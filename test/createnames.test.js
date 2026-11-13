import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { labelName, namesFromSelection, createNames, guessSides } from '../src/ooxml/createnames.js';
import { findName, deleteName } from '../src/ooxml/names.js';

const B = String.fromCharCode(92);

// Each label as Excel's Range.CreateNames named it (Top row, in a hidden Excel; null: no name made). Text only: the
// numbers, TRUE typed as a value and a formula's number named nothing.
const EXCEL = [
  ['Qty', 'Qty'], ['Unit Price', 'Unit_Price'], ['1st', '_1st'], ['A1', 'A1_'], ['R1C1', '_R1C1'], ['R', 'R_'],
  ['my-name', 'my_name'], ['Total!', 'Total'], ['  padded  ', 'padded'], ['a.b', 'a.b'], ['Price ($)', 'Price'],
  ['Ünïcode', 'Ünïcode'], ['x?', 'x?'], [B, B], ['', null], ['L'.repeat(300), 'L'.repeat(255)], ['A B  C', 'A_B__C'],
  ['Sales%', 'Sales'], ['XFE1', 'XFE1'], ['c', 'c_'], ['_ok', '_ok'], ['two\nlines', 'two_lines'], [`${B}x`, null],
  ['.dot', '_.dot'], ['?q', '_?q'], ['a__', 'a__'], ['a-b-', 'a_b'], ['(x)', 'x'], ['-a', 'a'], [`a${B}b`, `a${B}b`],
  [`x${B}`, `x${B}`], ['RC', 'RC_'], ['RC1', '_RC1'], ['C3', '_C3'], ['r2', '_r2'], ['R1C', '_R1C'], ['XFD1', 'XFD1_'],
  ['AB12', 'AB12_'], ['rc', 'rc_'], ['Q1-Q2', 'Q1_Q2'], ['_', null], ['-', null], ['!!', null], ['a.', 'a.'], ['é1', 'é1'],
  ['2020', '_2020'], ['TRUE', 'TRUE_'], ['A1B', 'A1B'], ['12abc', '_12abc'], ['R1C1x', '_R1C1x'], ['a b!', 'a_b'],
  [' 1x', '_1x'], ['x-_', 'x__'], ['__a', '__a'], ['a!-!b', 'a___b'], ['Rx', 'Rx'], ['XFD1048577', 'XFD1048577'],
  ['A0', 'A0'], ['$A$1', 'A_1'], ['Tax Rate %', 'Tax_Rate'], ['R1C1_', '_R1C1_'], ['R2D2', '_R2D2'], ['C3PO', '_C3PO'],
  ['RCx', 'RCx'], ['rc2x', '_rc2x'], ['Rc', 'Rc_'], ['RCC', 'RCC'], ['CR1', 'CR1_'], ['RR1', 'RR1_'], ['FALSE', 'FALSE_'],
  ['False_', 'False_'], [`x${B}y${B}`, `x${B}y${B}`], [B + B, null], ['.', '_.'], ['?', '_?'], ['__', '__'], ['a_', 'a_'],
  ['R1', '_R1'], ['L1', 'L1_'], ['B1', 'B1_'], ['T1', 'T1_'],
];

test('labels made into names as Excel makes them', () => {
  const wrong = EXCEL.filter(([label, name]) => labelName(label) !== name).map(([label, name]) => `${label.slice(0, 12)}: Excel ${name?.slice(0, 12)}, here ${labelName(label)?.slice(0, 12)}`);
  assert.deepEqual(wrong, []);
  // (Excel makes its built-in names of these; left out here.)
  assert.equal(labelName('Print_Area'), null);
  assert.equal(labelName('Print_Titles'), null);
});

const book = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  for (const d of [...wb.definedNames]) deleteName(wb, { name: d.name, sheet: d.sheet ?? null });
  return wb;
};
const put = (wb, si, cells) => wb.apply(si, Object.entries(cells).map(([ref, value]) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return { r: Number(m[2]) - 1, c: m[1].charCodeAt(0) - 65, value };
}));
const names = (list) => Object.fromEntries(list.map((x) => [x.name, x.refersTo.replace(/^=[^!]*!/, '')]));

test('the cells each name refers to, as Excel\'s CreateNames gave them', () => {
  const wb = book();
  const si = 0;
  // (Excel's grid: Top and Left, the corner's label naming all the cells inside.)
  put(wb, si, { B2: 'Grid', C2: 'Jan', D2: 'Feb', B3: 'North', B4: 'South', B5: 'East', C3: 1, D3: 1, C4: 1, D4: 1, C5: 1, D5: 1 });
  assert.deepEqual(names(namesFromSelection(wb, si, { r0: 1, c0: 1, r1: 4, c1: 3 }, { top: true, left: true })),
    { Jan: '$C$3:$C$5', Feb: '$D$3:$D$5', North: '$C$3:$D$3', South: '$C$4:$D$4', East: '$C$5:$D$5', Grid: '$C$3:$D$5' });
  // (Bottom row: the cells above; Right column: those left of it; Left with one row: the rest of it.)
  put(wb, si, { G2: 1, G3: 2, G4: 'Bot', H4: 'Bot2', J7: 5, K7: 'Righty', B10: 'RowName', C10: 3, D10: 3, E10: 3 });
  assert.deepEqual(names(namesFromSelection(wb, si, { r0: 1, c0: 6, r1: 3, c1: 7 }, { bottom: true })), { Bot: '$G$2:$G$3', Bot2_: '$H$2:$H$3' });
  assert.deepEqual(names(namesFromSelection(wb, si, { r0: 6, c0: 9, r1: 6, c1: 10 }, { right: true })), { Righty: '$J$7' });
  assert.deepEqual(names(namesFromSelection(wb, si, { r0: 9, c0: 1, r1: 9, c1: 4 }, { left: true })), { RowName: '$C$10:$E$10' });
  // (All four: every label, corners too, names the one cell inside.)
  put(wb, si, { B13: 'tl', C13: 'T1', D13: 'tr', B14: 'L1', D14: 'R1', B15: 'bl', C15: 'B1', D15: 'br' });
  const four = names(namesFromSelection(wb, si, { r0: 12, c0: 1, r1: 14, c1: 3 }, { top: true, left: true, bottom: true, right: true }));
  assert.deepEqual(Object.keys(four).sort(), ['B1_', 'bl', 'br', 'L1_', '_R1', 'T1_', 'tl', 'tr'].sort());
  assert.ok(Object.values(four).every((x) => x === '$C$14'));
  // (A selection the sides leave no cells in names nothing; Excel's CreateNames refused it.)
  assert.equal(namesFromSelection(wb, si, { r0: 19, c0: 0, r1: 19, c1: 0 }, { top: true }), null);
  assert.equal(namesFromSelection(wb, si, { r0: 1, c0: 1, r1: 4, c1: 3 }, {}), null);
});

test('names made workbook-wide, one there already replaced, formulas using them worked out', () => {
  const wb = book();
  const si = 0;
  put(wb, si, { A30: 'qty', A31: 4, B30: 'Qty', B31: 6 });
  wb.apply(si, [{ r: 32, c: 0, formula: 'SUM(Qty)' }]);
  createNames(wb, namesFromSelection(wb, si, { r0: 29, c0: 0, r1: 30, c1: 0 }, { top: true }));
  assert.equal(wb.sheet(si).cells.get(32).get(0).v, 4);
  // (The second label of the same name replaces the first, its capitals too, as Excel's did.)
  createNames(wb, namesFromSelection(wb, si, { r0: 29, c0: 1, r1: 30, c1: 1 }, { top: true }));
  const d = findName(wb, 'qty', null);
  assert.equal(d.name, 'Qty');
  assert.match(d.text, /!\$B\$31$/);
  assert.equal(wb.definedNames.filter((x) => x.name.toLowerCase() === 'qty').length, 1);
  assert.equal(wb.sheet(si).cells.get(32).get(0).v, 6);
});

test('the sides the box starts with', () => {
  const wb = book();
  put(wb, 0, { B40: 'Grid', C40: 'Jan', B41: 'North', C41: 1 });
  assert.deepEqual(guessSides(wb, 0, { r0: 39, c0: 1, r1: 40, c1: 2 }), { top: true, left: true, bottom: false, right: false });
  put(wb, 0, { E40: 'Head', E41: 2, E42: 3 });
  assert.deepEqual(guessSides(wb, 0, { r0: 39, c0: 4, r1: 41, c1: 4 }), { top: true, left: false, bottom: false, right: false });
});
