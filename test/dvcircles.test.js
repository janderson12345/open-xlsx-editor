import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readSheet } from '../src/ooxml/sheet.js';
import { invalidCells, breaksRule } from '../src/ooxml/validation.js';
import { setValidation } from '../src/ooxml/dvedit.js';
import { circleBoxes } from '../src/editor/dvcircles.js';
import { Axis } from '../src/editor/axis.js';
import { build, area } from './steps.js';
import { refName } from '../src/ooxml/refs.js';

/** A sheet built from `cells` with the rules `rules` ([range, rule]) written as Excel's box writes them. */
const sheetWith = (cells, rules) => {
  const { wb, si } = build({ cells });
  const { doc } = wb._part(si);
  for (const [range, rule] of rules) setValidation(doc, [area(range)], rule);
  wb._parsed.set(wb.sheets[si], { doc, model: readSheet(doc, wb.sharedStrings) });
  wb.calculate();
  return { wb, si };
};

test('the cells Circle Invalid Data circles, as Excel\'s CircleInvalid did', () => {
  // (Excel: whole numbers 1 to 10 over A1:A6 ignoring blanks and C1:C6 not; a list red,green,blue over E1:E3;
  // Information and Warning rules broken in G1 and G2; 99 in I2. It made 14 ovals: A2 50, A3 text, A5 2.5, A6 =40+2,
  // C2 50 and C3:C6 blank, E2 pink, E3 RED (a typed list matches case), G1, G2, I2.)
  const { wb, si } = sheetWith(
    { A1: '5', A2: '50', A3: 'abc', A5: '2.5', A6: '=40+2', C1: '5', C2: '50', E1: 'red', E2: 'pink', E3: 'RED', G1: '99', G2: '99', I2: '99' },
    [['A1:A6', { type: 'whole', op: 'between', f1: '1', f2: '10', ignoreBlank: true }],
      ['C1:C6', { type: 'whole', op: 'between', f1: '1', f2: '10', ignoreBlank: false }],
      ['E1:E3', { type: 'list', f1: 'red,green,blue' }],
      ['G1', { type: 'whole', op: 'between', f1: '1', f2: '10', error: { show: true, style: 'information' } }],
      ['G2', { type: 'whole', op: 'between', f1: '1', f2: '10', error: { show: true, style: 'warning' } }],
      ['I1:I3', { type: 'whole', op: 'between', f1: '1', f2: '10' }]]);
  const got = invalidCells(wb, si).map(({ r, c }) => refName(r, c)).sort();
  assert.deepEqual(got, ['A2', 'A3', 'A5', 'A6', 'C2', 'C3', 'C4', 'C5', 'C6', 'E2', 'E3', 'G1', 'G2', 'I2']);
  assert.equal(breaksRule(wb, si, 0, 0), false);
});

test('at most 255 circles, as Excel drew of 300 broken cells', () => {
  const cells = {};
  for (let r = 1; r <= 300; r++) cells[`A${r}`] = '99';
  const { wb, si } = sheetWith(cells, [['A1:A300', { type: 'whole', op: 'between', f1: '1', f2: '10' }]]);
  assert.equal(invalidCells(wb, si).length, 255);
});

test('each oval where Excel put its own, in pixels', () => {
  // (Excel, in points: B2 48,15 48x15 had its oval at 44.25,13.5 55.5x18; D2 135 wide 142.5; B4 39.75 tall 42.75;
  // A1 cut to 0,0 51.75x16.5; F6:G7 merged, 96x30, an oval 103.5x18, as tall as its first row.)
  const pt = (x) => x * 4 / 3;
  const cols = new Axis(10, 64, new Map([[3, 180]]));
  const rows = new Axis(10, 20, new Map([[3, 53]]));
  const boxes = circleBoxes([{ r: 1, c: 1 }, { r: 1, c: 3 }, { r: 3, c: 1 }, { r: 0, c: 0 }, { r: 5, c: 5 }], rows, cols, [{ r0: 5, c0: 5, r1: 6, c1: 6 }]);
  const near = (a, b) => Object.keys(b).every((k) => Math.abs(a[k] - b[k]) < 0.01);
  assert.ok(near(boxes[0], { x: pt(44.25), y: pt(13.5), w: pt(55.5), h: pt(18) }), JSON.stringify(boxes[0]));
  assert.ok(near(boxes[1], { x: pt(140.25), w: pt(142.5) }), JSON.stringify(boxes[1]));
  assert.ok(near(boxes[2], { y: pt(43.5), h: pt(42.75) }), JSON.stringify(boxes[2]));
  assert.ok(near(boxes[3], { x: 0, y: 0, w: pt(51.75), h: pt(16.5) }), JSON.stringify(boxes[3]));
  assert.ok(near(boxes[4], { w: pt(103.5), h: pt(18) }), JSON.stringify(boxes[4]));
});
