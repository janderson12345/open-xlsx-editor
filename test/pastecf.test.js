import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { kids } from '../src/ooxml/xml.js';
import { snapshot, moveCutParts, pasteConditionals } from '../src/ooxml/clipops.js';
import { build, area } from './steps.js';

// Excel's conditional formats after a copy and PasteSpecial (or a cut) in a hidden Excel, read from the file it saved
// (docs/CLIPBOARD.md): each conditionalFormatting's cells and its rules' priorities and formulas. The sheet: a rule
// "cell value > 5" (priority 1) and "=A1>B1" (2) over A1:A3; in some, "cell value < 1" over E1:E4 (3).
const GT = { kind: 'cellIs', operator: 'greaterThan', f1: '5', preset: 'lightRed' };
const LT = { kind: 'cellIs', operator: 'lessThan', f1: '1', preset: 'green' };
const expr = (f) => ({ kind: 'expression', formula: f, preset: 'bold' });

/** A sheet whose rules are `rules` ([range, rule], the first given priority 1), as Excel's FormatConditions.Add made them. */
function sheetWith(rules) {
  const { wb, si } = build({ cells: { A1: '7', A2: '7', A3: '7', B1: '3', B2: '3', B3: '3', E1: '9', E2: '9', E3: '9', E4: '9' } });
  for (const [range, rule] of [...rules].reverse()) wb.setConditional(si, [area(range)], rule);
  return { wb, si };
}
/** The sheet's conditional formats as Excel's saved file reads: "sqref: p<priority> <formula>, ...", sorted. */
function blocks(wb, si) {
  const doc = wb._part(si).doc;
  const ns = doc.documentElement.namespaceURI;
  return kids(doc.documentElement, 'conditionalFormatting', ns).map((cf) => `${cf.getAttribute('sqref')}: ${kids(cf, 'cfRule', ns)
    .map((r) => [Number(r.getAttribute('priority')), kids(r, 'formula', ns).map((f) => f.textContent).join('|')])
    .sort((a, b) => a[0] - b[0]).map(([p, f]) => `p${p} ${f}`).join(', ')}`).sort();
}
const A13 = [['A1:A3', GT], ['A1:A3', expr('A1>B1')]];
const WITH_E = [...A13, ['E1:E4', LT]];

const CASES = [
  // [name, rules, copied, pasted at, options, Excel's]
  ['All at C1', A13, 'A1:A3', 'C1:C3', {}, ['A1:A3: p3 5, p4 A1>B1', 'C1:C3: p1 5, p2 C1>D1']],
  ['All over E1:E4\'s rule at E1', WITH_E, 'A1:A3', 'E1:E3', {}, ['A1:A3: p3 5, p4 A1>B1', 'E1:E3: p1 5, p2 E1>F1', 'E4: p5 1']],
  ['All merging conditional formats there', WITH_E, 'A1:A3', 'E1:E3', { merge: true }, ['A1:A3: p3 5, p4 A1>B1', 'E1:E3: p1 5, p2 E1>F1', 'E1:E4: p5 1']],
  ['All over part of it at E2', WITH_E, 'A1:A3', 'E2:E4', {}, ['A1:A3: p3 5, p4 A1>B1', 'E1: p5 1', 'E2:E4: p1 5, p2 E2>F2']],
  ['All over its own cells at A2', A13, 'A1:A3', 'A2:A4', {}, ['A1:A4: p1 5, p2 A1>B1']],
  ['All just under them at A4', A13, 'A1:A3', 'A4:A6', {}, ['A1:A6: p1 5, p2 A1>B1']],
  ['All beside them at B1', A13, 'A1:A3', 'B1:B3', {}, ['A1:A3: p3 5, p4 A1>B1', 'B1:B3: p1 5, p2 B1>C1']],
  ['All a row apart at A5', A13, 'A1:A3', 'A5:A7', {}, ['A1:A3: p3 5, p4 A1>B1', 'A5:A7: p1 5, p2 A5>B5']],
  ['All across the corner at B4', A13, 'A1:A3', 'B4:B6', {}, ['A1:A3: p3 5, p4 A1>B1', 'B4:B6: p1 5, p2 B4>C4']],
  ['A2:A3 at A1, inside', A13, 'A2:A3', 'A1:A2', {}, ['A1:A3: p1 5, p2 A1>B1']],
  ['A2 far off at H9', A13, 'A2', 'H9', {}, ['A1:A3: p3 5, p4 A1>B1', 'H9: p1 5, p2 H9>I9']],
  ['A2 just under at A4', A13, 'A2', 'A4', {}, ['A1:A4: p1 5, p2 A1>B1']],
  ['Transposed at C1', A13, 'A1:A3', 'C1:E1', { transpose: true }, ['A1:A3: p3 5, p4 A1>B1', 'C1:E1: p1 5, p2 C1>D1']],
  ['A row\'s rule at its right, D1', [['A1:C1', GT], ['A1:C1', expr('A1>B1')]], 'A1:C1', 'D1:F1', {}, ['A1:F1: p1 5, p2 A1>B1']],
  ['A row\'s rule over itself at B1', [['A1:C1', GT], ['A1:C1', expr('A1>B1')]], 'A1:C1', 'B1:D1', {}, ['A1:D1: p1 5, p2 A1>B1']],
  ['A row\'s rule under it at A2', [['A1:C1', GT], ['A1:C1', expr('A1>B1')]], 'A1:C1', 'A2:C2', {}, ['A1:C1: p3 5, p4 A1>B1', 'A2:C2: p1 5, p2 A2>B2']],
  ['A column\'s rule above it, A3:A4 at A1', [['A3:A5', GT], ['A3:A5', expr('A3>B3')]], 'A3:A4', 'A1:A2', {}, ['A1:A5: p1 5, p2 A1>B1']],
  ['A block\'s rule under it at A3', [['A1:B2', GT], ['A1:B2', expr('A1>B1')]], 'A1:B2', 'A3:B4', {}, ['A1:B2: p3 5, p4 A1>B1', 'A3:B4: p1 5, p2 A3>B3']],
  ['A block\'s rule beside it at C1', [['A1:B2', GT], ['A1:B2', expr('A1>B1')]], 'A1:B2', 'C1:D2', {}, ['A1:B2: p3 5, p4 A1>B1', 'C1:D2: p1 5, p2 C1>D1']],
];

for (const [name, rules, from, to, opts, excel] of CASES) {
  test(`conditional formats pasted: ${name}, as Excel's`, () => {
    const { wb, si } = sheetWith(rules);
    const copy = snapshot(wb, si, area(from));
    pasteConditionals(wb, si, area(to), copy, opts);
    assert.deepEqual(blocks(wb, si), [...excel].sort());
  });
}

test('conditional formats cut and pasted: moved, the formula as a cut formula\'s, as Excel\'s', () => {
  const { wb, si } = sheetWith(A13);
  const cut = snapshot(wb, si, area('A1:A3'), true);
  moveCutParts(wb, si, area('C1:C3'), cut);
  assert.deepEqual(blocks(wb, si), ['C1:C3: p1 5, p2 C1>B1']);
});

test('conditional formats pasted on another sheet: its own rules there, the sheet copied from as it was', () => {
  const { wb, si } = sheetWith(A13);
  const other = wb.addSheet(wb.sheets.length, 'Other');
  const copy = snapshot(wb, si, area('A1:A3'));
  pasteConditionals(wb, other, area('C1:C3'), copy);
  assert.deepEqual(blocks(wb, si), ['A1:A3: p1 5, p2 A1>B1']);
  assert.deepEqual(blocks(wb, other), ['C1:C3: p1 5, p2 C1>D1']);
  // (Their formats copied: each pasted rule a format of its own, as Excel's new dxfs.)
  const ids = (k) => kids(wb._part(k).doc.documentElement, 'conditionalFormatting', wb._part(k).doc.documentElement.namespaceURI).flatMap((cf) => [...cf.getElementsByTagNameNS('*', 'cfRule')].map((r) => r.getAttribute('dxfId')));
  assert.equal(new Set([...ids(si), ...ids(other)]).size, 4);
});

test('a rule left without its first cell by a paste, written for its new one, as Excel\'s', () => {
  // (Excel: =E1>F1 over E1:E4; A1:A2, without rules, pasted at E1: =E3>F3 over E3:E4.)
  const pasted = sheetWith([['E1:E4', expr('E1>F1')]]);
  pasteConditionals(pasted.wb, pasted.si, area('E1:E2'), snapshot(pasted.wb, pasted.si, area('A1:A2')));
  assert.deepEqual(blocks(pasted.wb, pasted.si), ['E3:E4: p1 E3>F3']);
});

/** The extLst's rules: "sqref: type p<priority or -> id" sorted, and each base data bar's id with its sqref. */
function extRules(wb, si) {
  const doc = wb._part(si).doc;
  const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const x14 = [...doc.getElementsByTagNameNS(X14, 'cfRule')].map((r) => ({ sq: r.parentNode.getElementsByTagNameNS('*', 'sqref')[0].textContent, type: r.getAttribute('type'), p: r.getAttribute('priority'), id: r.getAttribute('id') }));
  const base = [...doc.getElementsByTagNameNS(doc.documentElement.namespaceURI, 'cfRule')].filter((r) => r.getAttribute('type') === 'dataBar')
    .map((r) => ({ sq: r.parentNode.getAttribute('sqref'), id: r.getElementsByTagNameNS(X14, 'id')[0]?.textContent }));
  return { x14, base };
}

test('data bars and Excel 2010 icon sets pasted with their extLst half, new ids, as Excel\'s', () => {
  // (Excel: a data bar over A1:A3 copied to C1: C1:C3 in both halves, a new id; at A4, both become A1:A6; three stars
  // over A1:A3 copied to C1: a new x14 rule over C1:C3, priority 1, the old one 2.)
  const bars = sheetWith([['A1:A3', { kind: 'dataBar', color: 'FF638EC6' }]]);
  pasteConditionals(bars.wb, bars.si, area('C1:C3'), snapshot(bars.wb, bars.si, area('A1:A3')));
  const b = extRules(bars.wb, bars.si);
  assert.deepEqual(b.x14.map((x) => x.sq).sort(), ['A1:A3', 'C1:C3']);
  assert.deepEqual(b.base.map((x) => x.sq).sort(), ['A1:A3', 'C1:C3']);
  for (const x of b.base) assert.equal(b.x14.find((y) => y.id === x.id)?.sq, x.sq, 'a base rule and its extLst half name each other');
  assert.equal(new Set(b.x14.map((x) => x.id)).size, 2);
  const line = sheetWith([['A1:A3', { kind: 'dataBar', color: 'FF638EC6' }]]);
  pasteConditionals(line.wb, line.si, area('A4:A6'), snapshot(line.wb, line.si, area('A1:A3')));
  assert.deepEqual(extRules(line.wb, line.si).x14.map((x) => x.sq), ['A1:A6']);
  assert.deepEqual(extRules(line.wb, line.si).base.map((x) => x.sq), ['A1:A6']);
  const stars = sheetWith([['A1:A3', { kind: 'iconSet', set: '3Stars' }]]);
  pasteConditionals(stars.wb, stars.si, area('C1:C3'), snapshot(stars.wb, stars.si, area('A1:A3')));
  const s = extRules(stars.wb, stars.si).x14.map((x) => `${x.sq} ${x.type} p${x.p}`).sort();
  assert.deepEqual(s, ['A1:A3 iconSet p2', 'C1:C3 iconSet p1']);
});

test('a data bar cut moves both halves; a rule partly cut keeps its cells left and the moved ones get their own, as Excel\'s', () => {
  // (Excel: a data bar over A1:A3 cut to C1: C1:C3 in both halves. =A1>B1 over A1:A3, A1:A2 cut to C1: =A3>B3 over A3
  // and, after it in priority, =C1>B1 over C1:C2; A1:A2 cut to A2: =A2>B1 over A2:A3.)
  const bars = sheetWith([['A1:A3', { kind: 'dataBar', color: 'FF638EC6' }]]);
  moveCutParts(bars.wb, bars.si, area('C1:C3'), snapshot(bars.wb, bars.si, area('A1:A3'), true));
  assert.deepEqual(extRules(bars.wb, bars.si).x14.map((x) => x.sq), ['C1:C3']);
  assert.deepEqual(extRules(bars.wb, bars.si).base.map((x) => x.sq), ['C1:C3']);
  const part = sheetWith([['A1:A3', expr('A1>B1')]]);
  moveCutParts(part.wb, part.si, area('C1:C2'), snapshot(part.wb, part.si, area('A1:A2'), true));
  assert.deepEqual(blocks(part.wb, part.si), ['A3: p1 A3>B3', 'C1:C2: p2 C1>B1']);
  const down = sheetWith([['A1:A3', expr('A1>B1')]]);
  moveCutParts(down.wb, down.si, area('A2:A3'), snapshot(down.wb, down.si, area('A1:A2'), true));
  assert.deepEqual(blocks(down.wb, down.si), ['A2:A3: p1 A2>B1']);
});
