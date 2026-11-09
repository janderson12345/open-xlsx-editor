import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
import { snapshot, pasteCopy, pasteCut, pasteText, pasteArea } from '../src/ooxml/clipops.js';
import { moveRefs } from '../src/calc/tokens.js';
import { formatEach } from '../src/ooxml/formatops.js';

// What Excel did with the same cells over COM (Range.Copy or Cut, then Worksheet.Paste; docs/CLIPBOARD.md).
const range = (text) => {
  const [a, b = a] = text.split(':').map(parseRef);
  return { r0: a.r, c0: a.c, r1: b.r, c1: b.c, ar: a.r, ac: a.c };
};
/** Excel's probe sheet, on two new sheets of formulas.xlsx: A1:A3 1, 2, 3; B2 =A2*2 (bold); B3 =$A$1+A3; F1 =A1+A2;
 *  F2 =SUM(A1:A3); F3 =SUM(A1:A2); the second sheet's A1 reading the first's A1. */
function setup() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  const s2 = wb.addSheet(wb.sheets.length);
  wb.apply(si, [{ r: 0, c: 0, value: 1 }, { r: 1, c: 0, value: 2 }, { r: 2, c: 0, value: 3 }, { r: 1, c: 1, formula: 'A2*2' }, { r: 2, c: 1, formula: '$A$1+A3' },
    { r: 0, c: 5, formula: 'A1+A2' }, { r: 1, c: 5, formula: 'SUM(A1:A3)' }, { r: 2, c: 5, formula: 'SUM(A1:A2)' }]);
  wb.apply(s2, [{ r: 0, c: 0, formula: `${wb.sheets[si].name}!A1` }]);
  wb.setStyles(si, formatEach(wb, si, range('B2'), { font: { b: true } }));
  return { wb, si, s2, name: wb.sheets[si].name };
}
const run = (wb, steps) => { for (const st of steps) { if (st.contents.length) wb.apply(st.si, st.contents); wb.setMerges(st.si, st.merges); if (st.styles.length) wb.setStyles(st.si, st.styles); } };
const f = (wb, si, ref) => { const { r, c } = parseRef(ref); const cell = wb.sheet(si).cells.get(r)?.get(c); return cell?.f != null ? `=${cell.f}` : String(cell?.v ?? ''); };
const bold = (wb, si, ref) => { const { r, c } = parseRef(ref); const s = wb.sheet(si).cells.get(r)?.get(c)?.s || 0; return wb.styles.fonts[wb.styles.cellXfs[s].fontId].b; };

test('copy: relative references follow, $ ones stay, the format comes along', () => {
  const { wb, si } = setup();
  run(wb, pasteCopy(wb, si, range('D5'), snapshot(wb, si, range('B2'))));
  assert.deepEqual([f(wb, si, 'D5'), bold(wb, si, 'D5')], ['=C5*2', true]);
  run(wb, pasteCopy(wb, si, range('A1'), snapshot(wb, si, range('B2'))));
  assert.equal(f(wb, si, 'A1'), '=#REF!*2');
});

test('a paste repeats the copy over a selection that is a whole number of copies, else pastes it once', () => {
  assert.deepEqual(pasteArea(range('D5:D8'), 2, 1), { r0: 4, c0: 3, r1: 7, c1: 3 });
  assert.deepEqual(pasteArea(range('D5:D7'), 2, 1), { r0: 4, c0: 3, r1: 5, c1: 3 });
  const { wb, si } = setup();
  run(wb, pasteCopy(wb, si, range('D5:D8'), snapshot(wb, si, range('B2:B3'))));
  assert.deepEqual(['D5', 'D6', 'D7', 'D8', 'D9'].map((r) => f(wb, si, r)), ['=C5*2', '=$A$1+C6', '=C7*2', '=$A$1+C8', '']);
  run(wb, pasteCopy(wb, si, range('H5:I6'), snapshot(wb, si, range('B2'))));
  assert.deepEqual(['H5', 'I5', 'H6', 'I6'].map((r) => f(wb, si, r)), ['=G5*2', '=H5*2', '=G6*2', '=H6*2']);
});

test('cut: the cells move, and every formula that reads them follows', () => {
  let { wb, si, s2, name } = setup();
  run(wb, pasteCut(wb, si, range('C1'), snapshot(wb, si, range('A1:A2'), true)));
  assert.deepEqual(['C1', 'C2', 'A1', 'F1', 'F2', 'F3', 'B2'].map((r) => f(wb, si, r)), ['1', '2', '', '=C1+C2', '=SUM(A1:A3)', '=SUM(C1:C2)', '=C2*2']);
  assert.equal(f(wb, s2, 'A1'), `=${name}!C1`);

  ({ wb, si } = setup());
  run(wb, pasteCut(wb, si, range('C5'), snapshot(wb, si, range('A2:B2'), true)));
  assert.deepEqual(['C5', 'D5', 'F1'].map((r) => f(wb, si, r)), ['2', '=C5*2', '=A1+C5']);
  assert.deepEqual([bold(wb, si, 'D5'), bold(wb, si, 'B2')], [true, false]);

  ({ wb, si } = setup());
  run(wb, pasteCut(wb, si, range('A3'), snapshot(wb, si, range('A1'), true))); // (over A3, which B3 reads)
  assert.deepEqual(['A3', 'F1', 'B3'].map((r) => f(wb, si, r)), ['1', '=A3+A2', '=$A$3+#REF!']);
});

test('references when cells move to another sheet', () => {
  const from = { sheet: 'Data', r0: 0, c0: 0, r1: 1, c1: 0 };
  const to = { sheet: 'Other', dr: 0, dc: 2 };
  assert.equal(moveRefs('A1+B1', { here: 'Data', from, to }), 'Other!C1+B1');
  assert.equal(moveRefs('A1*B1', { here: 'Data', home: 'Other', from, to }), 'C1*Data!B1');
  assert.equal(moveRefs("'Data'!A2", { here: 'X', from, to }), 'Other!C2');
});

test('text from another program, read cell by cell as if typed', () => {
  const { wb, si } = setup();
  run(wb, pasteText(wb, si, range('J1'), '12\tabc\r\n50%\t=J1*2\r\n'));
  assert.deepEqual(['J1', 'K1', 'J2', 'K2'].map((r) => f(wb, si, r)), ['12', 'abc', '0.5', '=J1*2']);
});

test('merged ranges come with a copy', () => {
  const { wb, si } = setup();
  wb.setMerges(si, [{ r0: 0, c0: 7, r1: 0, c1: 8 }]);
  run(wb, pasteCopy(wb, si, range('H10'), snapshot(wb, si, range('H1:I1'))));
  assert.ok(wb.sheet(si).merges.some((m) => m.r0 === 9 && m.c0 === 7 && m.c1 === 8));
});
