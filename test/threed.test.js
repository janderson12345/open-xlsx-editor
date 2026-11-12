import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
import { entryChanges } from '../src/ooxml/entry.js';
import { tokenize, renameSheetRefs } from '../src/calc/tokens.js';

// (3D references and intersection, as a hidden Excel 365 worked them out: probe/i3d.ps1's 86 formulas, 84 the same
// here; the formula probe has 17 more.)
function book() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const base = wb.sheets.length;
  ['Jan', 'Feb', 'My Mar', 'Apr'].forEach((n, k) => {
    wb.addSheet(base + k, n);
    wb.setCell(base + k, { r: 0, c: 0 }, (k + 1) * 100);
  });
  return { wb, base };
}
const value = (wb, si, ref) => { const { r, c } = parseRef(ref); return wb.sheet(si).cells.get(r)?.get(c)?.v; };

test('a 3D reference reads every sheet from its first to its last, and works out again when one of them changes', () => {
  const { wb, base } = book();
  wb.setFormula(base, 'B1', 'SUM(Jan:Apr!A1)');
  wb.setFormula(base, 'B2', "COUNT('Feb:My Mar'!A1)");
  // (Written either way round, as Excel reads it.)
  wb.setFormula(base, 'B3', 'SUM(Apr:Jan!A1)');
  assert.deepEqual([value(wb, base, 'B1'), value(wb, base, 'B2'), value(wb, base, 'B3')], [1000, 2, 1000]);
  wb.setCell(base + 2, 'A1', 5);
  assert.equal(value(wb, base, 'B1'), 705);
  // (On its own, a 3D reference is #REF!; one sheet that isn't there too.)
  wb.setFormula(base, 'B4', 'Jan:Apr!A1');
  wb.setFormula(base, 'B5', 'SUM(Jan:May!A1)');
  assert.deepEqual([value(wb, base, 'B4'), value(wb, base, 'B5')], ['#REF!', '#REF!']);
});

test('a sheet renamed at either end of a 3D reference is renamed in it, quoted when its name needs it', () => {
  assert.equal(renameSheetRefs('SUM(Jan:Apr!A1)', 'Jan', 'January'), 'SUM(January:Apr!A1)');
  assert.equal(renameSheetRefs('SUM(Jan:Apr!A1)', 'Apr', 'My Apr'), "SUM('Jan:My Apr'!A1)");
  assert.equal(renameSheetRefs("SUM('Feb:My Mar'!A1)", 'My Mar', 'Mar'), 'SUM(Feb:Mar!A1)');
  // (A sheet between the ends: nothing to change.)
  assert.equal(renameSheetRefs('SUM(Jan:Apr!A1)', 'Feb', 'X'), 'SUM(Jan:Apr!A1)');
  const { wb, base } = book();
  wb.setFormula(base, 'B1', 'SUM(Jan:Apr!A1)');
  wb.renameSheet(base + 3, 'April');
  assert.equal(wb.sheet(base).cells.get(0).get(1).f, 'SUM(Jan:April!A1)');
});

test('the intersection operator is a space between references, and only there', () => {
  const ops = (f) => tokenize(f).filter((t) => t.type === 'op').map((t) => t.text);
  assert.deepEqual(ops('A1:C3 B2:D4'), [' ']);
  assert.deepEqual(ops('SUM( A1:B2 )'), []);
  assert.deepEqual(ops('A1 + B1'), ['+']);
  assert.deepEqual(ops('INDEX(A1:C3,0,2) A2:C2'), [' ']);
  // (S1 reads as a cell: S1:S4!A1 is a range to S4!A1, as Excel reads it, not a 3D reference.)
  assert.ok(!tokenize('S1:S4!A1').some((t) => t.sheets));
  assert.deepEqual(tokenize('Jan:Apr!A1')[0].sheets, ['Jan', 'Apr']);
  const { wb, base } = book();
  wb.setFormula(base, 'E1', 'SUM(A1:C3 A1:A2)');
  wb.setFormula(base, 'E2', 'A1 C1');
  assert.deepEqual([value(wb, base, 'E1'), value(wb, base, 'E2'), wb.sheet(base).cells.get(0).get(4).f], [100, '#NULL!', 'SUM(A1:C3 A1:A2)']);
});

test('a 3D reference as sheets go and move, as Excel changes it: an end that leaves the span gives way inward', () => {
  const made = () => {
    const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
    const base = wb.sheets.length;
    ['Jan', 'Feb', 'Mar', 'Apr', 'Other'].forEach((n, k) => { wb.addSheet(base + k, n); wb.setCell(base + k, 'A1', (k + 1) * 100); });
    const o = wb.sheetIndexOf('Other');
    wb.setFormula(o, 'B1', 'SUM(Jan:Apr!A1)');
    wb.setFormula(o, 'B2', 'SUM(Feb:Mar!A1)');
    return wb;
  };
  const shown = (wb) => { const o = wb.sheetIndexOf('Other'); return [0, 1].map((r) => { const x = wb.sheet(o).cells.get(r).get(1); return `${x.f}=${x.v}`; }); };
  let wb = made();
  wb.deleteSheet(wb.sheetIndexOf('Jan'));
  assert.deepEqual(shown(wb), ['SUM(Feb:Apr!A1)=900', 'SUM(Feb:Mar!A1)=500']);
  wb = made();
  wb.deleteSheet(wb.sheetIndexOf('Feb'));
  assert.deepEqual(shown(wb), ['SUM(Jan:Apr!A1)=800', 'SUM(Mar!A1)=300']);
  // (Moved out past Apr: Feb:Apr. Moved after Feb, still among the others: as it was, now without Feb.)
  wb = made();
  wb.moveSheet(wb.sheetIndexOf('Jan'), wb.sheetIndexOf('Apr'));
  assert.deepEqual(shown(wb), ['SUM(Feb:Apr!A1)=900', 'SUM(Feb:Mar!A1)=500']);
  wb = made();
  wb.moveSheet(wb.sheetIndexOf('Jan'), wb.sheetIndexOf('Feb'));
  assert.deepEqual(shown(wb), ['SUM(Jan:Apr!A1)=800', 'SUM(Feb:Mar!A1)=600']);
  // (Typed: the workbook's order, and one sheet's a plain reference.)
  wb = made();
  const o = wb.sheetIndexOf('Other');
  for (const [ref, text] of [['C1', '=SUM(Apr:Jan!A1)'], ['C2', '=Jan:Jan!A1']]) {
    const { r, c } = parseRef(ref);
    wb.apply(o, [entryChanges(wb, o, r, c, text).content]);
  }
  assert.deepEqual([wb.sheet(o).cells.get(0).get(2).f, wb.sheet(o).cells.get(1).get(2).f], ['SUM(Jan:Apr!A1)', 'Jan!A1']);
});

test('rows inserted or deleted, or cells cut, on a sheet of a 3D reference leave it as it was, as in Excel', () => {
  // (Measured: =SUM(Jan:Apr!B2:B4) keeps its text whatever one sheet's rows, columns or cut cells do; it reads what is
  // there now. Each sheet's B1:B6 holds sheet*100 + row.)
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const base = wb.sheets.length;
  ['Jan', 'Feb', 'Mar', 'Apr', 'Other'].forEach((n, k) => { wb.addSheet(base + k, n); for (let r = 0; r < 6; r++) wb.setCell(base + k, { r, c: 1 }, (k + 1) * 100 + r + 1); });
  const o = wb.sheetIndexOf('Other');
  wb.setFormula(o, 'D1', 'SUM(Jan:Apr!B2:B4)');
  const d1 = () => { const x = wb.sheet(o).cells.get(0).get(3); return `${x.f}=${x.v}`; };
  assert.equal(d1(), 'SUM(Jan:Apr!B2:B4)=3036');
  wb.changeLines(wb.sheetIndexOf('Jan'), 'row', 0, 1);
  assert.equal(d1(), 'SUM(Jan:Apr!B2:B4)=3033');
  wb.changeLines(wb.sheetIndexOf('Feb'), 'col', 0, 1);
  // (Then a column on Feb: its B2:B4 is empty now, 609 less.)
  assert.equal(d1(), 'SUM(Jan:Apr!B2:B4)=2424');
});

test('the range operator between references: the range holding both, as Excel works it out (25 cases measured)', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'R');
  const other = wb.addSheet(wb.sheets.length, 'S');
  for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) wb.setCell(si, { r, c }, 100 + (r + 1) * 10 + c + 1);
  const at = (ref) => { const { r, c } = parseRef(ref); return wb.sheet(si).cells.get(r)?.get(c)?.v; };
  const type = (ref, text) => { const { r, c } = parseRef(ref); const ch = entryChanges(wb, si, r, c, text); if (ch.content) wb.apply(si, [ch.content]); return ch; };
  type('F1', '=SUM(A1:INDEX(A:A,3))');
  type('F2', '=SUM(OFFSET(A1,1,0):A4)');
  type('F3', '=SUM(INDEX(A:A,2):INDEX(C:C,3))');
  type('F4', '=SUM(A1:B2 B2:C3:D4)');
  type('F5', `=SUM(B2:${wb.sheets[other].name}!C3)`);
  assert.deepEqual(['F1', 'F2', 'F3', 'F4', 'F5'].map(at), [363, 393, 762, 122, '#VALUE!']);
  // (A number or text on one side isn't a formula, as Excel refuses it.)
  assert.ok(type('F6', '=SUM(A1:5)').error);
  // (Worked out again when a cell between the ends changes: B2 is in A1:INDEX(C:C,3).)
  type('F7', '=SUM(A1:INDEX(C:C,3))');
  const was = at('F7');
  wb.setCell(si, 'B2', 1000);
  assert.equal(at('F7'), was - 122 + 1000);
  // (Negation over a range spills, as =-A1:A2 does in Excel 365.)
  type('H1', '=-A1:INDEX(A:A,2)');
  assert.deepEqual([at('H1'), at('H2')], [-111, -121]);
});
