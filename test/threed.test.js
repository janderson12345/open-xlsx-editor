import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
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
  wb.setFormula(base, 'B1', 'SUM(A1:C3 A1:A2)');
  wb.setFormula(base, 'B2', 'A1 C1');
  assert.deepEqual([value(wb, base, 'B1'), value(wb, base, 'B2'), wb.sheet(base).cells.get(0).get(1).f], [100, '#NULL!', 'SUM(A1:C3 A1:A2)']);
});
