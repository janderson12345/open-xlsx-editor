import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { legacyAt } from '../src/calc/legacyat.js';
import { editText } from '../src/ooxml/input.js';

// (Excel 365's Formula2 of formulas set by Range.Formula at H10, H12, ..., with A1:A3 1 to 3, B1:B3 t1 to t3, C1:C3
// 10 to 30: 58 measured, 57 the same here. The one not: =SEQUENCE(A1:A3), whose A1:A3 cut to H's row is an error
// here, so no array comes out to mark; Excel shows =@SEQUENCE(@A1:A3).)
const EXCEL = [
  ['A1:A3', '@A1:A3'], ['A1:A3*2', '@A1:A3*2'], ['SUM(A1:A3)*2', 'SUM(A1:A3)*2'], ['A1+A2', 'A1+A2'], ['-A1:A3', '-@A1:A3'],
  ['ABS(A1:A3)', 'ABS(@A1:A3)'], ['SUM(A1:A3*2)', 'SUM(@A1:A3*2)'], ['SUMPRODUCT(A1:A3*2)', 'SUMPRODUCT(A1:A3*2)'],
  ['INDEX(A1:A3,2)', 'INDEX(A1:A3,2)'], ['OFFSET(A1,0,0,2)', '@OFFSET(A1,0,0,2)'],
  ['VLOOKUP(A1:A3,A1:B3,2,FALSE)', 'VLOOKUP(@A1:A3,A1:B3,2,FALSE)'], ['MATCH(2,A1:A3,0)', 'MATCH(2,A1:A3,0)'],
  ['ROW(A1:A3)', '@ROW(A1:A3)'], ['ROWS(A1:A3)', 'ROWS(A1:A3)'], ['N(A1:A3)', 'N(A1:A3)'], ['CHOOSE(1,A1:A3)', '@CHOOSE(1,A1:A3)'],
  ['MAX(A1:A3)+A1:A3', 'MAX(A1:A3)+@A1:A3'], ['TRANSPOSE(A1:A3)', '@TRANSPOSE(@A1:A3)'], ['INDEX(A1:A3,0,1)', '@INDEX(A1:A3,0,1)'],
  ['IFERROR(A1:A3,0)', 'IFERROR(@A1:A3,0)'], ['COUNTIF(A1:A3,A1:A3)', 'COUNTIF(A1:A3,@A1:A3)'], ['SUMIF(A1:A3,">1")', 'SUMIF(A1:A3,">1")'],
  ['SUM(A1:A3+C1:C3)', 'SUM(@A1:A3+@C1:C3)'], ['XLOOKUP(A1:A3,A1:A3,C1:C3)', 'XLOOKUP(@A1:A3,A1:A3,C1:C3)'],
  ['FILTER(A1:A3,A1:A3>1)', '@FILTER(A1:A3,A1:A3>1)'], ['SUM(OFFSET(A1,0,0,3))', 'SUM(OFFSET(A1,0,0,3))'],
];

test("Excel 365's @ in a formula of the old kind, where it cuts a range or an array to one value (measured)", () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'At');
  for (let r = 0; r < 3; r++) { wb.setCell(si, { r, c: 0 }, r + 1); wb.setCell(si, { r, c: 1 }, `t${r + 1}`); wb.setCell(si, { r, c: 2 }, (r + 1) * 10); }
  EXCEL.forEach(([f, want], i) => assert.equal(legacyAt(wb.calc, si, 9 + 2 * i, 7, f), want, f));
  // (The formula bar: an old formula with them, a dynamic one as it is.)
  const at = (shown) => legacyAt(wb.calc, si, 9, 7, shown);
  assert.equal(editText({ f: 'A1:A3*2' }, 'General', false, at), '=@A1:A3*2');
  assert.equal(editText({ f: 'A1:A3*2', dynamic: true }, 'General', false, at), '=A1:A3*2');
  assert.equal(editText({ f: 'A1:A3' }), '=@A1:A3');
});
