import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { formatLetters } from '../src/calc/cellinfo.js';

// (Excel 365, English, over COM: CELL("format"), and after a bar CELL("parentheses") and CELL("color") where either is 1,
// of a cell given each number format.)
const EXCEL = [
  ["General", "G"], ["0", "F0"], ["0.00", "F2"], ["#,##0", ",0"], ["#,##0.00", ",2"], ["#,##0_);(#,##0)", ",0"],
  ["#,##0_);[Red](#,##0)", ",0-|01"], ["#,##0.00_);(#,##0.00)", ",2"], ["#,##0.00_);[Red](#,##0.00)", ",2-|01"],
  ["$#,##0_);($#,##0)", "C0"], ["$#,##0_);[Red]($#,##0)", "C0-|01"], ["$#,##0.00_);($#,##0.00)", "C2"],
  ["$#,##0.00_);[Red]($#,##0.00)", "C2-|01"], ["0%", "P0"], ["0.00%", "P2"], ["0.00E+00", "S2"], ["##0.0E+0", "S1"],
  ["# ?/?", "G"], ["# ??/??", "G"], ["m/d/yyyy", "D4"], ["d-mmm-yy", "D1"], ["d-mmm", "D2"], ["mmm-yy", "D3"],
  ["h:mm AM/PM", "D7"], ["h:mm:ss AM/PM", "D6"], ["h:mm", "D9"], ["h:mm:ss", "D8"], ["m/d/yyyy h:mm", "D4"],
  ["mm:ss", "G"], ["[h]:mm:ss", "G"], ["mm:ss.0", "G"], ["@", "G"], ["0.000", "F3"], ["#,##0.000", ",3"],
  ["0.0%", "P1"], ["$#,##0", "C0"], ["$#,##0.00", "C2"], ["[Red]0", "F0"], ["0;(0)", "F0"], ["0;[Red]-0", "F0-|01"],
  ["#,##0;-#,##0", ",0"], ["0.0", "F1"], ["yyyy-mm-dd", "G"], ["dddd", "G"], ["mmmm d, yyyy", "D4"],
  ["[$-409]mmmm d, yyyy", "D2"], ["_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)", "C0"],
  ["_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)", "C2"],
  ["_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)", ",0"], ["_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", ",2"],
  ["\"x\"0", "F0"], ["0\" kg\"", "F0"], ["#", "F0"], ["?", "F0"], ["0.00;[Blue]-0.00", "F2-|01"],
  ["[Color10]0", "F0"], ["[>100]0;0.0", "F0"], ["0.00_)", "F2"], ["(0)", "F0()|10"], ["0;0;0;@", "F0"], ["yy", "G"],
  ["d", "G"], ["mmm d", "D5"], ["h", "G"], ["s", "G"], ["AM/PM h", "G"], ["m/d/yy", "D4"], ["mm/dd/yy", "D4"],
  ["mm/dd/yyyy", "D4"], ["d/m/yyyy", "D1"], ["dd/mm/yyyy", "D1"], ["yyyy/m/d", "G"], ["m/d", "D5"], ["mm/dd", "D5"],
  ["d/m", "D2"], ["m-d", "D5"], ["d-m-yy", "D1"], ["dd-mmm-yy", "D1"], ["dd-mmm-yyyy", "D1"], ["d-mmm-yyyy", "D1"],
  ["mmm-yyyy", "D3"], ["mmmm-yy", "D3"], ["mmm yy", "D3"], ["mmmm yyyy", "D3"], ["d mmm yyyy", "D1"],
  ["d mmmm yyyy", "D1"], ["mmm d, yyyy", "D4"], ["mmmm d", "D5"], ["d mmm", "D2"], ["dd mmm", "D2"], ["m/yy", "D3"],
  ["mm/yyyy", "D3"], ["yyyy-mm", "G"], ["yyyy/mm/dd", "G"], ["yy-mm-dd", "G"], ["dddd, mmmm d, yyyy", "D2"],
  ["ddd, mmm d", "D2"], ["[$-409]m/d/yy", "D2"], ["[$-409]d-mmm-yy", "G"], ["[$-409]mmm-yy", "D1"],
  ["[$-409]h:mm AM/PM", "G"], ["[$-F800]dddd, mmmm dd, yyyy", "G"], ["[$-F400]h:mm:ss AM/PM", "G"],
  ["h:mm:ss.000", "D8"], ["hh:mm", "D9"], ["hh:mm:ss", "D8"], ["h:mm AM/PM;@", "D7"], ["m/d/yyyy;@", "D4"],
  ["[h]:mm", "G"], ["[mm]:ss", "G"], ["h:mm:ss A/P", "D6"], ["h AM/PM", "G"], ["m/d/yy h:mm AM/PM", "D4"],
  ["yyyy-mm-dd hh:mm", "G"], ["d-mmm-yy h:mm", "D1"], ["mmm-yy;@", "D3"], ["mmmm", "G"], ["mmm", "G"], ["ddd", "G"],
  ["yyyy", "G"], ["m", "G"], ["mm", "G"], ["dd", "G"], ["[$-409]mmmm d", "D2"], ["[$-409]mmmm yyyy", "D1"],
  ["[$-409]d mmmm yyyy", "G"], ["[$-409]dddd", "G"], ["[$-409]yyyy-mm-dd", "G"],
];

test('CELL("format"), "parentheses" and "color" of 128 number formats, as Excel gives them', () => {
  for (const [code, want] of EXCEL) {
    const o = formatLetters(code);
    assert.equal(o.format + (o.parentheses || o.color ? `|${o.parentheses}${o.color}` : ''), want, code);
  }
});

test('CELL of another sheet names the workbook; prefix by alignment for typed text only; width and whether it is the default', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Other Sheet');
  wb.fileName = 'Book1';
  const at = (f, array = false) => wb.calc.formulaAt(0, 40, 0, f, array);
  wb.setCell(si, 'C3', 7);
  assert.equal(at(`CELL("address",'Other Sheet'!C3)`), "'[Book1]Other Sheet'!$C$3");
  wb.setCell(si, 'B2', 'text');
  wb.setFormula(si, 'B3', '"f"');
  const right = wb.styleWith(0, { align: { horizontal: 'right' } });
  wb.setStyles(si, [{ r: 1, c: 1, s: right }]);
  assert.equal(at(`CELL("prefix",'Other Sheet'!B2)`), '"');
  assert.equal(at(`CELL("prefix",'Other Sheet'!B3)`), '');
  assert.equal(at(`CELL("type",'Other Sheet'!B3)`), 'l');
  const w = at(`CELL("width",'Other Sheet'!A1)`, true);
  assert.deepEqual(w.rows ?? w, [[8, true]]);
});
