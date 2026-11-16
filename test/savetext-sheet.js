// The made-up sheet Excel saved as test/txt-excel-ansi.txt (Text, Tab delimited) and test/txt-excel-unicode.txt
// (Unicode Text) over COM (tools/excel-compare/excel-saveas.ps1), made the same way in test/csv.test.js.
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';

const TAB = String.fromCharCode(9);

/** A workbook (basic.xlsx with a sheet "Save" added) and the new sheet's index. */
export function saveSheet() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Save');
  const put = (ref, v, numFmt) => {
    if (typeof v === 'string' && v.startsWith('=')) wb.setFormula(si, ref, v.slice(1)); else wb.setCell(si, ref, v);
    if (numFmt) { const r = Number(ref.slice(1)) - 1; const c = ref.charCodeAt(0) - 65; wb.setStyles(si, [{ r, c, s: wb.styleWith(0, { numFmt }) }]); }
  };
  put('A1', 'plain'); put('B1', 'has, comma'); put('C1', 'say "hi"'); put('D1', 'two\nlines');
  put('A2', 1234.5, '#,##0.00'); put('B2', 0.25, '0%'); put('C2', 45356, 'm/d/yyyy'); put('D2', '=1/3');
  put('A3', true); put('B3', '=NA()'); put('D3', '  lead');
  put('A4', 1234567890123); put('B4', 12345678901234568); put('C4', 0.000001); put('D4', '007');
  put('B5', 'x');
  put('A6', 'café ü €'); put('C6', 12345); put('D6', -5, '0.00_);(0.00)');
  put('A7', 'semi;colon'); put('B7', 'quote"mid');
  put('A8', `tab${TAB}in`); put('B8', String.fromCharCode(0x65e5, 0x672c)); put('C8', 'a "q" b'); put('D8', '"lead');
  return { wb, si };
}
