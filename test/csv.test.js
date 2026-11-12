import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parseCsv, csvWorkbook, sheetNameOf, isCsv } from '../src/ooxml/csv.js';
import { formatCodeOf } from '../src/ooxml/styles.js';

// (test/csv-excel.csv, made-up fields, opened by Excel 365 over COM: each cell's value, number format and formula.
// Mar-24 is the 24th of March this year, as Excel took it in 2026.)
const EXCEL = {
  B2: [42, 'General'], C2: ['quoted, comma', 'General'], B3: [7, 'General'], C3: ['line\r\nbreak', 'General'],
  B4: [3.14, 'General'], C4: ['say "hi"', 'General'], B5: [100000, '0.00E+00'], C5: [true, 'General'],
  B6: [0.5, '0%'], C6: [false, 'General'], B7: [1234.5, '$#,##0.00_);[Red]($#,##0.00)'], C7: ['=1+2', 'General'],
  B8: [45356, 'm/d/yyyy'], C8: [45356, 'm/d/yyyy'], B9: [0.5729166666666666, 'h:mm'], C9: [46024, 'd-mmm'],
  B10: [-12, 'General'], C10: [-12, 'General'], B11: [12345678901234567000, 'General'], C11: [1.5, 'General'],
  B12: [7, 'General'], C12: ['  padded  ', 'General'], B13: [null, 'General'], C13: [null, 'General'],
  B14: [1234, '#,##0'], C14: ['1 234', 'General'], B15: [46024, 'd-mmm'], C15: [46105, 'd-mmm'],
  B16: ['#N/A', 'General'], C16: ['#DIV/0!', 'General'],
};

test('a CSV opened as Excel opens it: each field as if typed, quotes holding commas, line breaks and quotes (Excel\'s cells, measured)', () => {
  const wb = csvWorkbook(new Uint8Array(fs.readFileSync('test/csv-excel.csv')), 'csvprobe.csv', new Date(2026, 0, 15));
  assert.deepEqual(wb.sheets.map((s) => s.name), ['csvprobe']);
  const sh = wb.sheet(0);
  for (const [ref, [v, code]] of Object.entries(EXCEL)) {
    const r = Number(ref.slice(1)) - 1;
    const c = ref.charCodeAt(0) - 65;
    const cell = sh.cells.get(r)?.get(c);
    const got = cell == null ? null : cell.f != null ? `=${cell.f}` : cell.v;
    if (typeof v === 'number') assert.ok(Math.abs(got - v) < 1e-9 * Math.max(1, Math.abs(v)), `${ref}: ${got} for ${v}`);
    else assert.deepEqual(got, v, ref);
    // (Excel shows a built-in format without its quotes and backslashes: "$"#,##0.00 as $#,##0.00.)
    assert.equal(formatCodeOf(wb.styles, cell?.s || 0).replace(/["\\]/g, ''), code, `${ref}'s format`);
  }
  assert.equal(sh.cells.get(0).get(0).v, 'name');
  assert.equal(wb.edited, false);
});

test('CSV fields and names: quotes, empty fields, line ends; the sheet named after the file; a CSV told from a workbook', () => {
  assert.deepEqual(parseCsv('a,"b,c",\r\n"d""e",f\n\ng'), [['a', 'b,c', ''], ['d"e', 'f'], [''], ['g']]);
  assert.deepEqual(parseCsv('x,y\r\n'), [['x', 'y']]);
  assert.equal(sheetNameOf('C:/data/Sales [2026].csv'), 'Sales _2026_');
  assert.equal(sheetNameOf('a-very-long-file-name-for-a-sheet-tab.csv'), 'a-very-long-file-name-for-a-she');
  assert.equal(isCsv(new Uint8Array([0x61, 0x2c]), 'a.csv'), true);
  assert.equal(isCsv(new Uint8Array([0x50, 0x4b]), 'a.csv'), false);
  assert.equal(isCsv(new Uint8Array([0x61, 0x2c]), 'a.xlsx'), false);
});

test("a sheet saved as a CSV as Excel's Save As writes one, byte for byte: CSV UTF-8 and CSV (code page 1252)", async () => {
  const { csvBytes } = await import('../src/ooxml/csv.js');
  const { Workbook } = await import('../src/ooxml/workbook.js');
  // (The sheet Excel saved as test/csv-excel-utf8.csv and test/csv-excel-ansi.csv, made the same way here.)
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
  assert.deepEqual([...csvBytes(wb, si)], [...fs.readFileSync('test/csv-excel-utf8.csv')]);
  assert.deepEqual([...csvBytes(wb, si, { utf8: false })], [...fs.readFileSync('test/csv-excel-ansi.csv')]);
});
