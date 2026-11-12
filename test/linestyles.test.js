import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseXml } from '../src/ooxml/xml.js';
import { lineStyle } from '../src/ooxml/sheet.js';
import { protectSheet, isLocked } from '../src/ooxml/protection.js';

// (The sheet Excel 365 saved after Columns(2).Font.Bold, Range("D:E").Interior.Color and Rows(5).Font.Italic over
// COM: <col style>, <row s customFormat="1">, and a cell made where a formatted row crosses a formatted column.)
function book() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Lines');
  const bold = wb.styleWith(0, { font: { bold: true } });
  const fill = wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } });
  const italic = wb.styleWith(0, { font: { italic: true } });
  const open = wb.styleWith(0, { protection: { locked: false } });
  const { doc } = wb._part(si);
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const cols = parseXml(`<cols xmlns="${ns}"><col min="2" max="2" width="9.140625" style="${bold}"/><col min="4" max="5" width="9.140625" style="${fill}"/><col min="7" max="7" width="9.140625" style="${open}"/></cols>`).documentElement;
  root.insertBefore(doc.importNode(cols, true), [...root.childNodes].find((n) => n.localName === 'sheetData'));
  const data = [...root.childNodes].find((n) => n.localName === 'sheetData');
  const row = parseXml(`<row xmlns="${ns}" r="5" s="${italic}" customFormat="1"><c r="F5" s="${italic}" t="inlineStr"><is><t>f</t></is></c></row>`).documentElement;
  data.appendChild(doc.importNode(row, true));
  wb.dirty.add(wb.sheets[si].path);
  const back = Workbook.fromBytes(wb.toBytes());
  return { wb: back, si, bold, fill, italic, open };
}

test("a row's and a column's own format, read as Excel writes them: an empty cell has its row's, else its column's", () => {
  const { wb, si, bold, fill, italic } = book();
  const sh = wb.sheet(si);
  assert.equal(sh.rows.get(4).s, italic);
  assert.equal(lineStyle(sh, 0, 1), bold); // B1
  assert.equal(lineStyle(sh, 99, 3), fill); // D100
  assert.equal(lineStyle(sh, 99, 4), fill); // E100
  assert.equal(lineStyle(sh, 4, 0), italic); // A5
  assert.equal(lineStyle(sh, 0, 0), 0); // A1
  assert.equal(lineStyle(sh, 0, 5), 0); // F1
});

test("a cell typed into a formatted column or row gets that format, as Excel's does; Format Cells > Protection on a column counts", () => {
  const { wb, si, bold, fill, italic } = book();
  wb.setCell(si, 'B9', 'b');
  wb.setCell(si, 'E2', 'e');
  wb.setCell(si, 'A5', 'a');
  wb.setCell(si, 'C1', 'c');
  const s = (ref) => { const r = Number(ref.slice(1)) - 1; const c = ref.charCodeAt(0) - 65; return wb.sheet(si).cells.get(r)?.get(c)?.s || 0; };
  assert.deepEqual([s('B9'), s('E2'), s('A5'), s('C1')], [bold, fill, italic, 0]);
  protectSheet(wb, si);
  assert.equal(isLocked(wb, si, 0, 6), false); // G1, its column unlocked
  assert.equal(isLocked(wb, si, 0, 7), true); // H1
});
