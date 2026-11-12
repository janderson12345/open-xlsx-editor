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
  const bold = wb.styleWith(0, { font: { b: true } });
  const fill = wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } });
  const italic = wb.styleWith(0, { font: { i: true } });
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

test("formatting whole rows and columns as Excel does: the lines' own formats, the cells there, a cell where they cross (Excel's sheet)", async () => {
  const fmt = await import('../src/ooxml/formatops.js');
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Fmt');
  wb.setCell(si, 'B1', 1); wb.setCell(si, 'B3', 3); wb.setCell(si, 'D2', 4); wb.setCell(si, 'F5', 'f');
  const MAXR = 1048575;
  const MAXC = 16383;
  // (What commands.js formatting does with a selection: the probe's formats the lines', the rest the cells'.)
  const format = (sel, change) => {
    const probe = fmt.probeOf(wb.sheet(si), sel);
    const out = fmt.formatEach(wb, si, { ...sel, ar: sel.r0, ac: sel.c0 }, change);
    const cells = out.filter((x) => x.r !== probe.row && x.c !== probe.col);
    wb.setStyles(si, cells);
    wb.setLineStyles(si, { cols: out.filter((x) => x.r === probe.row).map(({ c, s }) => ({ c, s })), rows: out.filter((x) => x.c === probe.col).map(({ r, s }) => ({ r, s })) });
  };
  format({ r0: 4, c0: 0, r1: 4, c1: MAXC }, { font: { i: true } });
  format({ r0: 0, c0: 1, r1: MAXR, c1: 1 }, { font: { b: true } });
  format({ r0: 0, c0: 3, r1: MAXR, c1: 4 }, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } });
  const xml = serializeXml(wb._part(si).doc);
  const sh = wb.sheet(si);
  const s = (ref) => sh.cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65)?.s ?? null;
  const look = (x) => { const xf = wb.styles.cellXfs[x]; const f = wb.styles.fonts[xf.fontId]; return `${f.b ? 'b' : ''}${f.i ? 'i' : ''}${xf.fillId ? 'f' : ''}`; };
  assert.match(xml, /<col min="2" max="2" width="[\d.]+" style="\d+"\/><col min="4" max="5" width="[\d.]+" style="\d+"\/>/);
  assert.match(xml, /<row r="5"[^>]* s="\d+" customFormat="1"/);
  assert.equal(look(sh.cols.find((x) => x.min === 1).style), 'b');
  assert.equal(look(sh.cols.find((x) => x.min === 3).style), 'f');
  assert.equal(look(sh.rows.get(4).s), 'i');
  // (Excel's cells: B1 and B3 bold, D2 filled, F5 italic, and B5, D5, E5 made where row 5 crosses the columns.)
  assert.deepEqual(['B1', 'B3', 'D2', 'F5', 'B5', 'D5', 'E5'].map((ref) => look(s(ref))), ['b', 'b', 'f', 'i', 'bi', 'if', 'if']);
  assert.equal(s('B2'), null);
  assert.equal(lineStyle(sh, 99, 1), sh.cols.find((x) => x.min === 1).style);
  wb.setCell(si, 'E40', 'e');
  assert.equal(look(s('E40')), 'f');
});

test("the whole sheet formatted as Excel does: every column its own format, split where one had a width or a format; the formatted rows and the cells there changed; no cell made", async () => {
  const fmt = await import('../src/ooxml/formatops.js');
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'All');
  wb.setCell(si, 'B1', 1); wb.setCell(si, 'D3', 3);
  wb.setLines(si, 'col', 2, 2, { width: 20.7109375 });
  wb.setLines(si, 'row', 5, 5, { height: 30 });
  const format = (sel, change) => {
    const probe = fmt.probeOf(wb.sheet(si), sel);
    const out = fmt.formatEach(wb, si, { ...sel, ar: sel.r0, ac: sel.c0 }, change);
    wb.setStyles(si, out.filter((x) => x.r !== probe.row && x.c !== probe.col));
    wb.setLineStyles(si, { cols: out.filter((x) => x.r === probe.row).map(({ c, s }) => ({ c, s })), rows: out.filter((x) => x.c === probe.col && x.r !== probe.row).map(({ r, s }) => ({ r, s })) });
  };
  format({ r0: 7, c0: 0, r1: 7, c1: 16383 }, { font: { i: true } });
  format({ r0: 0, c0: 4, r1: 1048575, c1: 4 }, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } });
  const cellsBefore = [...wb.sheet(si).cells.values()].reduce((n, l) => n + l.size, 0);
  format({ r0: 0, c0: 0, r1: 1048575, c1: 16383 }, { font: { b: true } });
  const xml = serializeXml(wb._part(si).doc);
  const sh = wb.sheet(si);
  const look = (x) => { const xf = wb.styles.cellXfs[x]; const f = wb.styles.fonts[xf.fontId]; return `${f.b ? 'b' : ''}${f.i ? 'i' : ''}${xf.fillId ? 'f' : ''}`; };
  // (Excel's: <col min="1" max="2" style="1"/><col min="3" max="3" style="1" customWidth="1"/><col min="4" max="4" style="1"/><col min="5" max="5" style="2"/><col min="6" max="16384" style="1"/>.)
  const cols = [...xml.matchAll(/<col min="(\d+)" max="(\d+)"[^>]*?style="(\d+)"/g)].map((m) => `${m[1]}-${m[2]}:${look(Number(m[3]))}`);
  assert.deepEqual(cols, ['1-2:b', '3-3:b', '4-4:b', '5-5:bf', '6-16384:b']);
  assert.match(xml, /<col min="3" max="3" width="20.7109375" customWidth="1" style="\d+"\/>/);
  assert.equal(look(sh.rows.get(7).s), 'bi');
  assert.equal(sh.rows.get(5).s, undefined);
  const s = (ref) => sh.cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65)?.s;
  assert.deepEqual(['B1', 'D3', 'E8'].map((ref) => look(s(ref))), ['b', 'b', 'bif']);
  assert.equal([...sh.cells.values()].reduce((n, l) => n + l.size, 0), cellsBefore);
  assert.equal(look(lineStyle(sh, 99, 25)), 'b');
});
