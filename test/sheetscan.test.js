import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strToU8, strFromU8, unzipSync } from 'fflate';
import { readSheet, readSheetBytes } from '../src/ooxml/sheet.js';
import { readSharedStrings } from '../src/ooxml/strings.js';
import { SheetCells } from '../src/ooxml/sheetcells.js';
import { parseXml } from '../src/ooxml/xml.js';
import { Workbook } from '../src/ooxml/workbook.js';
import { filledRows, cellCount } from '../src/ooxml/sheetcells.js';
import { sheetProtection } from '../src/ooxml/protection.js';
import { readHyperlinks } from '../src/ooxml/links.js';
import { usedRange } from '../src/ooxml/special.js';
import { autoRowHeights, rowDescents } from '../src/editor/rowheights.js';

// A sheet's cells read from the part's bytes (sheetscan.js) against the same part parsed into a tree and walked, as
// every sheet was read before: the two models must be the same, cell for cell.

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sheet = (rows, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData>${rows}</sheetData>${extra}</worksheet>`;

/** A model with its cells as plain Maps (bands made into rows), to compare. */
const plain = (model) => ({ ...model, cells: new Map(model.cells) });

/** The part's text read both ways, with bands and without; the models compared. Returns the banded one. */
function same(text, sst = []) {
  const u8 = strToU8(text);
  const old = readSheet(parseXml(text), sst);
  const scanned = readSheetBytes(u8, sst, { compact: Infinity });
  assert.ok(scanned, 'read from its bytes');
  assert.ok(!(scanned.model.cells instanceof SheetCells));
  assert.deepEqual(scanned.model, old);
  const banded = readSheetBytes(u8, sst, { compact: 0 });
  const kept = banded.model.cells instanceof SheetCells ? banded.model.cells.bandCells : 0;
  assert.deepEqual(plain(banded.model), old);
  return { old, model: readSheetBytes(u8, sst, { compact: 0 }).model, kept };
}

test('every cell type, with and without a value', () => {
  const sst = Object.assign(['zero', 'one', 'two'], { runs: new Map([[2, [{ text: 'two', font: null }]]]) });
  same(sheet(
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="n"><v>-2.5</v></c><c r="C1" s="3"><v>1E-3</v></c><c r="D1"><v>0.1</v></c><c r="E1"><v>123456789012345678</v></c><c r="F1"><v>-0</v></c><c r="G1"><v></v></c><c r="H1"><v/></c><c r="I1"><v> 7 </v></c><c r="J1"><v>12.</v></c><c r="K1"><v>.5</v></c><c r="L1"><v>1.7976931348623157E308</v></c><c r="M1"><v>0.30000000000000004</v></c><c r="N1"><v>123456789.123456</v></c><c r="O1"><v>x</v></c></row>' +
    '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c><c r="C2" t="s"><v>9</v></c><c r="D2" t="s"/><c r="E2" t="s"><v></v></c><c r="F2" t="s" s="4"><v> 1 </v></c></row>' +
    '<row r="3"><c r="A3" t="b"><v>1</v></c><c r="B3" t="b"><v>0</v></c><c r="C3" t="b"><v>true</v></c><c r="D3" t="b"/><c r="E3" t="e"><v>#DIV/0!</v></c><c r="F3" t="e"/><c r="G3" t="str"><f>A1&amp;"x"</f><v>1x</v></c><c r="H3" t="str"><v>a &lt; b &amp; c &#65;&#x42; &quot;q&quot; &apos;</v></c><c r="I3" t="d"><v>2024-01-02T03:04:05</v></c><c r="J3" t="foo"><v>4</v></c><c r="K3" t=""><v>5</v></c><c r="L3" t="&#115;"><v>0</v></c></row>' +
    '<row r="4"><c r="A4" t="inlineStr"><is><t>plain</t></is></c><c r="B4" t="inlineStr"><is><t xml:space="preserve"> kept _x000D_ </t></is></c><c r="C4" t="inlineStr"><is><r><rPr><b/><sz val="12"/><rFont val="Arial"/></rPr><t>bold</t></r><r><t xml:space="preserve"> rest</t></r></is></c><c r="D4" t="inlineStr"/><c r="E4" t="inlineStr"><is/></c><c r="F4" t="inlineStr"><is><t/></is></c><c r="G4" t="inlineStr"><is><t>a</t><phoneticPr fontId="1"/></is></c><c r="H4" t="inlineStr"><is> <t>sp&amp;ced</t> </is></c><c r="I4" t="inlineStr"><v>3</v></c></row>',
  ), sst);
});

test('formulas: plain, shared, array, dynamic array and data table', () => {
  same(sheet(
    '<row r="1"><c r="A1"><f>B1*2</f><v>4</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1+$C$1</f><v>1</v></c><c r="C1"><f>IF(A1&lt;2,"a&amp;b",1)</f></c><c r="D1"><f/></c><c r="E1"><f></f><v>1</v></c></row>' +
    '<row r="2"><c r="B2"><f t="shared" si="0"/><v>2</v></c><c r="C2"><f t="shared" si="7"/></c><c r="D2"><f t="array" ref="D2:E3">A1:B2*2</f><v>1</v></c><c r="E2"><v>2</v></c><c r="F2" cm="1"><f t="array" ref="F2:F4">SORT(A1:A3)</f><v>1</v></c><c r="G2"><f t="array">SUM(A1:A2*B1:B2)</f><v>0</v></c></row>' +
    '<row r="3"><c r="B3"><f t="shared" si="0"></f><v>3</v></c><c r="H3"><f t="dataTable" ref="H3:H5" dt2D="0" dtr="0" r1="A1"/><v>1</v></c><c r="I3" cm="1"><f t="array">A1</f></c></row>',
  ));
});

test('rows and cells without r, out of order, repeated, and references that are not', () => {
  same(sheet('<row><c><v>1</v></c><c><v>2</v></c></row><row><c r="C2"><v>3</v></c><c><v>4</v></c></row><row r="10"><c><v>5</v></c></row><row><c r="b11"><v>6</v></c><c r="$D$11"><v>7</v></c><c r="zz"><v>8</v></c><c r="A0"><v>9</v></c><c r="XFE11"><v>10</v></c><c r="XFD11"><v>11</v></c><c r="AAAA1"><v>12</v></c><c r="A12345678"><v>13</v></c></row>'));
  same(sheet('<row r="5"><c r="B5"><v>1</v></c><c r="A5"><v>2</v></c><c r="B5"><v>3</v></c></row><row r="2"><c r="A2" s="1"/></row><row r="5"><c r="C5" s="2"/></row><row r=""><c s="3"/></row><row r="0"><c s="4"/></row>'));
  same(sheet('<row r="1"><c r="A1" s="1"/></row><row r="1"><c r="B1" s="1"/></row>'));
  same(sheet('<row r="2"><c r="B2" s="1"/><c r="A2" s="1"/></row><row r="3"><c r="A3" s="1"/><c r="A3" s="2"/></row>'));
});

test("rows' own attributes", () => {
  same(sheet(
    '<row r="1" ht="20.25" customHeight="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="1"/><row r="3" hidden="true" ht="0"/><row r="4" s="5" customFormat="1"><c r="A4" s="5"/></row>' +
    '<row r="5" s="5"/><row r="6" outlineLevel="2" collapsed="1"/><row r="7" outlineLevel="12"/><row r="8" thickBot="1" spans="1:3" x14ac:dyDescent="0.25" xmlns:x14ac="urn:x"/><row r="9" customHeight="1"/><row r="10" hidden="0" ht="12.5" customHeight="true"/>'.replace(' xmlns:x14ac="urn:x"', ''),
  ).replace('<worksheet ', '<worksheet xmlns:x14ac="urn:x" '));
});

test('quotes, spaces and line breaks in tags; entities in attributes; comments and other elements', () => {
  same(sheet(
    "<row r='1'  ht = \"15\" >\n <c r='A1'\n\ts='2' t='s' ><v>0</v></c> <!-- a > comment <c r=\"Z1\"/> --> <c r=\"B1\"\r\n s=\"&#51;\"><v>2</v></c><c r=\"C1\"><v>3</v><extLst><ext uri=\"x\"><c r=\"Q9\"><v>99</v></c></ext></extLst></c><c r=\"D1\"><v>4</v><v>5</v><f>1</f><f>2</f></c></row>" +
    '<?pi <row/> ?><row r="2"><c r="A2" t="str"><v>line\r\nbreak\rand\ttab</v></c><c r="B2" t="str"><v>&#13;&#10;</v></c><c r="C2" t="str"><f>"a\r\nb"</f><v>x</v></c></row><foo><row r="9"><c r="A9"><v>1</v></c></row></foo>',
  ), ['zero']);
  same(sheet('<row r="1"><c r="A1" t="str"><v>café 中文 😀</v></c><c r="B1" t="inlineStr"><is><t>über</t></is></c></row>'));
});

test('a prefixed sheet, and one with an empty or missing sheetData', () => {
  const p = (rows) => `<x:worksheet xmlns:x="${NS}"><x:sheetData>${rows}</x:sheetData><x:mergeCells><x:mergeCell ref="A1:B2"/></x:mergeCells></x:worksheet>`;
  same(p('<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:r><x:t>a</x:t></x:r><x:r><x:rPr><x:i/></x:rPr><x:t>b</x:t></x:r></x:is></x:c><x:c r="B1"><x:f>A1</x:f><x:v>1</x:v></x:c><x:c r="C1" s="2"/></x:row><x:row r="2"><x:c r="A2" s="2"/></x:row>'));
  same(sheet(''));
  same(`<worksheet xmlns="${NS}"><sheetData/><mergeCells><mergeCell ref="B2:C9"/></mergeCells></worksheet>`);
  assert.equal(readSheetBytes(strToU8(`<worksheet xmlns="${NS}"><dimension ref="A1"/></worksheet>`)), null);
});

test('XML the scan leaves to the parser', () => {
  const nope = (text) => assert.equal(readSheetBytes(strToU8(text)), null);
  nope(sheet('<row r="1"><c r="A1" t="str"><v><![CDATA[x]]></v></c></row>'));
  nope(`<!DOCTYPE worksheet [<!ENTITY e "x">]><worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" t="str"><v>&e;</v></c></row></sheetData></worksheet>`);
  nope(sheet(`<row r="1" xmlns:y="${NS}"><y:c r="A1"><y:v>1</y:v></y:c></row>`));
  nope(`<x:worksheet xmlns:x="${NS}"><x:sheetData><x:row x:r="3"><x:c><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>`);
  nope(`<x:worksheet xmlns:x="${NS}" xmlns="${NS}"><sheetData><row r="3"/></sheetData></x:worksheet>`);
  nope(sheet('<row r="1"><c r="A1" t="str"><v>a<b/>c</v></c></row>'));
});

test('rows of formatted empty cells are kept as bands, and read as the cells they stand for', () => {
  const rows = [];
  for (let r = 1; r <= 40; r++) {
    if (r === 7) rows.push(`<row r="7"><c r="A7" s="2"/><c r="B7" s="2"><v>5</v></c><c r="C7" s="2"/></row>`);
    else if (r === 20) rows.push('<row r="20"><c r="A20" s="2"/><c r="B20" s="3"/><c r="C20" s="2"/><c r="E20" s="2"/></row>');
    else if (r === 30) continue;
    else rows.push(`<row r="${r}"><c r="A${r}" s="2"/><c r="B${r}" s="2"/><c r="C${r}" s="2"/></row>`);
  }
  const { model, old, kept } = same(sheet(rows.join('')));
  const cells = model.cells;
  assert.ok(cells instanceof SheetCells);
  assert.deepEqual(cells.bands.map((b) => [b.r0, b.r1, b.runs]), [[0, 5, [0, 2, 2]], [7, 18, [0, 2, 2]], [19, 19, [0, 0, 2, 1, 1, 3, 2, 2, 2, 4, 4, 2]], [20, 28, [0, 2, 2]], [30, 39, [0, 2, 2]]]);
  assert.equal(kept, 37 * 3 + 4);
  assert.equal(cells.size, 39);
  assert.deepEqual([...cells.filled()].map(([r]) => r), [6]);
  assert.equal(cells.blankStyle(3, 1), 2);
  assert.equal(cells.blankStyle(19, 1), 3);
  assert.equal(cells.blankStyle(19, 3), -1);
  assert.equal(cells.blankStyle(6, 0), -1);
  assert.equal(cells.blankStyle(29, 0), -1);
  assert.equal(cells.has(29), false);
  assert.equal(cells.get(29), undefined);
  // (A row asked for is made, and is the Map's own from then on: a change to it stays.)
  assert.deepEqual(cells.get(10), old.cells.get(10));
  cells.get(10).get(1).s = 9;
  assert.equal(cells.get(10).get(1).s, 9);
  assert.equal(cells.blankStyle(10, 1), -1);
  assert.deepEqual(cells.bands.map((b) => [b.r0, b.r1]), [[0, 5], [7, 9], [11, 18], [19, 19], [20, 28], [30, 39]]);
  assert.equal(cells.delete(0), true);
  assert.equal(cells.has(0), false);
  assert.equal(cells.delete(0), false);
  cells.set(39, new Map());
  assert.equal(cells.get(39).size, 0);
  assert.equal(cells.size, 38);
  // (Walked, every row is there, in row order.)
  assert.deepEqual([...cells.keys()], [...old.cells.keys()].filter((r) => r !== 0));
  assert.equal(cells.bands.length, 0);
  assert.equal(cells.get(38).get(2).s, 2);
});

test('every fixture, read both ways', async () => {
  const dirs = ['test/fixtures', 'test/out'].filter((d) => fs.existsSync(d));
  let sheets = 0;
  for (const dir of dirs) {
    for (const name of fs.readdirSync(dir).filter((f) => /\.xls[xm]$/.test(f))) {
      let files;
      try { files = unzipSync(fs.readFileSync(`${dir}/${name}`)); } catch { continue; } // (an encrypted one)
      const sstPath = Object.keys(files).find((p) => /sharedStrings\.xml$/i.test(p));
      const sst = readSharedStrings(sstPath ? parseXml(strFromU8(files[sstPath])) : null);
      for (const [path, u8] of Object.entries(files)) {
        if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue;
        const old = readSheet(parseXml(strFromU8(u8)), sst);
        for (const compact of [Infinity, 0]) {
          const scanned = readSheetBytes(u8, sst, { compact });
          assert.ok(scanned, `${name} ${path}: read from its bytes`);
          assert.deepEqual(plain(scanned.model), old, `${name} ${path}`);
        }
        sheets++;
      }
    }
  }
  assert.ok(sheets >= 30, `${sheets} sheets`);
});

test('a million rows of formatted empty cells: read in bounded time and memory', () => {
  const ROWS = 1000000;
  const parts = [];
  for (let r = 1; r <= ROWS; r++) parts.push(r === 500 ? '<row r="500" spans="1:10"><c r="A500" s="1"/><c r="B500" s="1"><v>42</v></c></row>' : `<row r="${r}" spans="1:10"><c r="A${r}" s="1"/><c r="B${r}" s="1"/><c r="C${r}" s="1"/><c r="D${r}" s="2"/><c r="E${r}" s="1"/><c r="F${r}" s="1"/><c r="G${r}" s="1"/><c r="H${r}" s="1"/><c r="I${r}" s="1"/><c r="J${r}" s="1"/></row>`);
  const u8 = strToU8(sheet(parts.join('')));
  parts.length = 0;
  assert.ok(u8.length > 150e6);
  globalThis.gc?.();
  const heap = process.memoryUsage().heapUsed;
  const t0 = performance.now();
  const { model } = readSheetBytes(u8);
  const ms = performance.now() - t0;
  const grew = process.memoryUsage().heapUsed - heap;
  assert.ok(model.cells instanceof SheetCells);
  assert.equal(model.cells.bands.length, 2);
  assert.equal(model.cells.bandCells, (ROWS - 1) * 10);
  assert.equal(model.maxRow, ROWS - 1);
  assert.equal(model.maxCol, 9);
  assert.equal(model.cells.size, ROWS);
  assert.equal(model.cells.get(499).get(1).v, 42);
  assert.equal(model.cells.blankStyle(ROWS - 1, 3), 2);
  assert.deepEqual(model.cells.get(ROWS - 1).get(3), { t: 'n', v: null, s: 2, f: null });
  // (Seconds, not the minute and the gigabytes ten million cell objects and their tree took; with room for a slow machine.)
  assert.ok(ms < 20000, `${Math.round(ms)} ms`);
  if (process.env.SCAN_TIMES) console.log(`${Math.round(ms)} ms, ${Math.round(grew / 1e6)} MB`);
  assert.ok(grew < 64e6, `${Math.round(grew / 1e6)} MB`);
});

test('a workbook shown, worked out and saved without its sheets being parsed whole; an edit parses the one it changes', () => {
  const bytes = fs.readFileSync('test/fixtures/fills.xlsx');
  const wb = Workbook.fromBytes(bytes);
  wb.calculateMissing();
  const before = unzipSync(bytes);
  const model = wb.sheet(0);
  const part = wb._part(0);
  assert.equal(part.hasDoc, false);
  sheetProtection(wb, 0);
  readHyperlinks(wb, 0);
  wb.tabColor(0);
  usedRange(model);
  assert.equal(part.hasDoc, false);
  // (Unchanged, every part goes back as it came.)
  const saved = unzipSync(wb.toBytes());
  for (const [path, u8] of Object.entries(before)) assert.deepEqual(saved[path], u8, path);
  assert.equal(part.hasDoc, false);
  // (The model is the one the parsed part gives.)
  assert.deepEqual(plain(model), plain(readSheet(parseXml(strFromU8(before[wb.sheets[0].path])), wb.sharedStrings)));
  // An undo point taken before the edit, the edit, and back.
  const snap = wb.snapshot();
  const was = model.cells.get(0)?.get(0)?.v;
  wb.setCell(0, 'A1', 'changed here');
  assert.equal(part.hasDoc, true);
  assert.equal(wb.sheet(0).cells.get(0).get(0).v, 'changed here');
  const again = Workbook.fromBytes(wb.toBytes());
  assert.equal(again.sheet(0).cells.get(0).get(0).v, 'changed here');
  wb.restore(snap);
  assert.equal(wb.sheet(0).cells.get(0)?.get(0)?.v, was);
});

test('what reads a banded sheet without making its rows answers as for the rows themselves', () => {
  const rows = ['<row r="1"><c r="A1" t="str"><v>title</v></c></row>'];
  for (let r = 2; r <= 60; r++) rows.push(r === 9 ? '<row r="9" ht="30" customHeight="1"><c r="A9" s="1"/></row>' : r === 12 ? '<row r="12" s="2" customFormat="1" thickBot="1"><c r="B12" s="1"/><c r="C12" s="2"/></row>' : `<row r="${r}"><c r="B${r}" s="1"/><c r="C${r}" s="2"/><c r="D${r}" s="1"/></row>`);
  const text = sheet(rows.join(''), '<mergeCells count="2"><mergeCell ref="B20:C20"/><mergeCell ref="C30:D33"/></mergeCells>');
  const banded = readSheetBytes(strToU8(text), [], { compact: 0 }).model;
  const whole = readSheet(parseXml(text), []);
  assert.ok(banded.cells.bands.length > 3);
  // (Three fonts: 11, 20 and 14 points.)
  const deco = (s) => ({ font: { family: 'Calibri', sizePx: [14.67, 26.67, 18.67][s] ?? 14.67, css: '' }, align: {} });
  const cols = { size: () => 64 };
  const sorted = (m) => new Map([...m].sort((a, b) => a[0] - b[0]));
  assert.deepEqual(sorted(autoRowHeights(banded, deco, cols)), sorted(autoRowHeights(whole, deco, cols)));
  assert.deepEqual(sorted(rowDescents(banded, deco)), sorted(rowDescents(whole, deco)));
  assert.deepEqual(usedRange(banded), usedRange(whole));
  assert.equal(cellCount(banded.cells), cellCount(whole.cells));
  assert.equal(banded.cells.bands.length > 3, true, 'none of it made the rows');
  assert.deepEqual([...filledRows(banded.cells)].map(([r]) => r), [0]);
});
