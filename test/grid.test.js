import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Axis } from '../src/editor/axis.js';
import { sheetGeometry } from '../src/editor/geometry.js';
import { Workbook } from '../src/ooxml/workbook.js';

test('an axis finds the index under a position, passing over hidden ones', () => {
  const a = new Axis(10, 20, new Map([[2, 40], [3, 0]]));
  assert.equal(a.end, 20 * 8 + 40);
  assert.equal(a.start(4), 80);
  assert.equal(a.size(3), 0);
  assert.deepEqual([0, 19, 20, 39, 40, 79, 80, 99, 100, 10000].map((p) => a.indexAt(p)), [0, 0, 1, 1, 2, 2, 4, 4, 5, 9]);
});

test('sheet geometry from basic.xlsx: default 64px columns and 20px rows, its own widths and heights, hidden ones 0', () => {
  const sheet = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'))).sheet(0);
  const { rows, cols } = sheetGeometry(sheet);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((c) => cols.size(c)), [64, 145, 64, 0, 64, 215]);
  assert.deepEqual([0, 3, 5].map((r) => rows.size(r)), [20, 40, 0]);
  assert.ok(rows.count > sheet.maxRow && cols.count > sheet.maxCol);
});

// Excel's PDF of test/fixtures/rows.xlsx: rows without customHeight sized by their fonts and wrapped lines.
test('automatic row heights: tallest font, styled empty cells count, stored ht without customHeight ignored', async () => {
  const { fontRowPx } = await import('../src/editor/fonts.js');
  const { autoRowHeights, rowDescents } = await import('../src/editor/rowheights.js');
  const pt = (name, size) => fontRowPx(name, (size * 4) / 3) * 0.75;
  assert.deepEqual([['Calibri', 11], ['Calibri', 14], ['Calibri', 20], ['Calibri', 36], ['Cambria', 20], ['Times New Roman', 16], ['Arial', 14]].map(([n, s]) => pt(n, s)),
    [15, 18.75, 26.25, 46.5, 25.5, 20.25, 18]);
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/rows.xlsx')));
  const sheet = wb.sheet(0);
  const fonts = wb.styles.fonts;
  const deco = (s) => {
    const xf = wb.styles.cellXfs[s];
    const f = fonts[xf.fontId];
    return { font: { family: f.scheme ? 'Calibri' : f.name, sizePx: (f.size * 4) / 3, css: '' }, align: xf.align };
  };
  const h = autoRowHeights(sheet, deco, sheetGeometry(sheet).cols);
  // (Rows 1-based in the file: 2 Calibri 8, 5 Calibri 20, 12 a styled empty 20pt cell, 17 ht 30 not custom, 18 ht 10
  // not custom with a 20pt cell, 19 three lines by line breaks; Node counts wrapped lines by line breaks only.)
  assert.deepEqual([1, 4, 11, 16, 17, 18].map((r) => h.get(r)), [15, 26.25, 26.25, 15, 26.25, 45]);
  const d = rowDescents(sheet, deco);
  assert.ok(d.get(10) > d.get(2)); // (the mixed row sits on the 20pt baseline)
});

test("a row grows for its own format's font and a rich-text run's, as Excel's: 14pt to 18.75, a 16pt run to 21 (measured over COM)", async () => {
  const { autoRowHeights } = await import('../src/editor/rowheights.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Grow');
  const big = wb.styleWith(0, { font: { size: 14 } });
  wb.setLineStyles(si, { rows: [{ r: 6, s: big }] });
  wb.setCell(si, 'A1', 'Plain big');
  const sheet = wb.sheet(si);
  sheet.cells.get(0).get(0).runs = [{ text: 'Plain ', font: null }, { text: 'big', font: { name: 'Calibri', size: 16, scheme: 'minor', b: false, i: false, u: null, strike: false, color: null, vertAlign: null } }];
  const deco = (s) => {
    const f = wb.styles.fonts[wb.styles.cellXfs[s].fontId];
    return { font: { family: 'Calibri', sizePx: (f.size * 4) / 3, css: '' }, align: wb.styles.cellXfs[s].align };
  };
  const h = autoRowHeights(sheet, deco, sheetGeometry(sheet).cols);
  assert.equal(h.get(6), 18.75);
  assert.equal(h.get(0), 21);
});
