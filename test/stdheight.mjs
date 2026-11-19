// Browser check of a sheet's standard row height (src/editor/rowheights.js standardRowHeight) against Excel's own for
// tools/excel-compare/stdheight-book.mjs's workbooks (test/stdheight-excel.json, excel-stdheight.ps1): stored
// defaultRowHeights with and without customHeight, and column styles of smaller and larger fonts. Before, the stored
// defaultRowHeight was the least height of every row: a sheet saved on a screen at 125% (14.4) had all its rows 1px
// short of Excel's 15, and one whose columns are all styled in a smaller font a line too tall; 786 rows of 146
// customer workbooks were 1px off.
// Run after `npm run build`: node test/stdheight.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { STDHEIGHT_CASES, stdheightBook } from '../tools/excel-compare/stdheight-book.mjs';

const { page, check, finish } = await editingPage();
// (trimStart takes off the byte order mark PowerShell writes.)
const excel = JSON.parse(fs.readFileSync('test/stdheight-excel.json', 'utf8').trimStart()).books;

for (const c of STDHEIGHT_CASES) {
  const want = excel[c.name];
  const name = `${c.name}.xlsx`;
  await page.setInputFiles('input[type=file]', { name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(stdheightBook(c)) });
  await page.waitForFunction((n) => window.oxe.fileName === n, name);
  await page.waitForTimeout(150);
  // (Rows 1-13 and row 150, far below anything the file has: the standard height.)
  const ours = await page.evaluate((n) => { const { rows } = window.oxe.grid.geometry(); return [...Array.from({ length: n }, (_, r) => rows.size(r) * 0.75), rows.size(149) * 0.75]; }, want.heights.length);
  const off = [...want.heights, want.std].map((pt, k) => [k < want.heights.length ? `row ${k + 1}` : 'row 150', ours[k], pt]).filter(([, a, b]) => Math.abs(a - b) > 0.5);
  check(!off.length, `${c.name}: standard height ${want.std}pt and rows 1-${want.heights.length} as Excel's`, off.map(([r, a, b]) => `${r}: ${a} (Excel ${b})`).join('; '));
}

await finish();
