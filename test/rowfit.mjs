// Browser check of the row heights Excel works out itself (src/editor/rowheights.js), against Excel's own for
// tools/excel-compare/rowfit-book.mjs's workbook (test/rowfit-excel.json, excel-rowfit.ps1): a merged 16pt title row
// with a thick bottom border, wrapped bold headings, and 60 rows of a wrapped formula's long text and a wrapped
// sentence, none with customHeight. Before, a formula's text and a word wider than its column didn't count, and lines
// were counted at the canvas's widths, narrower than Excel's; a customer's crosswalk workbook showed every such row a
// line short.
// Run after `npm run build`: node test/rowfit.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { rowfitBook } from '../tools/excel-compare/rowfit-book.mjs';

const { page, check, finish } = await editingPage();
// (trimStart takes off the byte order mark PowerShell writes.)
const excel = JSON.parse(fs.readFileSync('test/rowfit-excel.json', 'utf8').trimStart()).heights;

await page.setInputFiles('input[type=file]', { name: 'rowfit.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(rowfitBook()) });
await page.waitForFunction(() => window.oxe.fileName === 'rowfit.xlsx');
await page.waitForTimeout(300);
const ours = await page.evaluate((n) => { const { rows } = window.oxe.grid.geometry(); return Array.from({ length: n }, (_, r) => rows.size(r) * 0.75); }, excel.length);
const off = ours.map((pt, r) => [r + 1, pt, excel[r]]).filter(([, a, b]) => Math.abs(a - b) > 0.5);
// (All 62 since lines are counted at Windows' whole-pixel widths and Excel's margin, rowheights.js excelLineCount:
// rows 32 and 45's 17-character identifiers, mostly digits, just fit on one line in Excel, whose digits are exactly
// 7px, and took two when counted at the canvas's widths.)
// (Exact on Windows. Elsewhere, Linux CI among them, a character the width table hasn't is measured in Calibri's
// stand-in, so rows are checked to be no shorter than Excel's, at most two lines taller, and most of them Excel's.)
if (process.platform === 'win32') {
  check(!off.length, `row heights: ${excel.length - off.length} of ${excel.length} rows as Excel's`, off.slice(0, 8).map(([r, a, b]) => `row ${r}: ${a} (Excel ${b})`).join('; '));
} else {
  const wrong = off.filter(([, a, b]) => a < b - 0.5 || a > b + 30.5);
  check(!wrong.length && off.length <= excel.length / 4, `row heights: ${excel.length - off.length} of ${excel.length} rows as Excel's, the rest at most two lines taller (not Windows: Calibri's stand-in is wider)`, (wrong.length ? wrong : off).slice(0, 8).map(([r, a, b]) => `row ${r}: ${a} (Excel ${b})`).join('; '));
}
check(Math.abs(ours[0] - 21.75) < 0.5, 'the merged 16pt title row with a thick bottom border: 21.75pt, as Excel', String(ours[0]));
// (Row 63: some hundred wrapped lines, which Excel stops at its tallest row, 409.5pt; it grew to thousands before.)
const tall = await page.evaluate(() => window.oxe.grid.geometry().rows.size(62) * 0.75);
check(tall === 409.5, 'a wrapped text taller than a row can be: 409.5pt, Excel\'s tallest row', String(tall));

await finish();
