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
// (The documented difference, docs/LAYOUT.md: rows 32 and 45's 17-character identifiers, mostly digits, just fit on
// one line in Excel, whose digits are exactly 7px; counted at EXCEL_WRAP_WIDTH they take two here.)
const KNOWN = new Set([32, 45]);
const unexpected = off.filter(([r]) => !KNOWN.has(r));
check(!unexpected.length, `row heights: ${excel.length - off.length} of ${excel.length} rows as Excel's, the others the documented ${[...KNOWN].join(' and ')}`, unexpected.slice(0, 8).map(([r, a, b]) => `row ${r}: ${a} (Excel ${b})`).join('; '));
check(Math.abs(ours[0] - 21.75) < 0.5, 'the merged 16pt title row with a thick bottom border: 21.75pt, as Excel', String(ours[0]));

await finish();
