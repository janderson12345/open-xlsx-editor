// Browser check: open each fixture in the demo (standards mode asserted), see the first sheet drawn, scroll the large
// one and check only the window in view is in the DOM, save unchanged (every part byte for byte) and after an edit.
// Run after `npm run build`: node test/e2e.mjs [outDir]
import fs from 'node:fs';
import { unzipSync } from 'fflate';
import { openDemo, openFixture, checker } from './browser.mjs';

const outDir = process.argv[2] || 'test/out';
fs.mkdirSync(outDir, { recursive: true });
const { page, errors, close } = await openDemo();
const t = checker();
const { check } = t;

/** The text of the drawn cell at a 0-based row and column, or null if it isn't in the DOM. */
const cellTextAt = (r, c) => page.evaluate(([r, c]) => {
  const el = document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="${c + 2}"]`);
  return el ? el.textContent : null;
}, [r, c]);
const box = (sel) => page.evaluate((s) => { const b = document.querySelector(s)?.getBoundingClientRect(); return b ? { x: b.x, y: b.y, w: b.width, h: b.height } : null; }, sel);
const domCells = () => page.evaluate(() => document.querySelectorAll('.oxe-cell').length);

for (const f of fs.readdirSync('test/fixtures').filter((n) => n.endsWith('.xlsx'))) {
  await openFixture(page, f);
  await page.screenshot({ path: `${outDir}/${f}.png` });
  const first = await page.evaluate(() => document.querySelector('.oxe-row .oxe-cell')?.textContent ?? null);
  check(first != null && first.length > 0, `${f}: first sheet drawn ("${first}")`);
  // Unchanged save: every part byte for byte.
  const src = unzipSync(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`)));
  const saved = new Uint8Array(await page.evaluate(() => Array.from(window.oxe.save())));
  fs.writeFileSync(`${outDir}/${f.replace('.xlsx', '.saved.xlsx')}`, saved);
  const out = unzipSync(saved);
  const changed = Object.keys(src).filter((k) => !out[k] || !Buffer.from(out[k]).equals(Buffer.from(src[k])));
  check(!changed.length && Object.keys(out).length === Object.keys(src).length, `${f}: saved unchanged, every part byte for byte`, changed.join(', '));
  // An edit through the API: only the sheet changes (the file goes to check-ooxml in CI).
  const edited = new Uint8Array(await page.evaluate(() => { window.oxe.workbook.setCell(window.oxe.sheetIndex, 'Z200', 'edited in the browser'); return Array.from(window.oxe.save()); }));
  fs.writeFileSync(`${outDir}/${f.replace('.xlsx', '.edited.xlsx')}`, edited);
  const eo = unzipSync(edited);
  const sheetPath = await page.evaluate(() => window.oxe.workbook.sheets[window.oxe.sheetIndex].path);
  const touched = Object.keys(src).filter((k) => !Buffer.from(eo[k]).equals(Buffer.from(src[k])));
  check(touched.length === 1 && touched[0] === sheetPath, `${f}: an edit rewrites only ${sheetPath}`, touched.join(', '));
}

// basic.xlsx: values, widths, a hidden column, a tall and a hidden row, a merged range, text running on.
await openFixture(page, 'basic.xlsx');
check(await cellTextAt(0, 0) === 'Item', 'basic: A1 "Item"');
check(await cellTextAt(10, 5) === '1234567.891', 'basic: F11 1234567.891 in General');
check(await cellTextAt(11, 4) === 'Rich bold', 'basic: rich text without its phonetic run');
check(await cellTextAt(1, 3) === null, 'basic: hidden column D not drawn');
check(await cellTextAt(5, 0) === null, 'basic: hidden row 6 not drawn');
const a = await box('.oxe-colhead[aria-colindex="2"]');
const b = await box('.oxe-colhead[aria-colindex="3"]');
check(a?.w === 64 && b?.w === 145, `basic: column A 64px (default), B 145px (width 20.7109375 at a 7px digit)`, JSON.stringify([a, b]));
const r4 = await box('.oxe-row[aria-rowindex="5"]');
const r3 = await box('.oxe-row[aria-rowindex="4"]');
check(r3?.h === 20 && r4?.h === 40, 'basic: row 3 20px (15pt), row 4 40px (30pt)', JSON.stringify([r3, r4]));
const merged = await box('.oxe-cell.oxe-merged');
check(merged?.w === 64 + 145 + 64 - 1 && merged?.h === 39, 'basic: A12:C13 drawn as one cell over three columns and two rows', JSON.stringify(merged));
const spill = await box('.oxe-row[aria-rowindex="16"] .oxe-cell');
check(spill?.w > 400, 'basic: long text in A15 runs on over the empty cells to its right', JSON.stringify(spill));
const rowHead = await box('.oxe-row[aria-rowindex="2"] .oxe-rowhead');
const cellA1 = await box('.oxe-row[aria-rowindex="2"] .oxe-cell[aria-colindex="2"]');
check(rowHead && cellA1 && Math.abs(rowHead.x + rowHead.w - cellA1.x) < 0.5, 'basic: A1 starts where the row heading ends');

// large.xlsx: 100,000 cells, only those in view drawn; scrolling draws the new window.
await openFixture(page, 'large.xlsx');
const n0 = await domCells();
check(n0 > 0 && n0 < 1500, `large: ${n0} cell elements in the DOM for 100,000 cells`);
await page.evaluate(() => window.oxe.grid.scrollToCell(4900, 12));
check(await cellTextAt(4900, 0) === null && await cellTextAt(4900, 12) === String(4901 * 100 + 12), 'large: scrolled to M4901, its value drawn');
await page.evaluate(() => { const g = document.querySelector('.oxe-grid'); g.scrollTop = 0; g.scrollLeft = 0; });
await page.mouse.move(600, 500);
await page.mouse.wheel(0, 20 * 300);
await page.waitForFunction(() => window.oxe.grid.window().r0 > 250);
const win = await page.evaluate(() => window.oxe.grid.window());
check(await cellTextAt(win.r0 + 5, 0) === `Row ${win.r0 + 6}`, `large: a wheel scroll redraws rows ${win.r0 + 1}-${win.r1 + 1}`);
const n1 = await domCells();
check(n1 < 1500, `large: still ${n1} cell elements after scrolling`);
const sticky = await box('.oxe-colheads');
check(sticky && Math.abs(sticky.y - (await box('.oxe-grid')).y) < 0.5, 'large: column headings stay at the top while scrolled');

// Printing: the pages are built, the grid is hidden in print media, the paper size reaches @page.
const printed = await page.evaluate(() => window.oxe.preparePrint());
check(printed === 336, `large: printed as ${printed} pages (112 down by 3 across)`);
await page.emulateMedia({ media: 'print' });
const printView = await page.evaluate(() => ({
  grid: getComputedStyle(document.querySelector('.oxe-grid')).display,
  pages: document.querySelectorAll('.oxe-page').length,
  first: document.querySelector('.oxe-page .oxe-cell')?.textContent,
  size: document.getElementById('oxe-page-style')?.textContent,
}));
check(printView.grid === 'none' && printView.pages === 336 && printView.first === 'Column A' && /size: 8.5in 11in/.test(printView.size), 'large: print media shows the pages, not the grid', JSON.stringify(printView));
await page.emulateMedia({ media: 'screen' });

check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
