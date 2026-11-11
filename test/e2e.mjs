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
  // (The first cell with text: the active cell, A1, is drawn even when empty.)
  const first = await page.evaluate(() => [...document.querySelectorAll('.oxe-row .oxe-cell')].find((e) => e.textContent)?.textContent ?? null);
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
  // (The workbook part may change too, only to ask Excel to work the formulas out on opening: when the sheet holds a
  // #SPILL! or #CALC!, saved as #VALUE! for Excel to show as its own, docs/DYNAMIC.md.)
  const book = (bytes) => Buffer.from(bytes).toString('utf8').replace(/<calcPr fullCalcOnLoad="1"\/>/, '').replace(/ fullCalcOnLoad="1"/, '');
  const touched = Object.keys(src).filter((k) => !Buffer.from(eo[k]).equals(Buffer.from(src[k])) && !(k === 'xl/workbook.xml' && book(eo[k]) === book(src[k])));
  check(touched.length === 1 && touched[0] === sheetPath, `${f}: an edit rewrites only ${sheetPath}`, touched.join(', '));
}

// formulas.xlsx is saved without formula results: worked out on opening, as Excel does (Calc!B2 =SUM over Data!B2:B9).
await openFixture(page, 'formulas.xlsx');
check(await cellTextAt(1, 1) === '3', 'formulas.xlsx: a formula saved without its result shows Excel\'s result on opening', await cellTextAt(1, 1));

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
check(printed === 342, `large: printed as ${printed} pages (114 down by 3 across, as Excel prints it)`);
await page.emulateMedia({ media: 'print' });
const printView = await page.evaluate(() => ({
  grid: getComputedStyle(document.querySelector('.oxe-grid')).display,
  pages: document.querySelectorAll('.oxe-page').length,
  first: document.querySelector('.oxe-page .oxe-cell')?.textContent,
  size: document.getElementById('oxe-page-style')?.textContent,
}));
check(printView.grid === 'none' && printView.pages === 342 && printView.first === 'Column A' && /size: 8.5in 11in/.test(printView.size), 'large: print media shows the pages, not the grid', JSON.stringify(printView));
await page.emulateMedia({ media: 'screen' });

// Notes: notes.xlsx's two notes marked red, its thread purple; B2 described by its note; the pointer resting on C3
// shows the thread with both people; Shift+F2 on B2 shows its note with the author in bold, and hides it again.
await openFixture(page, 'notes.xlsx');
const marks = await page.evaluate(() => ({ red: document.querySelectorAll('.oxe-grid > * .oxe-note-mark:not(.oxe-thread-mark)').length, purple: document.querySelectorAll('.oxe-grid > * .oxe-thread-mark').length }));
const b2 = await page.evaluate(() => { const cell = [...document.querySelectorAll('.oxe-row[aria-rowindex="3"] .oxe-cell')].find((x) => x.getAttribute('aria-colindex') === '3'); const id = cell?.getAttribute('aria-describedby'); return id && document.getElementById(id)?.textContent; });
check(marks.red === 2 && marks.purple === 1 && b2 === 'Note. Avery Chen: Check this total', 'notes: two red marks, one purple; B2 described by its note', `${JSON.stringify(marks)} ${b2}`);
const c3 = await page.evaluate(() => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const { rows, cols, rowHeadW, colHeadH } = g.geometry(); return { x: v.left + rowHeadW + cols.start(2) + 10, y: v.top + colHeadH + rows.start(2) + 6 }; });
await page.mouse.move(c3.x, c3.y);
await page.waitForSelector('.oxe-thread-box', { timeout: 2000 }).catch(() => {});
const thread = await page.evaluate(() => document.querySelector('.oxe-thread-box')?.textContent || '');
check(/Avery Chen/.test(thread) && /Sam Rivera/.test(thread) && /checked against the invoice/.test(thread), 'the pointer on C3: the thread, with its reply', thread);
await page.mouse.move(5, 5);
await page.evaluate(() => window.oxe.select('B2'));
// (A cell's menu, Show/Hide Note: the note beside it with the author in bold; again: hidden.)
const b2At = await page.evaluate(() => { const r = document.querySelector('.oxe-sel-active').getBoundingClientRect(); return { x: r.left + 10, y: r.top + 6 }; });
const showHide = async () => { await page.mouse.click(b2At.x, b2At.y, { button: 'right' }); await page.click('.oxe-menu >> text=Show/Hide Note'); };
await showHide();
const note = await page.evaluate(() => { const b = document.querySelector('.oxe-note-box:not(.oxe-thread-box)'); return b ? { text: b.textContent, bold: b.querySelector('strong')?.textContent } : null; });
await showHide();
check(note?.bold === 'Avery Chen:' && /Check this total/.test(note.text) && !(await page.isVisible('.oxe-note-box')), 'Show/Hide Note on B2: its note, the author in bold; again: hidden', JSON.stringify(note));
// Shift+F2 on B2: Excel's Edit Note, its text to type in; changed and OK: the note's new text, the author kept; Ctrl+Z.
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+F2');
await page.waitForSelector('dialog.oxe-note-edit');
const editing = await page.evaluate(() => ({ title: document.querySelector('dialog.oxe-note-edit h2').textContent, text: document.activeElement.value }));
await page.fill('dialog.oxe-note-edit textarea', 'Checked on Monday');
await page.click('dialog.oxe-note-edit >> text=OK');
const edited = await page.evaluate(() => document.getElementById(document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="3"]').getAttribute('aria-describedby'))?.textContent);
check(editing.title === 'Edit Note for B2' && editing.text === 'Check this total' && edited === 'Note. Avery Chen: Checked on Monday', 'Shift+F2: Edit Note for B2 with its text; OK: the new text, Avery Chen still its author', `${JSON.stringify(editing)} ${edited}`);
await page.keyboard.press('Control+z');
const undone = await page.evaluate(() => document.getElementById(document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="3"]').getAttribute('aria-describedby'))?.textContent);
check(undone === 'Note. Avery Chen: Check this total', 'Ctrl+Z: the note as it was', undone);

// Conditional formats: basic.xlsx's B2:B10 bold when over 5: B7 (6) bold, B5 (4) not; 50 typed into B2: bold at once.
await openFixture(page, 'basic.xlsx');
await page.evaluate(() => window.oxe.setReadOnly(false));
const boldIn = (ref) => page.evaluate((ref) => { const [, col, row] = /([A-Z]+)(\d+)/.exec(ref); const cell = [...document.querySelectorAll(`.oxe-grid .oxe-row[aria-rowindex="${Number(row) + 1}"] .oxe-cell`)].find((x) => x.getAttribute('aria-colindex') === String(col.charCodeAt(0) - 63)); return cell ? getComputedStyle(cell).fontWeight : null; }, ref);
const [b7, b5] = [await boldIn('B7'), await boldIn('B5')];
await page.evaluate(() => window.oxe.enter('B2', '50'));
const b2Bold = await boldIn('B2');
check(b7 === '700' && b5 === '400' && b2Bold === '700', 'conditional format: B7 (6) bold, B5 (4) not; B2 typed 50: bold', `${b7} ${b5} ${b2Bold}`);

// Charts: charts.xlsx's four drawn over the cells, each also a figure with its data for a screen reader; 500 typed into
// B7 (June, North): the first chart's axis goes to 600, as Excel's would.
await openFixture(page, 'charts.xlsx');
await page.evaluate(() => window.oxe.setReadOnly(false));
await page.waitForSelector('.oxe-chart svg');
const figs = await page.evaluate(() => [...document.querySelectorAll('.oxe-chart-figures figcaption')].map((x) => x.textContent));
await page.evaluate(() => window.oxe.enter('B7', '500'));
const topLabel = await page.evaluate(() => Math.max(...[...document.querySelector('.oxe-chart svg').querySelectorAll('text')].map((t) => Number(t.textContent)).filter(Number.isFinite)));
check(await page.locator('.oxe-chart').count() === 4 && figs.join('|') === 'Column chart: Sales by month|Line chart: Chart 2|Pie chart: North|Bar chart: Sales, stacked' && topLabel === 600,
  'charts: four drawn, each a figure for a screen reader; B7 typed 500: the first one\'s axis to 600', `${figs.join('|')} ${topLabel}`);

check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
