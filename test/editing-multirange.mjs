// Browser check of commands over several ranges selected (Ctrl+click), as Excel did them (test/multirange-excel.json):
// a new sheet with A1:F12 filled with row * 10 + column; Fill, Insert and Delete, Hide, Row Height, Merge on each
// range; Sort, Filter, Text to Columns, Remove Duplicates refused in Excel's words, Insert Table greyed out; Find
// within the ranges; the status bar's Sum; Print Selection's pages; names that refer to every range.
// Run after `npm run build`: node test/editing-multirange.mjs [outDir]
import { openFixture } from './browser.mjs';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, sel, ribbon } = await editingPage();
await openFixture(page, 'basic.xlsx');

const MULTI = "This can't be done on a multiple range selection.";
const run = (name, ...args) => page.evaluate(([n, a]) => window.oxe.command(n, ...a), [name, args]);
/** A fresh sheet shown, A1:F12 filled with row * 10 + column. */
const freshGrid = async () => {
  await run('insertSheet');
  await page.evaluate(() => { for (let r = 1; r <= 12; r++) for (let c = 0; c < 6; c++) window.oxe.enter(`${String.fromCharCode(65 + c)}${r}`, String(r * 10 + c + 1)); });
};
/** Ranges selected as Ctrl+click and drag selects them ('A2:B4,D6:E8'). */
const pick = (text) => page.evaluate((t) => {
  const at = (ref) => ({ r: Number(ref.replace(/[A-Z]+/, '')) - 1, c: ref.match(/[A-Z]+/)[0].charCodeAt(0) - 65 });
  t.split(',').forEach((part, k) => {
    const [a, b = a] = part.split(':').map(at);
    window.oxe.command(k ? 'addArea' : 'select', a.r, a.c);
    window.oxe.command('extendTo', b.r, b.c);
  });
}, text);
const values = (refs) => page.evaluate((list) => list.map((ref) => {
  const r = Number(ref.replace(/[A-Z]+/, '')) - 1;
  const c = ref.match(/[A-Z]+/)[0].charCodeAt(0) - 65;
  return window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c)?.v ?? null;
}), refs);
const dialogText = async () => { await page.waitForSelector('dialog.oxe-dialog[open]'); return page.textContent('dialog.oxe-dialog[open]'); };
const okDialog = () => page.click('dialog.oxe-dialog[open] >> text=OK');

// Fill Down (Ctrl+D) over A2:B4 and A7:B9: each column from its first cell selected, as Excel's (A7:B9 from A2:B2).
await freshGrid();
await pick('A2:B4,A7:B9');
await page.focus('.oxe-grid');
await page.keyboard.press('Control+d');
const filled = await values(['A3', 'B4', 'A5', 'A7', 'B9']);
check(filled.join() === '21,22,51,21,22', 'Fill Down over A2:B4 and A7:B9: A3:B4 and A7:B9 from A2:B2, A5 left, as Excel', filled.join());
await page.keyboard.press('Control+z');
check((await values(['A3', 'A7'])).join() === '31,71', 'Ctrl+Z: the fill taken back in one step');

// Delete Sheet Rows over A2:B3 and D6:E6: rows 2, 3 and 6 go, as Excel's.
await pick('A2:B3,D6:E6');
await run('lines', 'row', -1);
const rowsLeft = await values(['A2', 'A3', 'A4', 'A5']);
check(rowsLeft.join() === '41,51,71,81', 'Delete Sheet Rows over two ranges: rows 2, 3 and 6 deleted, as Excel', rowsLeft.join());
await page.keyboard.press('Control+z');
// Insert Cells shifted down over A2:B3 and D6:E6: each range's cells, as Excel's.
await pick('A2:B3,D6:E6');
await run('cells', 1, 'row');
const shifted = await values(['A2', 'A4', 'D6', 'D7']);
check(shifted.join() === ',21,,64', 'Insert Cells (shift down) over two ranges: each range blank, its cells moved down', shifted.join());
await page.keyboard.press('Control+z');
// Overlapping ranges: Excel's refusal.
await pick('A2:B4,B3:C5');
await run('cells', -1, 'row');
const overlap = await dialogText();
check(/Cannot use that command on overlapping selections\./.test(overlap), "Delete Cells over overlapping ranges: Excel's refusal", overlap);
await okDialog();

// Hide (Ctrl+9) and Row Height over two ranges: the rows of each.
await pick('A2:B3,D6:E6');
await page.focus('.oxe-grid');
await page.keyboard.press('Control+9');
const hidden = await page.evaluate(() => [1, 2, 3, 4, 5, 6].filter((r) => window.oxe.workbook.sheet(window.oxe.sheetIndex).rows.get(r)?.hidden).map((r) => r + 1).join());
check(hidden === '2,3,6', 'Hide Rows over two ranges: rows 2, 3 and 6 hidden, as Excel', hidden);
await page.keyboard.press('Control+z');
await pick('A2:A3,D6:E6');
await run('columnWidth', 20);
const wide = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cols.filter((x) => x.width > 15).map((x) => `${x.min}-${x.max}`).join());
check(wide === '0-0,3-4', 'Column Width over two ranges: columns A, D and E, as Excel', wide);
await page.keyboard.press('Control+z');

// Merge Cells over two ranges: Excel's question, then each range merged.
await pick('A2:B3,D6:E7');
await ribbon('Merge & Center', async () => {});
const asked = await dialogText();
await okDialog();
const merges = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).merges.map((m) => `${m.r0},${m.c0}:${m.r1},${m.c1}`).join(' '));
check(/Merging cells only keeps the upper-left value/.test(asked) && merges === '1,0:2,1 5,3:6,4', 'Merge & Center over two ranges: Excel’s question, then both merged', `${merges} / ${asked}`);
await page.keyboard.press('Control+z');

// Sort, Filter (Ctrl+Shift+L), Text to Columns and Remove Duplicates: Excel's refusal, nothing changed.
await pick('A2:B4,D6:E8');
await page.click('.oxe-rtab >> text=Data');
const refusals = [];
for (const title of ['Sort A to Z', 'Text to Columns...', 'Remove Duplicates...']) {
  await ribbon(title, async () => {});
  refusals.push(await dialogText());
  await okDialog();
}
await page.focus('.oxe-grid');
await page.keyboard.press('Control+Shift+L');
refusals.push(await dialogText());
await okDialog();
check(refusals.every((t) => t.includes(MULTI) && t.includes('Select a single range and try again.')) && (await values(['A2', 'D6'])).join() === '21,64' && !(await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).autoFilter)),
  "Sort A to Z, Text to Columns, Remove Duplicates, Filter over two ranges: Excel's refusal each time, nothing changed", refusals.join(' / '));
await page.click('.oxe-rtab >> text=Home');
// Insert Table (Ctrl+T): greyed out in Excel; nothing made.
await page.focus('.oxe-grid');
await page.keyboard.press('Control+t');
const opened = await page.isVisible('dialog[open]');
check(!opened && !(await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).tables?.length)), 'Ctrl+T over two ranges: no Create Table box, no table (Excel greys it out)');

// Find within the ranges: Find All finds only the cells in them.
await pick('A2:B3,D6:E7');
await page.focus('.oxe-grid');
await page.keyboard.press('Control+f');
await page.waitForSelector('.oxe-findbar:not([hidden])');
await page.fill('.oxe-findbar input[aria-label="Find what"]', '6');
await page.click('.oxe-findbar button[title="Find All"]');
const found = await page.evaluate(() => [...document.querySelectorAll('.oxe-findall-row span:nth-child(2)')].map((x) => x.textContent).join(' '));
check(found === '$D$6 $E$6', 'Find All with two ranges selected: D6 and E6, in them only (Excel: D6, E6)', found);
await page.click('.oxe-findbar button[title="Find Next (Enter)"]');
const kept = await sel();
check(kept.more?.length === 1 && kept.ar === 5 && [3, 4].includes(kept.ac), 'Find Next: the active cell on a match in the ranges, the two ranges kept selected', JSON.stringify(kept));
await page.keyboard.press('Escape');

// The status bar: Sum over A2:B3 and B3:C4, B3 counted twice, as Excel's.
await pick('A2:B3,B3:C4');
const status = await page.evaluate(() => document.querySelector('.oxe-status')?.textContent || '');
check(/Sum: 256/.test(status) && /Count: 8/.test(status) && /Average: 32/.test(status), 'Status bar over A2:B3 and B3:C4: Sum 256, Count 8, Average 32, as Excel', status);

// Print Selection: each range on a page of its own (Excel's PDF: two pages).
await pick('A2:B3,D6:E7');
const pages = await page.evaluate(() => window.oxe.preparePrint({ selection: true }));
check(pages === 2, 'Print Selection over two ranges: two pages, as Excel', String(pages));

// The Name Box: a name typed refers to both ranges.
await page.fill('.oxe-namebox', 'Both');
await page.press('.oxe-namebox', 'Enter');
const named = await page.evaluate(() => window.oxe.workbook.definedNames.find((d) => d.name === 'Both')?.text || '');
await page.evaluate(() => window.oxe.enter('H1', '=SUM(Both)'));
const total = (await values(['H1']))[0];
check(/\$A\$2:\$B\$3,.*\$D\$6:\$E\$7$/.test(named) && total === 21 + 22 + 31 + 32 + 64 + 65 + 74 + 75, 'Name Box over two ranges: the name refers to both, =SUM(Both) sums both', `${named} ${total}`);

await finish();
