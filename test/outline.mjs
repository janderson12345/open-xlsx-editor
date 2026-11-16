// Browser check of outlines (docs/OUTLINE.md) on a new sheet of test/fixtures/basic.xlsx: Shift+Alt+Right groups the
// rows selected and draws the outline bar at Excel's size (43px with one level, 19px more for each level after:
// excel-outline.ps1), its level buttons and the − at the summary row; the − and + hide and show the group, the level
// buttons show levels, each one step of the history; a range's Group asks Rows or Columns; Ctrl+8 hides and shows the bar; the Settings box's Automatic styles and Apply Styles; Data > Subtotal's box
// inserts Excel's SUBTOTAL rows with their outline, and Remove All takes them away. The result saved and read back.
// Run after `npm run build`: node test/outline.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';
import { Workbook } from '../src/ooxml/workbook.js';

const { page, check, finish, outDir, value } = await editingPage();
await openFixture(page, 'basic.xlsx');
check(await page.evaluate(() => document.compatMode) === 'CSS1Compat', 'standards mode');
await page.evaluate(() => window.oxe.insertSheet({ after: true }));

const select = (ref) => page.evaluate((ref) => { const [a, b] = ref.split(':'); window.oxe.select(a); if (b) { const m = /^([A-Z]+)(\d+)$/.exec(b); window.oxe.command('extendTo', Number(m[2]) - 1, m[1].charCodeAt(0) - 65); } window.oxe.grid.el.focus(); }, ref);
const row = (r) => page.evaluate((r) => window.oxe.workbook.sheet(window.oxe.sheetIndex).rows.get(r) || {}, r);
const bar = () => page.evaluate(() => { const s = getComputedStyle(window.oxe.grid.el); return { w: s.getPropertyValue('--oxe-outline-w').trim(), h: s.getPropertyValue('--oxe-outline-h').trim() }; });
const buttons = (cls) => page.evaluate((cls) => [...document.querySelectorAll(`.${cls}`)].map((b) => `${b.dataset.axis}:${b.textContent}`), cls);
const press = (sel) => page.click(sel);
const box = () => page.locator('dialog[open]');

// Typed: a list of made-up numbers in A1:A8.
for (let r = 1; r <= 8; r++) { await select(`A${r}`); await page.keyboard.type(String(r * 10)); await page.keyboard.press('Enter'); }

// Rows 2 to 4 selected whole and grouped with Shift+Alt+Right.
await select('A2:A4');
await page.evaluate(() => window.oxe.command('selectRows'));
await page.keyboard.press('Shift+Alt+ArrowRight');
check((await row(1)).level === 1 && (await row(3)).level === 1 && !(await row(4)).level, 'Shift+Alt+Right on rows 2:4: outlineLevel 1 on rows 2 to 4');
const one = await bar();
const levels = await buttons('oxe-ol-level');
const toggles = await buttons('oxe-ol-toggle');
check(one.w === '43px' && one.h === '0px' && levels.join() === 'row:1,row:2' && toggles.join() === 'row:−', 'the outline bar: 43px wide, buttons 1 and 2, a − at the summary row', JSON.stringify({ one, levels, toggles }));
await page.screenshot({ path: `${outDir}/outline-bar.png` });

// The − hides the group, the + shows it again; each one step.
await press('.oxe-ol-toggle');
const collapsed = [(await row(1)).hidden, (await row(3)).hidden, (await row(4)).collapsed];
const plus = await buttons('oxe-ol-toggle');
await press('.oxe-ol-toggle');
const opened = [(await row(1)).hidden, (await row(4)).collapsed];
await page.keyboard.press('Control+z');
const undone = (await row(1)).hidden;
await page.keyboard.press('Control+z');
check(collapsed.every(Boolean) && plus.join() === 'row:+' && opened.every((x) => !x) && undone === true && !(await row(1)).hidden,
  'the − hides rows 2:4 and marks row 5 collapsed, the + shows them, Ctrl+Z takes each back', JSON.stringify({ collapsed, plus, opened, undone }));

// A second level: rows 3:3 grouped again; button 2 hides it, 3 shows all.
await select('A3');
await page.evaluate(() => window.oxe.command('selectRows'));
await page.keyboard.press('Shift+Alt+ArrowRight');
const two = await bar();
await press('.oxe-ol-level[data-level="2"]');
const level2 = [(await row(2)).hidden, (await row(1)).hidden, (await row(3)).collapsed];
await press('.oxe-ol-level[data-level="1"]');
const level1 = (await row(1)).hidden && (await row(4)).collapsed;
await press('.oxe-ol-level[data-level="3"]');
check(two.w === '62px' && level2[0] && !level2[1] && level2[2] && level1 && !(await row(1)).hidden && !(await row(2)).hidden,
  'two levels: 62px; button 2 hides row 3 alone, 1 hides rows 2:4, 3 shows all', JSON.stringify({ two, level2, level1 }));

// A range (not whole lines): Group asks Rows or Columns.
await select('B1:C1');
await page.keyboard.press('Shift+Alt+ArrowRight');
await box().getByLabel('Columns').check();
await box().getByRole('button', { name: 'OK' }).click();
const cols = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cols.filter((x) => x.level).map((x) => [x.min, x.max, x.level]));
check(JSON.stringify(cols) === '[[1,2,1]]' && (await bar()).h === '43px', 'B1:C1 grouped by Columns from the box: columns B:C level 1, the bar 43px high over the headings', JSON.stringify(cols));
// Ctrl+8: the outline symbols hidden (the bar gone, the outline kept) and shown again, as Excel's DisplayOutline.
await select('A1');
await page.keyboard.press('Control+8');
const off = [await bar(), await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).outline.symbols), (await row(1)).level];
await page.keyboard.press('Control+8');
check(off[0].w === '0px' && off[0].h === '0px' && off[1] === false && off[2] === 1 && (await bar()).w === '62px', 'Ctrl+8 hides the outline bar (the outline kept), Ctrl+8 again shows it', JSON.stringify(off));
// Data > Outline's Settings: Automatic styles ticked and Apply Styles on A5, the summary of rows 2:4: RowLevel_1, bold.
await select('A5');
await page.click('.oxe-rtab >> text="Data"');
await press('.oxe-toolbar button[title="Outline Settings"]');
await box().getByLabel('Automatic styles').check();
await box().getByRole('button', { name: 'Apply Styles' }).click();
const look = await page.evaluate(() => { const wb = window.oxe.workbook; const sh = wb.sheet(window.oxe.sheetIndex); const xf = wb.styles.cellXfs[sh.cells.get(4).get(0).s]; return { bold: !!wb.styles.fonts[xf.fontId].b, styles: sh.outline.styles }; });
await page.keyboard.press('Control+z');
const undoneStyles = await page.evaluate(() => { const wb = window.oxe.workbook; const sh = wb.sheet(window.oxe.sheetIndex); return { s: sh.cells.get(4).get(0).s || 0, styles: sh.outline.styles }; });
check(look.bold && look.styles && undoneStyles.s === 0 && !undoneStyles.styles, 'Settings: Automatic styles and Apply Styles make A5 bold (RowLevel_1), one step of the history', JSON.stringify({ look, undoneStyles }));
await select('A1');
await page.evaluate(() => window.oxe.command('clearOutline'));
check((await bar()).w === '0px' && (await bar()).h === '0px', 'Clear Outline: no outline, no bar');

// Data > Subtotal over a list: Excel's rows and outline; Remove All takes them away.
const list = [['Region', 'Units'], ['East', '4'], ['East', '2'], ['West', '7']];
for (let r = 0; r < list.length; r++) for (let c = 0; c < 2; c++) { await select(`${'DE'[c]}${r + 1}`); await page.keyboard.type(list[r][c]); await page.keyboard.press('Enter'); }
await select('D2');
await page.click('.oxe-rtab >> text="Data"');
await press('.oxe-toolbar button[title="Subtotal"]');
await box().getByRole('button', { name: 'OK' }).click();
const made = [await value('D4'), await value('E4'), await value('D6'), await value('E7')];
const formula = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(6).get(4).f);
check(made.join() === 'East Total,6,West Total,13' && formula === 'SUBTOTAL(9,E2:E5)' && (await row(1)).level === 2 && (await row(3)).level === 1 && (await bar()).w === '62px',
  'Subtotal: East Total 6, West Total, Grand Total =SUBTOTAL(9,E2:E5) 13, rows at level 2 under level 1', JSON.stringify({ made, formula }));
const bytes = await page.evaluate(async () => Array.from(await window.oxe.save()));
fs.writeFileSync(`${outDir}/outline.xlsx`, Buffer.from(bytes));
const back = Workbook.fromBytes(new Uint8Array(bytes));
const saved = back.sheet(await page.evaluate(() => window.oxe.sheetIndex));
check(saved.rows.get(1)?.level === 2 && saved.rows.get(3)?.level === 1, 'saved and read back with its outline');
await press('.oxe-toolbar button[title="Subtotal"]');
await box().getByRole('button', { name: 'Remove All' }).click();
check(await value('D4') === 'West' && !(await row(1)).level, 'Remove All: the subtotal rows and the outline gone', String(await value('D4')));
await page.keyboard.press('Control+z');
check(await value('D4') === 'East Total', 'Ctrl+Z brings the subtotals back in one step');

await finish();
