// Browser check of Insert > Sparklines and the Sparkline tab (docs/SPARKLINES.md) on test/fixtures/sparklines.xlsx:
// the tab shown only while the active cell holds a sparkline; Create Sparklines filled from the selection, a shape
// Excel refuses refused; a style from the gallery; the type, points, colour, axis, Group, Ungroup and Clear each one step of the history; the
// result saved (to the output folder, for tools/excel-compare/excel-sparkedit-check.ps1) and read back.
// Run after `npm run build`: node test/sparkline-tab.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';
import { Workbook } from '../src/ooxml/workbook.js';

const { page, check, finish, outDir } = await editingPage();
await openFixture(page, 'sparklines.xlsx');
check(await page.evaluate(() => document.compatMode) === 'CSS1Compat', 'standards mode');

/** A cell or a range selected (its first cell active), as a drag would. */
const select = (ref) => page.evaluate((ref) => { const [a, b] = ref.split(':'); window.oxe.select(a); if (b) { const m = /^([A-Z]+)(\d+)$/.exec(b); window.oxe.command('extendTo', Number(m[2]) - 1, m[1].charCodeAt(0) - 65); } }, ref);
const tabShown = () => page.evaluate(() => [...document.querySelectorAll('.oxe-rtab')].some((b) => b.textContent === 'Sparkline' && !b.hidden));
const groups = () => page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).sparklines.map((g) => ({ type: g.type, high: g.high, axis: g.axis, minType: g.minType, min: g.min, series: g.colors.series, cells: g.sparklines.map((s) => `${String.fromCharCode(65 + s.c)}${s.r + 1} ${s.source}`) })));
const groupAt = async (cell) => (await groups()).find((g) => g.cells.some((c) => c.startsWith(`${cell} `)));
const tab = (name) => page.click(`.oxe-rtab >> text="${name}"`);
const press = (title) => page.click(`.oxe-toolbar button[title="${title}"]`);
const menuItem = (text) => page.click(`.oxe-pop button:has(span:text-is("${text}"))`);
const box = () => page.locator('dialog[open]');
const fill = async (label, text) => { await box().getByLabel(label).fill(text); };

await select('A1');
const hidden = !(await tabShown());
await select('O2');
check(hidden && await tabShown(), 'the Sparkline tab: shown with a sparkline in the active cell, not without');

// Insert > Sparklines, an empty cell selected: the box's location filled with it.
await select('Q2');
await tab('Insert');
await press('Insert Line Sparkline');
const location = await box().getByLabel('Location Range:').inputValue();
await fill('Data Range:', 'B2:M2');
await box().getByRole('button', { name: 'OK' }).click();
const made = await groupAt('Q2');
check(location === '$Q$2' && made?.type === 'line' && made.cells[0] === 'Q2 Sparklines!B2:M2' && await page.locator('.oxe-fills svg.oxe-spark').count() > 0,
  'Insert > Line with Q2 selected: Location Range $Q$2, the data typed, a line drawn', JSON.stringify({ location, made }));
await page.keyboard.press('Control+z');
check(!(await groupAt('Q2')), 'one step of the history (Ctrl+Z takes it away)');

// Data selected: the box's data filled with it; a location Excel refuses refused in the box.
await select('B2:M3');
await press('Insert Column Sparkline');
const dataField = await box().getByLabel('Data Range:').inputValue();
await fill('Location Range:', 'Q2:Q4');
await box().getByRole('button', { name: 'OK' }).click();
const refused = await box().getByRole('alert').textContent();
await fill('Location Range:', 'Q2:Q3');
await box().getByRole('button', { name: 'OK' }).click();
const two = await groupAt('Q3');
check(dataField === 'B2:M3' && /not valid/.test(refused) && two?.type === 'column' && two.cells.join() === 'Q2 Sparklines!B2:M2,Q3 Sparklines!B3:M3',
  'Insert > Column with B2:M3 selected: the data filled in, Q2:Q4 refused, Q2:Q3 one group of two by rows', JSON.stringify({ dataField, refused, two }));

// The Sparkline tab on Q2.
await select('Q2');
await tab('Sparkline');
await page.screenshot({ path: `${outDir}/sparkline-tab.png` });
await press('Sparkline Styles');
const shown = await page.locator('.oxe-pop button.oxe-spark-style').count();
await page.screenshot({ path: `${outDir}/sparkline-styles.png` });
await page.click('.oxe-pop button[title="Sparkline Style Colorful #4"]');
const styled = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).sparklines.find((x) => x.sparklines.some((s) => s.r === 1 && s.c === 16)).colors);
check(shown === 36 && styled.series.rgb === 'FF00B050' && styled.markers.rgb === 'FF0070C0' && styled.low.rgb === 'FFFF0000',
  'the style gallery: 36 styles; Colorful #4 gives the group its colours', JSON.stringify({ shown, series: styled.series, markers: styled.markers }));
await press('Win/Loss');
await press('High Point');
await press('Sparkline Color');
await page.click('.oxe-pop button.oxe-swatch[title="Red"]');
let g = await groupAt('Q2');
check(g.type === 'stacked' && g.high && g.series.rgb === 'FFFF0000', 'Win/Loss, High Point and a standard colour', JSON.stringify(g));
await press('Sparkline Axis');
await menuItem('Show Axis');
await press('Sparkline Axis');
await page.locator('.oxe-pop button:has(span:text-is("Custom Value..."))').first().click();
await box().locator('input').fill('-3');
await box().getByRole('button', { name: 'OK' }).click();
g = await groupAt('Q2');
check(g.axis && g.minType === 'custom' && g.min === -3, 'Axis: Show Axis, the minimum a custom -3', JSON.stringify(g));
await press('Ungroup Selected Sparklines');
const apart = (await groupAt('Q2')) !== (await groupAt('Q3')) && (await groupAt('Q2')).cells.length === 1;
await select('Q2:Q3');
await press('Group Selected Sparklines');
const together = (await groupAt('Q2')).cells.length === 2;
check(apart && together, 'Ungroup puts Q2 in a group of its own; Group with Q2:Q3 selected joins them again');
await select('Q3');
await press('Clear Sparklines');
await menuItem('Clear Selected Sparklines');
check(!(await groupAt('Q3')) && !!(await groupAt('Q2')), 'Clear Selected Sparklines takes Q3 alone');

// Saved and read back: what the tab wrote.
const bytes = new Uint8Array(await page.evaluate(() => [...window.oxe.save()]));
fs.mkdirSync(`${outDir}/sparklines`, { recursive: true });
fs.writeFileSync(`${outDir}/sparklines/sparkline-edits.xlsx`, bytes);
const back = Workbook.fromBytes(bytes).sheet(0).sparklines.find((x) => x.sparklines.some((s) => s.r === 1 && s.c === 16));
check(back?.type === 'stacked' && back.high && back.axis && back.minType === 'custom' && back.min === -3 && back.colors.series.rgb === 'FFFF0000' && back.sparklines.length === 1,
  'saved and read back as the tab left it', JSON.stringify(back));
await finish();
