// Browser check of tables (docs/TABLES.md) on test/fixtures/tables.xlsx: drawn in their styles (a header's fill and
// font, a stripe); Ctrl+T's Create Table box filled from the data round the active cell, the table made in the default
// style; the Table Design tab shown only in a table, its options, the total row and its drop-down, the style gallery,
// Table Name and Convert to Range, each one step of the history; a table's own filter button; Home > Format as Table;
// Home > Cell Styles; the result saved (to the output folder) and read back.
// Run after `npm run build`: node test/table-tab.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';
import { Workbook } from '../src/ooxml/workbook.js';
import { readTables } from '../src/ooxml/tables.js';

const { page, check, finish, outDir, ribbon } = await editingPage();
await openFixture(page, 'tables.xlsx');
check(await page.evaluate(() => document.compatMode) === 'CSS1Compat', 'standards mode');

const select = (ref) => page.evaluate((ref) => { window.oxe.select(ref); window.oxe.grid.el.focus({ preventScroll: true }); }, ref);
const tabShown = () => page.evaluate(() => [...document.querySelectorAll('.oxe-rtab')].some((b) => b.textContent === 'Table Design' && !b.hidden));
const tab = (name) => page.click(`.oxe-rtab >> text="${name}"`);
// (Through the kit's ribbon(), which opens More when the ribbon has folded the button's group away, as on a narrower
// screen or GitHub's Ubuntu runner with its wider fonts.)
const press = (title) => ribbon(title);
const box = () => page.locator('dialog[open]');
const tables = () => page.evaluate(() => window.oxe.workbook.pkg.files && Object.keys(window.oxe.workbook.pkg.files).filter((k) => k.startsWith('xl/tables/')).length);
const part = (n) => page.evaluate((n) => window.oxe.workbook.pkg.text(`xl/tables/table${n}.xml`), n);
const cellCss = (ref) => page.evaluate((ref) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  const id = document.querySelector(`[id$="-${Number(m[2]) - 1}-${m[1].charCodeAt(0) - 65}"]`);
  return id ? { color: getComputedStyle(id).color, weight: getComputedStyle(id).fontWeight } : null;
}, ref);

// Drawn: Table2 (TableStyleLight2, every option on) and Table1 (Light1, banded rows).
await select('A1');
const fills = await page.locator('.oxe-fill.oxe-tbl-fill').count();
const edges = await page.locator('.oxe-edge.oxe-tbl-edge').count();
const head = await cellCss('H2');
check(fills > 10 && edges > 10 && head?.weight === '700', 'tables drawn in their styles: fills, lines, a bold header', JSON.stringify({ fills, edges, head }));

// The Table Design tab: only in a table.
const away = !(await tabShown());
await select('C4');
check(away && await tabShown(), 'the Table Design tab: shown with the active cell in a table, not outside');

// Ctrl+T on new data: Create Table filled from the data round the active cell.
const before = await tables();
await page.evaluate(() => { const o = window.oxe; [['L3', 'Name'], ['M3', 'Score'], ['L4', 'ann'], ['M4', 5], ['L5', 'bob'], ['M5', 7]].forEach(([r, v]) => o.setCell(r, v)); o.select('L4'); o.grid.el.focus(); });
await page.keyboard.press('Control+t');
const range = await box().locator('input[type=text]').inputValue();
const headers = await box().locator('input[type=checkbox]').isChecked();
await box().getByRole('button', { name: 'OK' }).click();
const made = await part(before + 1);
check(range === '=$L$3:$M$5' && headers && /ref="L3:M5"/.test(made) && /name="TableStyleMedium2"/.test(made) && /<autoFilter ref="L3:M5"\/>/.test(made),
  'Ctrl+T: Create Table with =$L$3:$M$5 and My table has headers; the table made as Excel makes it', JSON.stringify({ range, headers, made }));
await page.keyboard.press('Control+z');
check(await tables() === before, 'one step of the history (Ctrl+Z takes it away)');
await page.keyboard.press('Control+y');
check(await tables() === before + 1, 'and Ctrl+Y puts it back');

// The tab's options on the new table.
await select('L4');
await tab('Table Design');
await page.screenshot({ path: `${outDir}/table-tab.png` });
await press('Total Row');
let xml = await part(before + 1);
const total = await page.evaluate(() => { const c = window.oxe.workbook.sheet(0).cells.get(5); return [c?.get(11)?.v, c?.get(12)?.f]; });
check(/totalsRowCount="1"/.test(xml) && total[0] === 'Total' && total[1] === 'SUBTOTAL(109,Table63[Score])', 'Total Row: "Total" and a SUM of the numbers', JSON.stringify(total));
await select('M6');
await page.locator('.oxe-dv-btn').click();
await page.locator('dialog[open], .oxe-pick').getByText('Average', { exact: true }).click();
const avg = await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(5)?.get(12)?.f);
check(avg === 'SUBTOTAL(101,Table63[Score])', "the total cell's drop-down: Average", avg);
await select('L4');
await press('Banded Columns');
await press('First Column');
xml = await part(before + 1);
check(/showFirstColumn="1"/.test(xml) && /showColumnStripes="1"/.test(xml), 'Banded Columns and First Column', xml);
await press('Table Styles');
const swatches = await page.locator('.oxe-pop button.oxe-tstyle').count();
await page.screenshot({ path: `${outDir}/table-styles.png` });
await page.click('.oxe-pop button[title="Table Style Dark 3"]');
xml = await part(before + 1);
check(swatches === 61 && /name="TableStyleDark3"/.test(xml), 'the style gallery: 60 built-in styles and the workbook\'s own; Dark 3 picked', JSON.stringify({ swatches }));
const name = page.locator('.oxe-table-name');
await name.fill('1st');
await name.press('Enter');
const refused = await page.locator('dialog[open]').textContent();
await page.locator('dialog[open]').getByRole('button', { name: 'OK' }).click();
await name.fill('Scores');
await name.press('Enter');
xml = await part(before + 1);
check(/not valid/.test(refused) && /displayName="Scores"/.test(xml), "Table Name: 1st refused in Excel's words, Scores taken", refused);

// A table's own filter button.
await page.locator('.oxe-filter-btn[data-table="Scores"][data-col="12"]').click();
const listShown = await page.locator('.oxe-pop, dialog[open]').count();
await page.keyboard.press('Escape');
const filtered = await page.evaluate(() => window.oxe.command('filterBy', 12, { values: ['5'], blank: false }, 'Scores'));
const hidden = await page.evaluate(() => !!window.oxe.workbook.sheet(0).rows.get(4)?.hidden);
xml = await part(before + 1);
check(listShown > 0 && filtered && hidden && /<filterColumn colId="1"><filters><filter val="5"\/>/.test(xml), "a table's filter button: its list, and the filter kept in the table's part", xml);
await page.keyboard.press('Control+z');

// Saved and read back.
const bytes = new Uint8Array(await page.evaluate(() => [...window.oxe.save()]));
fs.mkdirSync(`${outDir}/tables`, { recursive: true });
fs.writeFileSync(`${outDir}/tables/table-edits.xlsx`, bytes);
const back = readTables(Workbook.fromBytes(bytes)).find((t) => t.name === 'Scores');
check(back && back.totals === 1 && back.style.name === 'TableStyleDark3' && back.style.first && back.style.cols && back.totalsRow[1].fn === 'average',
  'saved and read back as the tab left it', JSON.stringify(back && { totals: back.totals, style: back.style, totalsRow: back.totalsRow }));

// Convert to Range: asked first, the look kept as formats.
await select('L4');
await press('Convert to Range');
await box().getByRole('button', { name: 'Yes' }).click();
const left = await tables();
const kept = await page.evaluate(() => (window.oxe.workbook.sheet(0).cells.get(2)?.get(11)?.s || 0) > 0);
check(left === before && kept, 'Convert to Range: the table gone, its header keeping a format of its own', JSON.stringify({ left, kept }));

// Home > Format as Table on plain data: Create Table, then the style picked.
await page.evaluate(() => { const o = window.oxe; [['L20', 'Fruit'], ['M20', 'Qty'], ['L21', 'fig'], ['M21', 3], ['L22', 'kiwi'], ['M22', 4]].forEach(([r, v]) => o.setCell(r, v)); o.select('L21'); });
await tab('Home');
// (Past the window's width the Styles group is in the ribbon's overflow: its button pressed as the overflow would.)
await page.evaluate(() => document.querySelector('.oxe-toolbar button[title="Format as Table"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true })));
await page.click('.oxe-pop button[title="Table Style Light 9"]');
await box().getByRole('button', { name: 'OK' }).click();
const fmt = await page.evaluate(() => Object.keys(window.oxe.workbook.pkg.files).filter((k) => k.startsWith('xl/tables/')).map((k) => window.oxe.workbook.pkg.text(k)).find((x) => /ref="L20:M22"/.test(x)));
check(/name="TableStyleLight9"/.test(fmt || ''), 'Home > Format as Table: Create Table, the table in Light 9', fmt || await page.evaluate(() => Object.keys(window.oxe.workbook.pkg.files).filter((k) => k.startsWith('xl/tables/')).map((k) => /ref="[^"]*"/.exec(window.oxe.workbook.pkg.text(k))[0]).slice(-3).join() + ' ' + document.querySelectorAll('dialog[open]').length));
// Home > Cell Styles: the gallery under Excel's headings, Good given to a cell as Excel gives it.
await select('N3');
await page.evaluate(() => document.querySelector('.oxe-toolbar button[title="Cell Styles"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true })));
const chips = await page.locator('.oxe-pop button.oxe-cstyle').count();
await page.screenshot({ path: `${outDir}/cell-styles.png` });
await page.click('.oxe-pop button.oxe-cstyle[title="Good"]');
const good = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const xf = wb.styles.cellXfs[wb.sheet(0).cells.get(2).get(13).s];
  const fill = wb.styles.fills[xf.fillId];
  return { fill: fill.fg?.rgb, font: wb.styles.fonts[xf.fontId].color?.rgb, style: /<cellStyle name="Good" xfId="\d+" builtinId="26"\/>/.test(new XMLSerializer().serializeToString(wb.stylesDoc)) };
});
check(chips === 47 && good.fill === 'FFC6EFCE' && good.font === 'FF006100' && good.style, 'Home > Cell Styles: 47 styles; Good gives the cell its fill, font and the cellStyle Excel writes', JSON.stringify({ chips, good }));
await page.keyboard.press('Control+z');
const undone = await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(2)?.get(13)?.s || 0);
check(undone === 0, 'one step of the history', String(undone));
await finish();
