// Browser check of editing, continued: a workbook saved with a password, PivotTables (refreshed and made), the Data
// tab, conditional formatting, several ranges selected, Go To Special, sheets undone, rich text drawn, and a column's own format.
// Run after `npm run build`: node test/editing-more.mjs [outDir]
import fs from 'node:fs';
import { encryptWorkbook } from '../src/ooxml/encryption.js';
import { pivotBook } from './pivots.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, shown, sel, clickCell, pointAt, ribbon, menu } = await editingPage();

// A workbook saved with a password (made here with a test password): Excel's Password box; a wrong one gets Excel's
// words and nothing opens; the right one opens it, and it saves encrypted again.
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/encrypted.xlsx', encryptWorkbook(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')), 'Test-pass 1', null, { spins: 1000 }));
const askPassword = async (pw) => {
  await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
  await page.setInputFiles('input[type=file]', 'test/out/encrypted.xlsx');
  await page.waitForSelector('dialog.oxe-dialog input[type=password]');
  const text = await page.textContent('dialog.oxe-dialog');
  await page.fill('dialog.oxe-dialog input[type=password]', pw);
  await page.keyboard.press('Enter');
  return text;
};
const asked = await askPassword('not it');
await page.waitForSelector('dialog.oxe-dialog >> text=The password you supplied is not correct');
await page.keyboard.press('Enter');
await page.waitForFunction(() => /was not opened/.test(document.getElementById('status').textContent));
check(/'encrypted\.xlsx' is protected\./.test(asked) && await page.evaluate(() => window.oxe.fileName) !== 'encrypted.xlsx', 'a password-protected workbook: Excel’s Password box; a wrong password: Excel’s words, not opened', asked);
await askPassword('Test-pass 1');
await page.waitForFunction(() => window.oxe.fileName === 'encrypted.xlsx');
const saved = await page.evaluate(() => ({ encrypted: window.oxe.encrypted, sheets: window.oxe.workbook.sheets.length, head: [...window.oxe.save().slice(0, 4)] }));
check(saved.encrypted && saved.sheets === 3 && saved.head.join() === '208,207,17,224', 'the right password: opened, and saved encrypted again (a compound file)', JSON.stringify(saved));

// A PivotTable (Excel's, of made-up sales: test/pivot-excel.json): its source changed, Alt+F5 on it refreshes it as
// Excel does; Ctrl+Z takes the refresh back; Data > Refresh All does it again.
fs.writeFileSync('test/out/pivot.xlsx', pivotBook(0).wb.toBytes());
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/out/pivot.xlsx', 'pivot.xlsx');
await page.evaluate(() => { window.oxe.showSheet('Src'); window.oxe.enter('D2', '100'); window.oxe.enter('A2', 'Central'); window.oxe.showSheet('Pivot'); window.oxe.select('A4'); });
await page.focus('.oxe-grid');
await page.keyboard.press('Alt+F5');
const pivotShown = async () => [await shown(5, 1), await shown(6, 0), await shown(6, 1), await shown(7, 1)].join();
check(await pivotShown() === '150,Central,100,434', 'Alt+F5 in the PivotTable: refreshed as Excel refreshes it (South 150, Central added after, Grand Total 434)', await pivotShown());
await page.keyboard.press('Control+z');
check(await shown(5, 1) === '168' && await shown(6, 0) === 'Grand Total', 'Ctrl+Z: the refresh taken back', await pivotShown());
await page.click('.oxe-rtab >> text=Data');
await page.click('.oxe-toolbar button[title="Refresh All (Ctrl+Alt+F5)"]');
check(await pivotShown() === '150,Central,100,434', 'Data > Refresh All: refreshed again', await pivotShown());
// Data > Sort & Filter: Filter on the source sheet turns the filter on over its table of data.
await page.evaluate(() => { window.oxe.showSheet('Src'); window.oxe.select('B3'); });
await page.click('.oxe-toolbar button[title="Filter (Ctrl+Shift+L)"]');
check(await page.evaluate(() => { const f = window.oxe.workbook.sheet(window.oxe.sheetIndex).autoFilter; return f && `${f.r0},${f.c0},${f.r1},${f.c1}`; }) === '0,0,18,4', 'Data > Filter: the filter on over A1:E19', await page.evaluate(() => JSON.stringify(window.oxe.workbook.sheet(window.oxe.sheetIndex).autoFilter)));
// Data > Data Validation... on G2:G4: Allow List, Source Yes,No, OK: the rule on the three cells; Maybe refused.
await page.evaluate(() => { window.oxe.select('G2'); window.oxe.command('extendTo', 3, 6); });
await page.click('.oxe-toolbar button[title="Data Validation..."]');
await page.waitForSelector('dialog.oxe-dv-box');
await page.selectOption('dialog.oxe-dv-box select >> nth=0', 'list');
await page.fill('dialog.oxe-dv-box .oxe-dv-panel:not([hidden]) input[type=text] >> nth=0', 'Yes,No');
await page.click('dialog.oxe-dv-box >> text=OK');
const dv = await page.evaluate(() => { const v = window.oxe.workbook.sheet(window.oxe.sheetIndex).validations; return v.map((x) => `${x.type}:${x.ranges.map((r) => `${r.r0},${r.c0},${r.r1},${r.c1}`).join(' ')}:${x.formulas[0]}`).join(';'); });
const maybe = await page.evaluate(() => window.oxe.enter('G3', 'Maybe'));
check(dv === 'list:1,6,3,6:"Yes,No"' && maybe?.invalid?.style === 'stop', 'Data Validation box: a list rule on G2:G4; Maybe refused (Stop)', `${dv} ${JSON.stringify(maybe)}`);
// Home > Conditional Formatting > Greater Than...: 20 on the Units (D2:D19), light red: the cells over 20 drawn red.
await page.click('.oxe-rtab >> text=Home');
await page.evaluate(() => { window.oxe.select('D2'); window.oxe.command('extendTo', 18, 3); });
// (The box answered inside the ribbon helper: the More panel closes after it.)
await ribbon('Conditional Formatting', async () => {
  await page.click('.oxe-pop >> text=Greater Than...');
  await page.waitForSelector('dialog.oxe-cf-box');
  await page.fill('dialog.oxe-cf-box input', '20');
  await page.click('dialog.oxe-cf-box >> text=OK');
});
const cfLooks = await page.evaluate(() => { const looks = []; const sh = window.oxe.workbook.sheet(window.oxe.sheetIndex); for (let r = 1; r <= 18; r++) { const v = sh.cells.get(r)?.get(3)?.v; const el = document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="5"]`); looks.push(`${v}:${el?.classList.contains('oxe-cf-color') ? getComputedStyle(el).color : '-'}`); } return looks; });
check(cfLooks.length === 18 && cfLooks.every((x) => { const [v, color] = x.split(':'); return Number(v) > 20 ? color === 'rgb(156, 0, 6)' : color === '-'; }), 'Conditional Formatting > Greater Than 20: the cells over 20 in dark red, the rest as they were', cfLooks.join(' '));
await page.click('.oxe-rtab >> text=Data');
await page.click('.oxe-rtab >> text=Home');

// Ctrl+click: a second range, both shaded; Bold goes to both, Copy refuses them, as Excel does.
await clickCell(1, 0); // A2
await page.keyboard.down('Control');
await clickCell(3, 2); // C4
await page.keyboard.up('Control');
const two = await sel();
const shaded = await page.evaluate(() => document.querySelectorAll('.oxe-sel-range').length);
check(two.more?.length === 1 && two.ar === 3 && two.ac === 2 && shaded === 2, 'Ctrl+click: A2 and C4 selected, both shaded, C4 active', `${JSON.stringify(two)} ${shaded}`);
const boldOf = (r, c) => page.evaluate(([r, c]) => { const wb = window.oxe.workbook; const s = wb.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c)?.s || 0; return !!wb.styles.fonts[wb.styles.cellXfs[s].fontId]?.b; }, [r, c]);
await page.keyboard.press('Control+b');
check(await boldOf(1, 0) && await boldOf(3, 2), 'Bold (Ctrl+B) on two ranges: both cells bold');
await page.keyboard.press('Control+c');
const manyRefusal = await page.textContent('dialog.oxe-dialog');
check(/won't work on multiple selections/.test(manyRefusal || ''), "Copy over two ranges: Excel's refusal", manyRefusal);
await page.click('dialog.oxe-dialog >> text=OK');
await page.keyboard.press('Control+z');
// Home > Find & Select > Go To Special...: Constants, Numbers only, from A1: every number typed in, and only those.
await clickCell(0, 0);
await ribbon('Find & Select', async () => {
  await page.click('.oxe-pop >> text=Go To Special...');
  await page.waitForSelector('dialog.oxe-special-box');
  await page.check('dialog.oxe-special-box >> text=Constants');
  for (const t of ['Text', 'Logicals', 'Errors']) await page.uncheck(`dialog.oxe-special-box label:text-is("${t}")`);
  await page.click('dialog.oxe-special-box >> text=OK');
});
const special = await page.evaluate(() => {
  const s = window.oxe.selection;
  const sh = window.oxe.workbook.sheet(window.oxe.sheetIndex);
  const areas = [...(s.more || []), s];
  const inSel = (r, c) => areas.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1);
  let wrong = 0; let numbers = 0;
  for (const [r, line] of sh.cells) for (const [c, cell] of line) {
    const want = cell.f == null && cell.t === 'n' && cell.v != null;
    if (want) numbers++;
    if (want !== inSel(r, c) && !sh.merges.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1)) wrong++;
  }
  return { numbers, wrong, ranges: areas.length };
});
check(special.numbers > 0 && special.wrong === 0, 'Go To Special > Constants, Numbers: the number constants selected, nothing else', JSON.stringify(special));

// Insert > PivotTable: a small table on a sheet of its own, Region to Rows and Units to Values; the pivot on a new
// sheet at A3, its items sorted, as Excel makes one.
await page.evaluate(() => {
  window.oxe.command('insertSheet');
  const data = [['Region', 'Units'], ['South', 5], ['North', 7], ['South', 3], ['East', 4], ['North', 1]];
  data.forEach((row, r) => row.forEach((v, c) => window.oxe.command('setValue', r, c, v)));
  window.oxe.select('A1');
});
const sourceSheet = await page.evaluate(() => window.oxe.workbook.sheets[window.oxe.sheetIndex].name);
await page.click('.oxe-rtab >> text=Insert');
await ribbon('PivotTable', async () => {
  await page.waitForSelector('dialog.oxe-pivot-box');
  const guess = await page.inputValue('dialog.oxe-pivot-box input[type="text"]');
  check(guess === `${sourceSheet}!$A$1:$B$6`, 'Insert > PivotTable: the table round the active cell offered as the source', guess);
  await page.click('dialog.oxe-pivot-box >> button:text-is("OK")');
  await page.waitForSelector('dialog.oxe-pivot-fields');
  await page.selectOption('dialog.oxe-pivot-fields select[aria-label="Region: area"]', 'rows');
  await page.selectOption('dialog.oxe-pivot-fields select[aria-label="Units: area"]', 'values');
  await page.click('dialog.oxe-pivot-fields >> button:text-is("OK")');
});
const madePivot = await page.evaluate(() => {
  const sh = window.oxe.workbook.sheet(window.oxe.sheetIndex);
  const out = [];
  for (let r = 2; r <= 6; r++) out.push([0, 1].map((c) => sh.cells.get(r)?.get(c)?.v ?? '').join('='));
  return { sheet: window.oxe.workbook.sheets[window.oxe.sheetIndex].name, cells: out.join(' ') };
});
check(madePivot.sheet !== sourceSheet && madePivot.cells === 'Row Labels=Sum of Units East=4 North=8 South=8 Grand Total=20', 'PivotTable made on a new sheet: Row Labels, East, North, South, Grand Total 20', JSON.stringify(madePivot));
await page.click('.oxe-rtab >> text=Home');
// Ctrl+Z takes the PivotTable's new sheet back; Ctrl+Y makes it again.
const tabNames = () => page.evaluate(() => [...document.querySelectorAll('.oxe-tab')].map((x) => x.textContent).join(','));
const withPivot = await tabNames();
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const pivotUndone = await tabNames();
await page.keyboard.press('Control+y');
check(!pivotUndone.includes(madePivot.sheet) && await tabNames() === withPivot, 'Ctrl+Z: the PivotTable\'s sheet taken back; Ctrl+Y: made again', `${pivotUndone} / ${await tabNames()}`);
// Sheets renamed and deleted: one undo step each, the sheet back with its cells and its tab.
await page.evaluate((name) => window.oxe.showSheet(name), sourceSheet);
await page.evaluate((name) => window.oxe.renameSheet(name, 'Regions'), sourceSheet);
await page.keyboard.press('Control+z');
const renameUndone = await tabNames();
await page.evaluate(() => window.oxe.command('deleteSheet', window.oxe.sheetIndex));
const deleted = await tabNames();
await page.keyboard.press('Control+z');
const back = await page.evaluate(() => ({ name: window.oxe.workbook.sheets[window.oxe.sheetIndex].name, a2: window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(1)?.get(0)?.v }));
check(renameUndone.includes(sourceSheet) && !renameUndone.includes('Regions') && !deleted.includes(sourceSheet) && back.name === sourceSheet && back.a2 === 'South' && (await tabNames()).includes(sourceSheet),
  'Ctrl+Z after Rename and after Delete: the sheet as it was, its cells and tab back', JSON.stringify({ renameUndone, deleted, back }));

// A dynamic array formula typed in the grid: =SORT(B2:B6) in H2 spills down, drawn; a value typed in its way makes it
// #SPILL!, and Ctrl+Z lets it spill again.
await page.evaluate(() => window.oxe.select('H2'));
await page.keyboard.type('=SORT(B2:B6)');
await page.keyboard.press('Enter');
const spilled = async () => { const out = []; for (let r = 1; r <= 5; r++) out.push(await shown(r, 7)); return out.join(','); };
const sorted = await spilled();
check(sorted === '1,3,4,5,7', 'a dynamic array formula typed in H2: =SORT(B2:B6) spills down H2:H6, drawn', sorted);
await page.evaluate(() => window.oxe.select('H4'));
await page.keyboard.type('x');
await page.keyboard.press('Enter');
const blocked = await shown(1, 7);
await page.keyboard.press('Control+z');
check(blocked === '#SPILL!' && await spilled() === sorted, 'a value typed in its way: #SPILL!; Ctrl+Z: it spills again', `${blocked} / ${await spilled()}`);

// Show Field List on the PivotTable made above (right-click one of its cells): Region moved from Rows to Columns; the
// pivot laid out again across, as Excel's field list lays it out.
await page.evaluate((name) => { window.oxe.showSheet(name); window.oxe.select('A4'); }, madePivot.sheet);
const a4 = await pointAt(3, 0);
await page.mouse.click(a4.x, a4.y, { button: 'right' });
await page.click('.oxe-menu [role="menuitem"]:text-is("Show Field List")');
await page.waitForSelector('dialog.oxe-pivot-fields');
const startArea = await page.inputValue('dialog.oxe-pivot-fields select[aria-label="Region: area"]');
await page.selectOption('dialog.oxe-pivot-fields select[aria-label="Region: area"]', 'cols');
await page.click('dialog.oxe-pivot-fields >> button:text-is("OK")');
const across = await page.evaluate(() => { const sh = window.oxe.workbook.sheet(window.oxe.sheetIndex); return [2, 3, 4].map((r) => [0, 1, 2, 3, 4].map((c) => sh.cells.get(r)?.get(c)?.v ?? '').join('|')).join(' / '); });
check(startArea === 'rows' && across === '|Column Labels||| / |East|North|South|Grand Total / Sum of Units|4|8|8|20', 'Show Field List: Region from Rows to Columns, the pivot laid out across', `${startArea} ${across}`);

// The Format Cells box (Ctrl+1) on B2:C3 of a new sheet: Alignment 45 degrees, Font bold, Border Outline (thin),
// Fill a solid colour; OK gives the four cells those, the outline on the range's edge only. Then Format Painter: B2's
// formats picked up and painted on E5 with one click.
await page.evaluate(() => window.oxe.command('insertSheet'));
await page.evaluate(() => { window.oxe.enter('B2', 'abc'); window.oxe.select('B2'); window.oxe.command('extendTo', 2, 2); });
await page.focus('.oxe-grid');
await page.keyboard.press('Control+1');
await page.waitForSelector('dialog.oxe-format-box');
await page.click('dialog.oxe-format-box [role=tab]:text-is("Alignment")');
await page.fill('dialog.oxe-format-box .oxe-fc-panel:not([hidden]) input[type=number] >> nth=1', '45');
await page.click('dialog.oxe-format-box [role=tab]:text-is("Font")');
await page.selectOption('dialog.oxe-format-box .oxe-fc-panel:not([hidden]) select >> nth=0', 'b');
await page.click('dialog.oxe-format-box [role=tab]:text-is("Border")');
await page.click('dialog.oxe-format-box .oxe-fc-panel:not([hidden]) button:text-is("Outline")');
await page.click('dialog.oxe-format-box [role=tab]:text-is("Fill")');
await page.selectOption('dialog.oxe-format-box .oxe-fc-panel:not([hidden]) select >> nth=0', 'FFFFFF00');
await page.click('dialog.oxe-format-box >> button:text-is("OK")');
const boxed = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const sh = wb.sheet(window.oxe.sheetIndex);
  return [[1, 1], [1, 2], [2, 1], [2, 2]].map(([r, c]) => {
    const xf = wb.styles.cellXfs[sh.cells.get(r)?.get(c)?.s || 0];
    const b = wb.styles.borders[xf.borderId];
    return [xf.align.rotation, wb.styles.fonts[xf.fontId].b ? 'b' : '-', wb.styles.fills[xf.fillId].fg?.rgb, ['left', 'top', 'right', 'bottom'].filter((k) => b[k].style).join('+')].join(' ');
  }).join(' / ');
});
check(boxed === '45 b FFFFFF00 left+top / 45 b FFFFFF00 top+right / 45 b FFFFFF00 left+bottom / 45 b FFFFFF00 right+bottom',
  'Format Cells on B2:C3: rotation, bold and fill on every cell, the outline on the edge only', boxed);
const turned = await page.evaluate(() => getComputedStyle(document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="3"] > .oxe-tx')).transform);
check(turned !== 'none', 'B2 drawn turned', turned);
await page.evaluate(() => window.oxe.select('B2'));
await page.click('.oxe-rtab >> text=Home');
await ribbon('Format Painter: press twice to keep it on');
await clickCell(4, 4);
const painted = await page.evaluate(() => { const sh = window.oxe.workbook.sheet(window.oxe.sheetIndex); return [sh.cells.get(1).get(1).s, sh.cells.get(4)?.get(4)?.s, document.querySelector('.oxe-root').classList.contains('oxe-painting')]; });
check(painted[0] === painted[1] && !painted[2], 'Format Painter: the format of B2 on E5 with one click, then off', JSON.stringify(painted));
await page.keyboard.press('Control+z');
const unpainted = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(4)?.get(4)?.s || 0);
check(unpainted === 0, 'Ctrl+Z takes the painted format away', String(unpainted));

// Insert > Charts on a new sheet's table: a clustered column chart beside it, drawn; picked with a click and dragged
// a column right; Edit Chart (Enter) gives it a title; Delete deletes it and Ctrl+Z brings it back.
await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90]].forEach((row, r) => row.forEach((v, c) => o.enter(String.fromCharCode(65 + c) + (r + 1), String(v))));
  o.select('A1');
});
await page.click('.oxe-rtab >> text=Insert');
await menu('Insert Column or Bar Chart', 'Clustered Column');
await page.waitForSelector('.oxe-chart');
const chartsNow = () => page.evaluate(() => window.oxe.workbook && [...document.querySelectorAll('.oxe-chart')].map((d) => d.dataset.oxeChart).join(','));
const anchorNow = () => page.evaluate(() => { const s = document.querySelector('.oxe-chart'); return s ? s.style.getPropertyValue('--oxe-x') : null; });
check(await chartsNow() === '0', 'Insert > Charts > Clustered Column: the chart drawn beside the table', await chartsNow());
const chartX0 = await anchorNow();
const chartBox0 = await page.locator('.oxe-chart').boundingBox();
await page.mouse.move(chartBox0.x + 100, chartBox0.y + 100);
await page.mouse.down();
await page.mouse.move(chartBox0.x + 140, chartBox0.y + 100, { steps: 4 });
await page.mouse.move(chartBox0.x + 170, chartBox0.y + 100, { steps: 4 });
await page.mouse.up();
const chartX1 = await anchorNow();
check(parseFloat(chartX1) > parseFloat(chartX0) + 50, 'the chart dragged right: its anchor moved', `${chartX0} -> ${chartX1}`);
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-chart-box');
await page.selectOption('dialog.oxe-chart-box select >> nth=1', 'text');
await page.fill('dialog.oxe-chart-box input[type=text]', 'Sales by region');
await page.click('dialog.oxe-chart-box >> button:text-is("OK")');
const titled = await page.evaluate(() => [...document.querySelectorAll('.oxe-chart svg text')].some((t) => t.textContent === 'Sales by region'));
check(titled, 'Edit Chart: the title drawn');
await page.keyboard.press('Delete');
const chartGone = await chartsNow();
await page.keyboard.press('Control+z');
await page.waitForSelector('.oxe-chart');
check(chartGone === '' && await chartsNow() === '0', 'Delete deletes the picked chart; Ctrl+Z brings it back', `${chartGone} / ${await chartsNow()}`);

// A comment thread (Ctrl+Alt+M) on B2 of a new sheet: posted, replied to, the reply edited, then deleted; the thread
// read back from the workbook at each step, and Ctrl+Z undoing the deletion.
await page.evaluate(() => { window.oxe.command('insertSheet'); window.oxe.select('B2'); });
await page.focus('.oxe-grid');
const threadNow = () => page.evaluate(() => {
  const wb = window.oxe.workbook;
  const path = wb.sheets[window.oxe.sheetIndex].path;
  const rel = wb.pkg.readRels(path).find((r) => /threadedComment$/.test(r.type));
  const xml = rel ? wb.pkg.text(rel.target) : '';
  return [...xml.matchAll(/<text>([\s\S]*?)<\/text>/g)].map((m) => m[1]).join(' | ');
});
const box = 'dialog.oxe-thread-box';
await page.keyboard.press('Control+Alt+m');
await page.waitForSelector(box);
await page.fill(`${box} textarea`, 'Check these totals');
await page.click(`${box} >> button:text-is("Post")`);
await page.waitForSelector(`${box} .oxe-thread-item`);
await page.fill(`${box} > textarea`, 'Done');
await page.click(`${box} >> button:text-is("Reply")`);
await page.waitForFunction((b) => document.querySelectorAll(`${b} .oxe-thread-item`).length === 2, box);
check(await threadNow() === 'Check these totals | Done', 'New Comment, then a reply: the thread written', await threadNow());
await page.click(`${box} button[aria-label^="Edit reply 1"]`);
await page.fill(`${box} .oxe-thread-item textarea`, 'Done, checked');
await page.click(`${box} >> button:text-is("Save")`);
await page.waitForFunction((b) => [...document.querySelectorAll(`${b} .oxe-thread-text`)].some((x) => x.textContent === 'Done, checked'), box);
check(await threadNow() === 'Check these totals | Done, checked', 'the reply edited', await threadNow());
await page.click(`${box} button[aria-label^="Delete reply 1"]`);
await page.waitForFunction((b) => document.querySelectorAll(`${b} .oxe-thread-item`).length === 1, box);
const oneLeft = await threadNow();
await page.click(`${box} >> button:text-is("Close")`);
await page.keyboard.press('Control+z');
check(oneLeft === 'Check these totals' && await threadNow() === 'Check these totals | Done, checked', 'the reply deleted; Ctrl+Z brings it back', `${oneLeft} / ${await threadNow()}`);

// Conditional formatting on a new sheet's numbers 1 to 5: a blue data bar from the menu (drawn: the largest bar the
// widest), an icon set (5's icon the top one), New Rule's formula rule, and the Rules Manager deleting the data bar.
await page.evaluate(() => { const o = window.oxe; o.command('insertSheet'); for (let r = 1; r <= 5; r++) o.enter(`A${r}`, String(r)); o.select('A1'); o.command('extendTo', 4, 0); });
await page.click('.oxe-rtab >> text=Home');
await menu('Conditional Formatting', 'Blue Data Bar (Gradient Fill)');
await menu('Conditional Formatting', '3 Arrows (Colored)');
const barWidths = await page.evaluate(() => [...document.querySelectorAll('.oxe-cf-bar')].map((b) => parseFloat(b.style.getPropertyValue('--oxe-w'))));
const icons = await page.evaluate(() => document.querySelectorAll('.oxe-cf-icon').length);
check(barWidths.length === 5 && barWidths[4] > barWidths[0] * 4 && icons === 5, 'Data Bars and Icon Sets from the menu: five bars, the largest widest, five icons', `${barWidths} / ${icons}`);
await menu('Conditional Formatting', 'New Rule...');
await page.waitForSelector('dialog.oxe-cf-new');
await page.selectOption('dialog.oxe-cf-new select.oxe-dialog-list', 'formula');
await page.fill('dialog.oxe-cf-new .oxe-cf-panel input[type=text]', '=MOD(A1,2)=0');
await page.click('dialog.oxe-cf-new >> button:text-is("OK")');
const typesNow = () => page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).condFormats.flatMap((c) => c.rules).sort((a, b) => a.priority - b.priority).map((r) => r.type).join(','));
check(await typesNow() === 'expression,iconSet,dataBar', 'New Rule (Use a formula): first, before the others', await typesNow());
await menu('Conditional Formatting', 'Manage Rules...');
await page.waitForSelector('dialog.oxe-cf-manager');
await page.click('dialog.oxe-cf-manager tr.oxe-names-row:has-text("Data Bar") td >> nth=0');
await page.click('dialog.oxe-cf-manager >> button:text-is("Delete Rule")');
await page.waitForFunction(() => document.querySelectorAll('dialog.oxe-cf-manager tr.oxe-names-row').length === 2);
await page.click('dialog.oxe-cf-manager >> button:text-is("Close")');
check(await typesNow() === 'expression,iconSet' && await page.evaluate(() => document.querySelectorAll('.oxe-cf-bar').length) === 0, 'Rules Manager: the data bar deleted, its bars gone', await typesNow());

// Filters on a new sheet's table: Number Filters > greater than 2 on Qty, then Advanced Filter in place with a criteria
// range (Team blue), and Clear showing the rows again.
await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  [['Name', 'Qty', 'Team'], ['apple', 1, 'red'], ['pear', 2, 'blue'], ['fig', 3, 'red'], ['kiwi', 10, 'blue']].forEach((row, r) => row.forEach((v, c) => o.enter(String.fromCharCode(65 + c) + (r + 1), String(v))));
  o.enter('F1', 'Team'); o.enter('F2', 'blue');
  o.select('A1');
});
await page.focus('.oxe-grid');
await page.keyboard.press('Control+Shift+L');
const hiddenNow = () => page.evaluate(() => [...window.oxe.workbook.sheet(window.oxe.sheetIndex).rows].filter(([, r]) => r.hidden).map(([r]) => r + 1).sort((a, b) => a - b).join(','));
await page.evaluate(() => window.oxe.select('B1'));
await page.keyboard.press('Alt+ArrowDown');
await page.waitForSelector('dialog.oxe-filter-box');
await page.click('dialog.oxe-filter-box >> button:text-is("Number Filters...")');
await page.waitForSelector('dialog.oxe-custom-filter');
await page.selectOption('dialog.oxe-custom-filter fieldset select >> nth=0', '>');
await page.fill('dialog.oxe-custom-filter input[aria-label="First value"]', '2');
await page.click('dialog.oxe-custom-filter >> button:text-is("OK")');
check(await hiddenNow() === '2,3', 'Number Filters: Qty greater than 2 hides the rows of 1 and 2', await hiddenNow());
await page.keyboard.press('Control+Shift+L');
await page.click('.oxe-rtab >> text=Data');
await page.evaluate(() => window.oxe.select('A1'));
await ribbon('Advanced...');
await page.waitForSelector('dialog.oxe-adv-filter');
await page.fill('dialog.oxe-adv-filter .oxe-fc-field:has-text("Criteria range") input', '$F$1:$F$2');
await page.click('dialog.oxe-adv-filter >> button:text-is("OK")');
check(await hiddenNow() === '2,4', 'Advanced Filter in place: Team blue keeps pear and kiwi', await hiddenNow());
await page.locator('.oxe-toolbar button[title="Clear"]:visible').click();
check(await hiddenNow() === '', 'Clear shows the rows again', await hiddenNow());

// The data tools on a new sheet: Fill > Series (1 by 3 down A1:A4), Flash Fill (Ctrl+E) of first names, Remove
// Duplicates on a list, Text to Columns at commas, and Move or Copy > Create a copy, then Ctrl+Z taking the copy back.
await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  o.enter('A1', '1');
  [['Full', 'First'], ['Avery Chen', 'Avery'], ['Sam Lee', ''], ['Jo Park', '']].forEach((row, r) => row.forEach((v, c) => { if (v) o.enter(String.fromCharCode(67 + c) + (r + 1), v); }));
  [['Fruit'], ['pear'], ['fig'], ['Pear'], ['kiwi']].forEach((row, r) => o.enter(`F${r + 1}`, row[0]));
  o.enter('H1', 'a,b,c');
  o.select('A1'); o.command('extendTo', 3, 0);
});
const cellText = (ref) => page.evaluate((ref) => { const r = Number(ref.slice(1)) - 1; const c = ref.charCodeAt(0) - 65; const x = window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c); return x?.v == null ? '' : String(x.v); }, ref);
await page.click('.oxe-rtab >> text=Home');
await menu('Fill', 'Series...');
await page.waitForSelector('dialog.oxe-series-box');
await page.fill('dialog.oxe-series-box .oxe-fc-field:has-text("Step value") input', '3');
await page.click('dialog.oxe-series-box >> button:text-is("OK")');
check([await cellText('A2'), await cellText('A4')].join() === '4,10', 'Fill > Series: 1, 4, 7, 10', [await cellText('A2'), await cellText('A4')].join());
await page.evaluate(() => window.oxe.select('D3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+e');
check([await cellText('D3'), await cellText('D4')].join() === 'Sam,Jo', 'Flash Fill (Ctrl+E): first names', [await cellText('D3'), await cellText('D4')].join());
await page.evaluate(() => { window.oxe.select('F1'); window.oxe.command('extendTo', 4, 5); });
await page.click('.oxe-rtab >> text=Data');
await ribbon('Remove Duplicates...');
await page.waitForSelector('dialog.oxe-dups-box');
await page.click('dialog.oxe-dups-box >> button:text-is("OK")');
await page.waitForSelector('dialog.oxe-dialog >> text=duplicate values found and removed');
await page.keyboard.press('Enter');
check([await cellText('F2'), await cellText('F3'), await cellText('F4'), await cellText('F5')].join() === 'pear,fig,kiwi,', 'Remove Duplicates: Pear gone, kiwi moved up', [await cellText('F4'), await cellText('F5')].join());
await page.evaluate(() => window.oxe.select('H1'));
await ribbon('Text to Columns...');
await page.waitForSelector('dialog.oxe-ttc-box');
await page.locator('dialog.oxe-ttc-box label:text-is("Comma")').click();
await page.click('dialog.oxe-ttc-box >> button:text-is("OK")');
check([await cellText('H1'), await cellText('I1'), await cellText('J1')].join() === 'a,b,c', 'Text to Columns: a, b, c', [await cellText('H1'), await cellText('I1')].join());
const sheetsNow = () => page.evaluate(() => window.oxe.workbook.sheets.map((s) => s.name).join('|'));
const before = await sheetsNow();
const tab = page.locator('.oxe-tab[aria-selected="true"]').first();
await tab.click({ button: 'right' });
await page.click('.oxe-menu [role="menuitem"]:text-is("Move or Copy...")');
await page.waitForSelector('dialog.oxe-movecopy-box');
await page.locator('dialog.oxe-movecopy-box label:text-is("Create a copy")').click();
await page.click('dialog.oxe-movecopy-box >> button:text-is("OK")');
const after = await sheetsNow();
await page.keyboard.press('Control+z');
check(after.split('|').length === before.split('|').length + 1 && /\(2\)/.test(after) && await sheetsNow() === before, 'Move or Copy > Create a copy: a "(2)" sheet; Ctrl+Z takes it back', after);

// View > Show: Gridlines, Formula Bar and Headings turned off and on again (Excel's saved form is tested in
// view.test.js): the gridlines not drawn, the bar hidden, the headings taking no room; the buttons follow.
await page.click('.oxe-rtab >> text="View"');
const showState = () => page.evaluate(() => {
  const g = window.oxe.grid.el;
  const pressed = (l) => document.querySelector(`.oxe-toolbar button[title="${l}"]`)?.getAttribute('aria-pressed');
  return [g.querySelectorAll('.oxe-lines > *').length > 0, !document.querySelector('.oxe-bar').hidden, g.querySelector('.oxe-corner').getBoundingClientRect().width > 0,
    pressed('Gridlines'), pressed('Formula Bar'), pressed('Headings'), window.oxe.workbook.sheet(window.oxe.sheetIndex ?? 0).showHeadings].join();
});
const shownBefore = await showState();
for (const l of ['Gridlines', 'Formula Bar', 'Headings']) await ribbon(l);
const shownOff = await showState();
for (const l of ['Gridlines', 'Formula Bar', 'Headings']) await ribbon(l);
check(shownBefore === 'true,true,true,true,true,true,true' && shownOff === 'false,false,false,false,false,false,false' && await showState() === shownBefore,
  'View > Show: Gridlines, Formula Bar and Headings off and on again', JSON.stringify({ shownBefore, shownOff }));

// Page Layout > Sheet Options: Print ticked for gridlines and headings (kept for printing), View the same switch as
// View > Show.
await page.click('.oxe-rtab >> text="Page Layout"');
await ribbon('Gridlines: Print'); await ribbon('Headings: Print'); await ribbon('Gridlines: View');
const sheetOpts = await page.evaluate(() => { const s = window.oxe.workbook.sheet(window.oxe.sheetIndex); return [s.page.gridLines, s.page.headings, s.showGridLines].join(); });
const printPressed = await page.getAttribute('.oxe-toolbar button[title="Gridlines: Print"]', 'aria-pressed');
await ribbon('Gridlines: Print'); await ribbon('Headings: Print'); await ribbon('Gridlines: View');
check(sheetOpts === 'true,true,false' && printPressed === 'true' && await page.evaluate(() => { const s = window.oxe.workbook.sheet(window.oxe.sheetIndex); return [s.page.gridLines, s.page.headings, s.showGridLines].join(); }) === 'false,false,true',
  'Page Layout > Sheet Options: Print gridlines and headings ticked and unticked; View the gridlines switch', sheetOpts);


// AutoCorrect as a cell is typed in (autocorrect.js; the list Excel's own): a word corrected when it ends and when the
// cell is entered; Ctrl+Z while typing gives back what was typed; the spelling underlines on for text, off for a formula.
await page.evaluate(() => window.oxe.select('K60'));
await page.keyboard.type('teh cat on monday');
const typedNow = await page.inputValue('.oxe-editor');
const spelling = await page.evaluate(() => document.querySelector('.oxe-editor').spellcheck);
await page.keyboard.press('Enter');
await page.evaluate(() => window.oxe.select('K61'));
await page.keyboard.type('adn ');
await page.keyboard.press('Control+z');
const undone = await page.inputValue('.oxe-editor');
await page.keyboard.press('Escape');
await page.keyboard.type('=1+1');
const formulaSpelling = await page.evaluate(() => document.querySelector('.oxe-editor').spellcheck);
await page.keyboard.press('Escape');
const k60 = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(59)?.get(10)?.v);
check(typedNow === 'the cat on monday' && k60 === 'the cat on Monday' && undone === 'adn ' && spelling === true && formulaSpelling === false,
  'AutoCorrect: teh to the as the word ends, monday to Monday on Enter, Ctrl+Z gives back what was typed; underlines for text only',
  JSON.stringify({ typedNow, k60, undone, spelling, formulaSpelling }));

// Rich text (made here as Excel writes it: richtext.test.js): each run drawn in its own font over the cell's.
{
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.setCell(0, 'J12', 'x');
  const { doc } = wb._part(0);
  const ns = doc.documentElement.namespaceURI;
  const c = [...doc.getElementsByTagNameNS(ns, 'c')].find((e) => e.getAttribute('r') === 'J12');
  while (c.firstChild) c.removeChild(c.firstChild);
  c.setAttribute('t', 'inlineStr');
  const rPr = (inner) => `<rPr>${inner}<sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr>`;
  const { parseXml } = await import('../src/ooxml/xml.js');
  const is = parseXml(`<is xmlns="${ns}"><r><t xml:space="preserve">Plain </t></r><r>${rPr('<b/>')}<t>bold</t></r>`
    + `<r><rPr><i/><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/><family val="3"/></rPr><t xml:space="preserve"> red</t></r></is>`).documentElement;
  c.appendChild(doc.importNode(is, true));
  wb.dirty.add(wb.sheets[0].path);
  fs.writeFileSync('test/out/rich.xlsx', wb.toBytes());
}
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/out/rich.xlsx', 'rich.xlsx');
const runs = await page.evaluate(() => {
  const el = document.querySelector('.oxe-row[aria-rowindex="13"] .oxe-cell[aria-colindex="11"] .oxe-runs');
  return el ? [...el.children].map((x) => { const st = getComputedStyle(x); return [x.textContent, st.fontWeight, st.fontStyle, st.fontSize, st.color, st.fontFamily.split(',')[0]]; }) : null;
});
check(runs?.length === 3 && runs[0][1] === '400' && runs[1][0] === 'bold' && runs[1][1] === '700' && runs[2][2] === 'italic' && runs[2][3] === '21.3333px' && runs[2][4] === 'rgb(255, 0, 0)' && /Courier New/.test(runs[2][5]),
  'rich text: each run in its own font (bold; italic, 16pt, red, Courier New) over the cell’s', JSON.stringify(runs));

// A filled column's empty cells drawn filled, as Excel draws a column's own format (linestyles.test.js).
let lineFill;
{
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  lineFill = wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } });
  const { doc } = wb._part(0);
  const root = doc.documentElement;
  const old = [...root.childNodes].find((n) => n.localName === 'cols');
  if (old) root.removeChild(old);
  const cols = parseXml(`<cols xmlns="${root.namespaceURI}"><col min="12" max="12" width="9.140625" style="${lineFill}"/></cols>`).documentElement;
  root.insertBefore(doc.importNode(cols, true), [...root.childNodes].find((n) => n.localName === 'sheetData'));
  wb.dirty.add(wb.sheets[0].path);
  fs.writeFileSync('test/out/linefill.xlsx', wb.toBytes());
}
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/out/linefill.xlsx', 'linefill.xlsx');
const filled = await page.evaluate((f) => [...document.querySelectorAll(`.oxe-f${f}`)].filter((e) => !e.classList.contains('oxe-cell')).length, lineFill);
const rowsShown = await page.evaluate(() => document.querySelectorAll('.oxe-row').length);
check(filled >= rowsShown - 2, 'a column’s own fill: its empty cells drawn filled', JSON.stringify({ filled, rowsShown }));

await finish();
