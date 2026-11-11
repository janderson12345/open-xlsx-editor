// Browser check of editing, continued: a workbook saved with a password, PivotTables (refreshed and made), the Data
// tab, conditional formatting, several ranges selected, Go To Special, and sheets undone.
// Run after `npm run build`: node test/editing-more.mjs [outDir]
import fs from 'node:fs';
import { encryptWorkbook } from '../src/ooxml/encryption.js';
import { pivotBook } from './pivots.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, shown, sel, clickCell, pointAt, ribbon } = await editingPage();

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

await finish();
