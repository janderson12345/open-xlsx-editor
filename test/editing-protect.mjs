// Browser check of protection: Review > Protect Sheet (a locked cell refused, a hidden cell's formula not shown) and
// Protect Workbook (a new sheet refused).
// Run after `npm run build`: node test/editing-protect.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish, shown, ribbon } = await editingPage();
await openFixture(page, 'basic.xlsx');
await page.focus('.oxe-grid');
const bar = () => page.inputValue('.oxe-formula');

// A hidden cell (Format Cells > Protection): its formula shown until the sheet is protected, then the formula bar shows
// nothing there and Show Formulas shows it as nothing, as in Excel (measured); FORMULATEXT gives #N/A.
await page.evaluate(() => { window.oxe.enter('K72', '=1+2'); window.oxe.enter('L72', '=FORMULATEXT(K72)'); window.oxe.select('K72'); window.oxe.command('format', { protection: { hidden: true } }); });
const barBefore = await bar();
await page.click('.oxe-rtab >> text="Review"');
await ribbon('Protect Sheet');
await page.waitForSelector('dialog.oxe-dialog >> text=Allow all users of this worksheet to:');
const protectItems = await page.$$eval('dialog.oxe-dialog .oxe-protect-list label', (ls) => ls.map((l) => l.textContent).join('|'));
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.evaluate(() => { window.oxe.select('K72'); window.oxe.command('calculate', 'all'); });
const barHidden = await bar();
const textHidden = await shown(71, 11);
await page.evaluate(() => window.oxe.command('showFormulas', true));
const formulasHidden = await shown(71, 10);
await page.evaluate(() => window.oxe.command('showFormulas', false));
check(barBefore === '=1+2' && barHidden === '' && textHidden === '#N/A' && formulasHidden === '',
  'a hidden cell of a protected sheet: no formula in the formula bar, none with Show Formulas; FORMULATEXT gives #N/A', JSON.stringify({ barBefore, barHidden, textHidden, formulasHidden }));
await page.evaluate(() => window.oxe.select('K70'));
const k70Before = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(69)?.get(10)?.v ?? null);
await page.keyboard.type('9');
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-dialog[role="alertdialog"]');
const refusal = await page.textContent('dialog.oxe-dialog[role="alertdialog"]');
await page.keyboard.press('Enter');
const k70After = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(69)?.get(10)?.v ?? null);
const unprotectLabel = await page.getAttribute('.oxe-toolbar button[title="Unprotect Sheet"]', 'aria-label');
await ribbon('Unprotect Sheet');
await page.evaluate(() => window.oxe.select('K70'));
await page.focus('.oxe-grid');
await page.keyboard.type('9');
await page.keyboard.press('Enter');
const k70Unprotected = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(69)?.get(10)?.v ?? null);
check(protectItems.startsWith('Select locked cells|Select unlocked cells|Format cells') && protectItems.endsWith('Edit objects|Edit scenarios') && /on a protected sheet/.test(refusal || '')
  && k70After === k70Before && unprotectLabel === 'Unprotect Sheet' && k70Unprotected === 9,
  'Review > Protect Sheet: Excel’s box; a locked cell refused in Excel’s words; Unprotect Sheet lets it change', JSON.stringify({ protectItems, refusal, k70Before, k70After, unprotectLabel, k70Unprotected }));


// Review > Protect Workbook: the structure protected, a new sheet refused, the button pressed; again, unprotected.
await ribbon('Protect Workbook');
await page.waitForSelector('dialog.oxe-dialog >> text=Protect Structure and Windows');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const sheetCount = () => page.evaluate(() => window.oxe.workbook.sheets.length);
const sheetsBefore = await sheetCount();
await page.click('.oxe-tab-add');
await page.waitForSelector('dialog.oxe-dialog[role="alertdialog"]');
const bookRefusal = await page.textContent('dialog.oxe-dialog[role="alertdialog"]');
await page.keyboard.press('Enter');
const pressed = await page.getAttribute('.oxe-toolbar button[title="Protect Workbook"]', 'aria-pressed');
const afterRefusal = await sheetCount();
await ribbon('Protect Workbook');
await page.click('.oxe-tab-add');
check(/protected/.test(bookRefusal || '') && pressed === 'true' && afterRefusal === sheetsBefore && await sheetCount() === sheetsBefore + 1,
  'Review > Protect Workbook: a new sheet refused while protected; unprotected, it is added', JSON.stringify({ bookRefusal, pressed, sheetsBefore, afterRefusal }));

await finish();
