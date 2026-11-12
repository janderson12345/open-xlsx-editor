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

// Review > Allow Edit Ranges: a range of the selection's cells (Range1), one with a password; protected, the first
// changes though locked, the second asks for its password (Unlock Range), then changes; the box can't change them.
const cellV = (r, c) => page.evaluate(([r, c]) => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c)?.v ?? null, [r, c]);
const newRange = async (password) => {
  await page.click('dialog.oxe-dialog >> button:text-is("New...")');
  await page.waitForSelector('dialog.oxe-dialog >> text=Refers to cells:');
  const offered = await page.$$eval('dialog.oxe-dialog[open] input', (xs) => xs.slice(-3).map((x) => x.value));
  if (password) {
    await page.fill('dialog.oxe-dialog[open]:last-of-type input[type=password]', password);
    await page.fill('dialog.oxe-dialog[open]:last-of-type input >> nth=1', '=$K$82');
  }
  await page.click('dialog.oxe-dialog[open]:last-of-type >> button:text-is("OK")');
  if (password) {
    await page.waitForSelector('dialog.oxe-dialog >> text=Reenter password to proceed.');
    await page.keyboard.type(password);
    await page.keyboard.press('Enter');
  }
  return offered;
};
await page.evaluate(() => window.oxe.select('K80'));
await ribbon('Allow Edit Ranges');
await page.waitForSelector('dialog.oxe-dialog >> text=Allow Users to Edit Ranges');
const offered = await newRange('');
await newRange('Range-pass 2');
const listed = await page.$$eval('dialog.oxe-dialog .oxe-names-row', (rs) => rs.map((r) => r.textContent).join('|'));
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await ribbon('Protect Sheet');
await page.waitForSelector('dialog.oxe-dialog >> text=Allow all users of this worksheet to:');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.evaluate(() => window.oxe.select('K80'));
await page.focus('.oxe-grid');
await page.keyboard.type('5');
await page.keyboard.press('Enter');
const k80 = await cellV(79, 10);
await page.evaluate(() => window.oxe.select('K82'));
await page.focus('.oxe-grid');
await page.keyboard.type('6');
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-dialog >> text=Unlock Range');
const unlockText = await page.textContent('dialog.oxe-dialog[open]');
await page.keyboard.type('Range-pass 2');
await page.keyboard.press('Enter');
await page.waitForFunction(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(81)?.get(10)?.v === 6);
await ribbon('Allow Edit Ranges');
await page.waitForSelector('dialog.oxe-dialog >> text=Allow Users to Edit Ranges');
const lockedButtons = await page.$$eval('dialog.oxe-dialog button', (bs) => bs.filter((b) => b.disabled).map((b) => b.textContent).join('|'));
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
check(offered.join() === 'Range1,=$K$80,' && listed === 'Range1=$K$80|Range2=$K$82' && k80 === 5 && /password protected/.test(unlockText) && await cellV(81, 10) === 6 && lockedButtons === 'New...|Modify...|Delete',
  'Allow Edit Ranges: Range1 of the selection; protected, it changes though locked; one with a password asks Unlock Range, then changes; the box locked',
  JSON.stringify({ offered, listed, k80, unlockText, lockedButtons }));

await finish();
