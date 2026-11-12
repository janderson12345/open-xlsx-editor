// Browser check of hyperlinks: a click on a hyperlink's cell follows it (an address opened, a place in the workbook
// gone to), a click held selects the cell instead, Excel's ScreenTip and hand over it, and the cell menu's Open
// Hyperlink. Run after `npm run build`: node test/editing-links.mjs [outDir]
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { addHyperlink } from '../src/ooxml/links.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, sel, clickCell, pointAt } = await editingPage();

// (Made here, as Excel's Hyperlinks.Add writes them: links.test.js.)
const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
addHyperlink(wb, 0, 2, 9, { address: 'https://example.com/page' }); // J3
addHyperlink(wb, 0, 3, 9, { location: "'Notes & more'!C30", text: 'Go there' }); // J4
addHyperlink(wb, 0, 4, 9, { address: 'example.com', tooltip: 'Tip here' }); // J5
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/links.xlsx', wb.toBytes());
await openFixtureFrom(page, 'test/out/links.xlsx', 'links.xlsx');
await page.evaluate(() => { window.opened = []; window.open = (url) => { window.opened.push(url); return null; }; });
const opened = () => page.evaluate(() => window.opened.slice());

// Over J3: the hand and Excel's ScreenTip; a click opens the address.
await page.evaluate(() => window.oxe.select('J1'));
const p = await pointAt(2, 9);
await page.mouse.move(p.x, p.y);
const tip = await page.getAttribute('.oxe-grid', 'title');
const hand = await page.evaluate(() => getComputedStyle(document.querySelector('.oxe-grid')).cursor);
await clickCell(2, 9);
const first = await opened();
check(tip === 'https://example.com/page - Click once to follow. Click and hold to select this cell.' && hand === 'pointer' && first.join() === 'https://example.com/page',
  'a hyperlink: the hand and Excel’s ScreenTip over it; a click opens its address', JSON.stringify({ tip, hand, first }));

// Held, it selects the cell and opens nothing.
await page.mouse.move(p.x, p.y);
await page.mouse.down();
await page.waitForTimeout(700);
await page.mouse.up();
const held = await opened();
check(held.length === 1 && (await sel()).ar === 2, 'a hyperlink clicked and held: its cell selected, nothing opened', JSON.stringify({ held, sel: await sel() }));

// J4, a place in the workbook: gone to.
await clickCell(3, 9);
await page.waitForFunction(() => window.oxe.sheetIndex === 1);
const there = await sel();
check(there.ar === 29 && there.ac === 2, 'a hyperlink to a place in the workbook: that sheet shown, the cell selected', JSON.stringify(there));

// J5, a file's path: its own ScreenTip; refused in Excel's words; the cell menu's Open Hyperlink does the same.
await page.evaluate(() => window.oxe.showSheet(0));
const q = await pointAt(4, 9);
await page.mouse.move(q.x + 1, q.y + 1);
const ownTip = await page.getAttribute('.oxe-grid', 'title');
await clickCell(4, 9);
await page.waitForSelector('dialog.oxe-dialog[role="alertdialog"]');
const refusal = await page.textContent('dialog.oxe-dialog[role="alertdialog"]');
await page.keyboard.press('Enter');
const j3 = await pointAt(2, 9);
await page.mouse.click(j3.x, j3.y, { button: 'right' });
await page.waitForSelector('.oxe-menu >> text=Open Hyperlink');
const items = await page.$$eval('.oxe-menu [role="menuitem"]', (xs) => xs.map((x) => x.textContent.trim()).join('|'));
await page.click('.oxe-menu >> text=Open Hyperlink');
check(ownTip.startsWith('Tip here - ') && /Cannot open the specified file\./.test(refusal || '') && /Open Hyperlink\|Remove Hyperlink/.test(items) && (await opened()).length === 2,
  'a file’s path: its own ScreenTip, refused in Excel’s words; the cell menu’s Open Hyperlink opens J3’s address', JSON.stringify({ ownTip, refusal, items }));

await finish();
