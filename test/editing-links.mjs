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

// Ctrl+K on an empty cell: Excel's Insert Hyperlink box; a web address and text to display; then Edit Hyperlink... on
// the cell's menu, to a place in the workbook; then its Remove Link.
const linkOf = (r, c) => page.evaluate(([r, c]) => {
  const wbk = window.oxe.workbook;
  const { doc } = wbk._part(window.oxe.sheetIndex);
  const el = [...doc.getElementsByTagNameNS(doc.documentElement.namespaceURI, 'hyperlink')].find((h) => h.getAttribute('ref') === String.fromCharCode(65 + c) + (r + 1));
  if (!el) return null;
  const id = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
  return { location: el.getAttribute('location'), address: id ? wbk.pkg.readRels(wbk.sheets[window.oxe.sheetIndex].path).find((x) => x.id === id).rawTarget : null };
}, [r, c]);
await page.keyboard.press('Escape');
await page.evaluate(() => window.oxe.select('J8'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+k');
await page.waitForSelector('dialog.oxe-dialog >> text=Insert Hyperlink');
const kinds = await page.$$eval('dialog.oxe-dialog fieldset label', (ls) => ls.map((l) => l.textContent).join('|'));
await page.fill('dialog.oxe-dialog input[id$="-show"]', 'New link');
await page.fill('dialog.oxe-dialog input[id$="-addr"]', 'www.example.com/new');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const made = await linkOf(7, 9);
const j8 = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(7)?.get(9)?.v);
const j8Look = await page.evaluate(() => { const el = document.querySelector('.oxe-row[aria-rowindex="9"] .oxe-cell[aria-colindex="11"]'); return el ? getComputedStyle(el).textDecorationLine : null; });
check(kinds === 'Existing File or Web Page|Place in This Document|E-mail Address' && made?.address === 'http://www.example.com/new' && j8 === 'New link' && /underline/.test(j8Look || ''),
  'Ctrl+K: Excel’s Insert Hyperlink box; a web address (http:// before www.) and its text, underlined in the Hyperlink style', JSON.stringify({ kinds, made, j8, j8Look }));
const j8p = await pointAt(7, 9);
await page.mouse.click(j8p.x, j8p.y, { button: 'right' });
await page.click('.oxe-menu >> text=Edit Hyperlink...');
await page.waitForSelector('dialog.oxe-dialog >> text=Edit Hyperlink');
const shownAddr = await page.inputValue('dialog.oxe-dialog input[id$="-addr"]');
await page.click('dialog.oxe-dialog label:text-is("Place in This Document")');
await page.fill('dialog.oxe-dialog input[id$="-ref"]', 'c5');
await page.selectOption('dialog.oxe-dialog select', 's:Notes & more');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const edited = await linkOf(7, 9);
await page.keyboard.press('Control+k');
await page.waitForSelector('dialog.oxe-dialog >> text=Edit Hyperlink');
await page.click('dialog.oxe-dialog >> button:text-is("Remove Link")');
const removed = await linkOf(7, 9);
await page.keyboard.press('Control+z');
const undone = await linkOf(7, 9);
check(shownAddr === 'http://www.example.com/new' && edited?.location === "'Notes & more'!C5" && !edited.address && removed === null && undone?.location === "'Notes & more'!C5",
  'Edit Hyperlink: the link shown, changed to a place in the workbook; Remove Link; Ctrl+Z brings it back', JSON.stringify({ shownAddr, edited, removed, undone }));

await finish();
