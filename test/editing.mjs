// Browser check of editing: selection by keys and mouse, typing over a cell, F2, Esc, Tab, Delete, undo and redo,
// typed numbers and text, formulas typed and worked out again, sheet tabs and managing sheets, read-only; the saved
// workbook read back.
// Standards mode asserted.
// Run after `npm run build`: node test/editing.mjs [outDir]
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../src/ooxml/xml.js';
import { Workbook } from '../src/ooxml/workbook.js';
import { openDemo, openFixture, checker } from './browser.mjs';

setXmlImpl({ DOMParser, XMLSerializer });
const outDir = process.argv[2] || 'test/out';
fs.mkdirSync(outDir, { recursive: true });
const { page, errors, close } = await openDemo({ width: 1100, height: 760 });
const t = checker();
const { check } = t;

const shown = (r, c) => page.evaluate(([r, c]) => document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="${c + 2}"]`)?.textContent ?? null, [r, c]);
const sel = () => page.evaluate(() => window.oxe.selection);
const value = (ref) => page.evaluate((ref) => {
  const { r, c } = { r: Number(ref.slice(1)) - 1, c: ref.charCodeAt(0) - 65 };
  const cell = window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c);
  return cell ? cell.v : null;
}, ref);
const live = () => page.evaluate(() => document.querySelector('.oxe-live')?.textContent || '');
/** Where cell (r, c) is on the page, from the grid's geometry (empty cells have no element to click). */
const pointAt = (r, c) => page.evaluate(([r, c]) => {
  const g = window.oxe.grid;
  const v = g.el.getBoundingClientRect();
  const { rows, cols, rowHeadW, colHeadH } = g.geometry();
  return { x: v.left + rowHeadW + cols.start(c) - g.el.scrollLeft + 6, y: v.top + colHeadH + rows.start(r) - g.el.scrollTop + 6 };
}, [r, c]);
const clickCell = async (r, c) => { const p = await pointAt(r, c); await page.mouse.click(p.x, p.y); };
// A ribbon button, by its title: groups that don't fit at this width (it depends on the fonts) are behind the ribbon's
// More button, opened for the click and closed after `then` (a menu choice, a box answered).
const ribbon = async (title, then = async () => {}) => {
  const btn = `.oxe-toolbar button[title="${title}"]`;
  const more = !(await page.isVisible(btn)) && await page.isVisible('.oxe-more');
  if (more) await page.click('.oxe-more');
  await page.click(btn);
  await then();
  if (more && await page.isVisible('.oxe-more')) await page.click('.oxe-more');
};
const menu = (title, item) => ribbon(title, () => page.click(`.oxe-pop >> text=${item}`));

await openFixture(page, 'basic.xlsx');
check((await sel()).ar === 0 && (await sel()).ac === 0, 'opens with A1 selected');
await page.focus('.oxe-grid');

// Typing over a cell, Enter moves down.
await clickCell(1, 1); // B2
await page.keyboard.type('77');
await page.keyboard.press('Enter');
let s = await sel();
check(await value('B2') === 77 && s.ar === 2 && s.ac === 1, 'typing over B2 and Enter: 77 stored, B3 active');
check(await shown(1, 1) === '77', 'B2 shows 77');

// F2 keeps the contents and adds to them; Esc gives up.
await clickCell(1, 0); // A2 "Item 1"
await page.keyboard.press('F2');
await page.keyboard.type(' x');
await page.keyboard.press('Enter');
check(await value('A2') === 'Item 1 x', 'F2 then typing adds to "Item 1"');
await clickCell(2, 0); // A3
await page.keyboard.type('zzz');
await page.keyboard.press('Escape');
check(await value('A3') === 'Item 2', 'Esc leaves A3 as it was');

// Tab enters and moves right; typed entries read as Excel reads them.
await clickCell(15, 0); // A16, empty
for (const entry of ['50%', '1,234', "'007", 'true']) {
  await page.keyboard.type(entry);
  await page.keyboard.press('Tab');
}
check(JSON.stringify([await value('A16'), await value('B16'), await value('C16'), await value('D16')]) === JSON.stringify([0.5, 1234, '007', true]),
  "typed 50%, 1,234, '007, true: 0.5, 1234, text 007, TRUE");
check((await sel()).ac === 4, 'Tab moved along the row to E16');

// Delete empties a selected range; undo brings it back, redo empties it again.
await clickCell(1, 4); // E2
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Shift+ArrowDown');
s = await sel();
check(s.r0 === 1 && s.r1 === 3 && s.c0 === 4 && s.c1 === 4, 'Shift+Down twice selects E2:E4');
await page.keyboard.press('Delete');
check(await value('E2') === null && await value('E3') === null && await value('E4') === null, 'Delete empties E2:E4');
await page.keyboard.press('Control+z');
check(await value('E2') === 'No' && await value('E3') === 'Yes', 'Ctrl+Z brings them back');
await page.keyboard.press('Control+y');
check(await value('E2') === null, 'Ctrl+Y empties them again');
await page.keyboard.press('Control+z');

// Formulas: shown in the bar and the editor, worked out again when what they read changes, typed as Excel takes them.
const bar = () => page.evaluate(() => document.querySelector('.oxe-formula').value);
await clickCell(2, 2); // C3, a shared formula's cell (C2:C10 share B2*2)
check(await bar() === '=B3*2', 'the bar shows C3\'s own formula, moved from the shared one in C2', await bar());
await page.keyboard.press('F2');
check(await page.evaluate(() => document.querySelector('.oxe-editor').value) === '=B3*2', 'F2 edits the formula');
await page.keyboard.press('Escape');
await clickCell(2, 1); // B3
await page.keyboard.type('10');
await page.keyboard.press('Enter');
check(await value('C3') === 20 && await shown(2, 2) === '20', 'B3 set to 10: C3 (=B3*2) shows 20');
check(await value('B11') === 129 && await shown(10, 1) === '129', 'and the total in B11 (=SUM(B2:B10)) is 129', String(await value('B11')));
await clickCell(15, 4); // E16
await page.keyboard.type('=sum(b2:b3');
await page.keyboard.press('Enter');
check(await value('E16') === 87 && await shown(15, 4) === '87', '=sum(b2:b3 typed: 87');
await clickCell(15, 4);
check(await bar() === '=SUM(B2:B3)', 'and kept as Excel keeps it, =SUM(B2:B3)', await bar());
await page.keyboard.type('=1+');
await page.keyboard.press('Enter');
await page.waitForTimeout(250);
check(await value('E16') === 87 && /problem with this formula/.test(await live()), '=1+ is refused and said so', await live());
check(await page.evaluate(() => document.querySelector('.oxe-editor')?.value) === '=1+', 'and stays in the editor to be put right');
await page.keyboard.press('Escape');
await clickCell(1, 1); // B2
await page.keyboard.type('=B3/0');
await page.keyboard.press('Enter');
check(await shown(1, 1) === '#DIV/0!' && await value('B11') === '#DIV/0!', 'an error in B2 reaches the total', String(await value('B11')));
await page.keyboard.press('Control+z');
check(await value('B2') === 77 && await value('B11') === 129, 'Ctrl+Z puts back 77, and the total with it');

// Moving: Ctrl+Down to the end of a block, Ctrl+Home back.
await clickCell(0, 0);
await page.keyboard.press('Control+ArrowDown');
check((await sel()).ar === 11, 'Ctrl+Down from A1 goes to the end of the block (A12, the merged range)', JSON.stringify(await sel()));
await page.keyboard.press('Control+Home');
check((await sel()).ar === 0 && (await sel()).ac === 0, 'Ctrl+Home goes to A1');

// The mouse: drag to select, double click to edit.
const a = await pointAt(1, 0);
const b = await pointAt(3, 2);
await page.mouse.move(a.x, a.y);
await page.mouse.down();
await page.mouse.move(b.x, b.y, { steps: 4 });
await page.mouse.up();
s = await sel();
check(s.r0 === 1 && s.c0 === 0 && s.r1 === 3 && s.c1 === 2, 'dragging from A2 to C4 selects A2:C4', JSON.stringify(s));
await page.mouse.dblclick(a.x, a.y);
check(await page.evaluate(() => document.activeElement?.classList.contains('oxe-editor')), 'double click opens the editor');
await page.keyboard.press('Escape');

// The active cell is the grid's active descendant, empty or not.
await clickCell(0, 0);
await page.keyboard.press('Control+End');
await page.keyboard.press('ArrowRight');
const ad = await page.evaluate(() => { const id = document.querySelector('.oxe-grid').getAttribute('aria-activedescendant'); return id && document.getElementById(id)?.getAttribute('role'); });
check(ad === 'gridcell', 'an empty active cell is still the grid\'s active descendant');

// Sheet tabs: the hidden sheet has none; clicking, the arrow keys on the tabs and Ctrl+Page Up/Down change sheet, each
// sheet keeping its selection; undo reaches a change on another sheet by showing it.
const tabs = () => page.evaluate(() => [...document.querySelectorAll('.oxe-tab')].map((el) => ({ name: el.textContent, on: el.getAttribute('aria-selected') === 'true', color: el.style.getPropertyValue('--oxe-tab-color') })));
const label = () => page.evaluate(() => document.querySelector('.oxe-grid').getAttribute('aria-label'));
const first = await tabs();
check(first.length === 2 && first[0].name === 'Data' && first[0].on && first[1].name === 'Notes & more' && first[1].color === '#F1A983', 'two tabs (the hidden sheet has none), Data marked, the second in its tab colour', JSON.stringify(first));
await page.evaluate(() => window.oxe.select('C5')); // (scrolled away from it: no clicking)
await page.click('.oxe-tab >> text=Notes & more');
await page.waitForTimeout(250);
check(await label() === 'Notes & more' && await shown(0, 0) === 'Notes' && (await sel()).ar === 0 && (await tabs())[1].on, 'clicking the tab shows Notes & more, A1 selected', await label());
check(/^Notes & more, A1, Notes/.test(await live()), 'and says which sheet, and the cell', await live());
await clickCell(1, 1); // B2
await page.keyboard.type('7');
await page.keyboard.press('Enter');
check(await value('B2') === 7, 'typing on the second sheet');
await page.keyboard.press('Control+PageUp');
check(await label() === 'Data' && (await sel()).ar === 4 && (await sel()).ac === 2 && await page.evaluate(() => document.activeElement.classList.contains('oxe-grid')), 'Ctrl+Page Up: back on Data with C5 selected, the grid still focused');
await page.keyboard.press('Control+PageUp');
check(await label() === 'Data', 'Ctrl+Page Up on the first sheet stays there');
await page.keyboard.press('Control+z');
check(await label() === 'Notes & more' && await value('B2') === -3.5, 'Ctrl+Z on Data undoes the change on Notes & more, showing it');
await page.keyboard.press('Control+PageDown');
check(await label() === 'Notes & more', 'Ctrl+Page Down on the last sheet shown stays there (the hidden one is passed over)');
await page.focus('.oxe-tab[aria-selected="true"]');
await page.keyboard.press('Home');
check(await label() === 'Data' && await page.evaluate(() => document.activeElement.textContent) === 'Data', 'Home on the tabs: the first sheet, its tab focused');
await page.keyboard.press('ArrowRight');
check(await label() === 'Notes & more' && await page.evaluate(() => [...document.querySelectorAll('.oxe-tab')].map((el) => el.tabIndex).join()) === '-1,0', 'Right arrow: the next sheet, one tab in the tab order');
await page.keyboard.press('ArrowLeft');
await page.focus('.oxe-grid');

// Managing sheets: the + button, renaming in place (a refused name said in Excel's words, the box kept), the tab menu
// by mouse and keyboard (colour, hide, unhide, delete with a question first), dragging a tab, Shift+F11.
const names = async () => (await tabs()).map((x) => x.name).join(',');
const menuItem = (label) => page.click(`.oxe-menu [role="menuitem"]:text-is("${label}")`);
const dialogText = () => page.evaluate(() => document.querySelector('.oxe-dialog')?.textContent || '');
await page.click('.oxe-tab-add');
check(await names() === 'Data,Sheet1,Notes & more' && await label() === 'Sheet1', 'the + button: Sheet1 after the sheet shown, and shown', await names());
await page.dblclick('.oxe-tab >> text=Sheet1');
await page.keyboard.type('a/b');
await page.keyboard.press('Enter');
await page.waitForTimeout(100);
check(/invalid name for a sheet/.test(await dialogText()) && await page.evaluate(() => document.activeElement.textContent) === 'OK', 'a/b refused in a message box, Excel\'s words, OK focused');
await page.keyboard.press('Enter');
await page.waitForTimeout(100);
check(await page.evaluate(() => document.activeElement.className) === 'oxe-tab-edit', 'OK: back in the name box to put it right');
await page.keyboard.type('Summary');
await page.keyboard.press('Enter');
check(await names() === 'Data,Summary,Notes & more' && await label() === 'Summary', 'renamed Summary in place');
await page.evaluate(() => window.oxe.select('A1'));
await page.keyboard.type("='Notes & more'!B2*2");
await page.keyboard.press('Enter');
check(await value('A1') === -7, 'a formula reading Notes & more');
await page.click('.oxe-tab >> text=Notes & more', { button: 'right' });
check(await label() === 'Notes & more' && await page.evaluate(() => document.activeElement.textContent) === 'Insert', 'right click: the sheet shown, the menu open on Insert');
await menuItem('Rename');
await page.keyboard.press('Control+a');
await page.keyboard.type('Notes');
await page.keyboard.press('Enter');
check(await page.evaluate(() => window.oxe.workbook.sheet(1).cells.get(0).get(0).f) === 'Notes!B2*2', 'renamed Notes: the formula on Summary follows (Notes!B2*2)', await page.evaluate(() => window.oxe.workbook.sheet(1).cells.get(0).get(0).f));
await page.click('.oxe-tab >> text=Summary', { button: 'right' });
await menuItem('Tab Color');
await page.click('.oxe-swatch[aria-label="Red"]');
check((await tabs())[1].color === '#FF0000', 'Tab Color, Red: the tab is red');
const box = async (name) => page.evaluate((n) => { const r = [...document.querySelectorAll('.oxe-tab')].find((t) => t.textContent === n).getBoundingClientRect(); return { x: r.left, y: r.top + r.height / 2, w: r.width }; }, name);
const from = await box('Summary');
const past = await box('Notes');
await page.mouse.move(from.x + 10, from.y);
await page.mouse.down();
await page.mouse.move(from.x + 40, from.y, { steps: 4 });
await page.mouse.move(past.x + past.w - 4, past.y, { steps: 4 });
await page.mouse.up();
check(await names() === 'Data,Notes,Summary' && await label() === 'Summary', 'dragging Summary past Notes moves it there', await names());
await page.focus('.oxe-tab >> text=Notes');
await page.keyboard.press('Shift+F10');
check(await page.evaluate(() => document.activeElement.textContent) === 'Insert', 'Shift+F10 on a tab opens its menu');
await page.keyboard.press('Escape');
check(await page.evaluate(() => document.activeElement.textContent) === 'Notes' && !(await page.$('.oxe-menu')), 'Esc closes it, focus back on the tab');
await page.keyboard.press('Shift+F10');
for (const k of ['End', 'ArrowUp']) await page.keyboard.press(k);
await page.keyboard.press('Enter'); // (Hide)
check(await names() === 'Data,Summary' && await label() === 'Summary', 'Hide from the keyboard: Notes hidden, the next sheet shown', await names());
await page.click('.oxe-tab >> text=Data', { button: 'right' });
await menuItem('Unhide...');
check(await page.evaluate(() => [...document.querySelectorAll('.oxe-dialog option')].map((o) => o.textContent).join()) === 'Notes,Hidden', 'Unhide lists the hidden sheets');
await page.keyboard.press('Enter');
check(await names() === 'Data,Notes,Summary' && await label() === 'Notes', 'Enter: Notes back, and shown');
await page.click('.oxe-tab >> text=Summary', { button: 'right' });
await menuItem('Delete');
check(/permanently deleted/.test(await dialogText()), 'Delete on a sheet with something in it asks first');
await page.click('.oxe-dialog-button >> text=Delete');
check(await names() === 'Data,Notes' && await label() === 'Notes', 'deleted; the sheet before it shown', await names());
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+F11');
check(await names() === 'Data,Sheet2,Notes' && await label() === 'Sheet2', 'Shift+F11: Sheet2 before the sheet shown', await names());
await page.evaluate(() => window.oxe.deleteSheet('Sheet2'));
await page.evaluate(() => window.oxe.showSheet('Data'));
check(await names() === 'Data,Notes' && await label() === 'Data', 'deleted from the API without a question');
await page.focus('.oxe-grid');

// Pointing while a formula is typed: arrows (Shift extends), F4, the formula carried to another sheet and entered back
// on its own, a drag on the grid; the formula's ranges outlined, the one pointed at dashed.
const editorText = () => page.evaluate(() => document.querySelector('.oxe-editor')?.value ?? null);
const outlines = () => page.evaluate(() => [...document.querySelectorAll('.oxe-mark')].map((m) => (m.classList.contains('oxe-mark-point') ? 'dashed' : 'solid')).join());
await page.evaluate(() => window.oxe.select('H2'));
await page.focus('.oxe-grid');
await page.keyboard.type('=');
await page.keyboard.press('ArrowLeft');
check(await editorText() === '=G2' && await outlines() === 'dashed', '= then Left: G2 pointed at, outlined dashed', await editorText());
await page.keyboard.press('ArrowLeft');
await page.keyboard.press('Shift+ArrowDown');
check(await editorText() === '=F2:F3', 'Left, Shift+Down: F2:F3', await editorText());
await page.keyboard.type('+');
await page.keyboard.press('ArrowUp');
await page.keyboard.press('F4');
check(await editorText() === '=F2:F3+$H$1' && await outlines() === 'solid,dashed', '+, Up, F4: $H$1; F2:F3 outlined too', `${await editorText()} ${await outlines()}`);
await page.keyboard.type('*');
await page.click('.oxe-tab >> text=Notes');
check(await label() === 'Notes' && await bar() === '=F2:F3+$H$1*Notes!' && await page.evaluate(() => document.activeElement.className) === 'oxe-formula', 'a tab clicked: Notes shown, its name in the formula, which goes on in the bar', await bar());
await page.keyboard.press('ArrowDown');
check(await bar() === '=F2:F3+$H$1*Notes!B3', 'Down from Notes\'s active cell (B2, from before): Notes!B3', await bar());
await page.keyboard.press('Enter');
check(await label() === 'Data' && await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(1).get(7).f) === 'F2:F3+$H$1*Notes!B3' && await value('H2') === 0, 'Enter: back on Data, entered into H2');
await page.evaluate(() => window.oxe.select('H5'));
await page.focus('.oxe-grid');
await page.keyboard.type('=SUM(');
const p0 = await pointAt(1, 1);
const p1 = await pointAt(4, 2);
await page.mouse.move(p0.x, p0.y);
await page.mouse.down();
await page.mouse.move(p1.x, p1.y, { steps: 5 });
await page.mouse.up();
check(await editorText() === '=SUM(B2:C5', 'a drag from B2 to C5 points at B2:C5', await editorText());
await page.keyboard.press('Enter');
check(await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(4).get(7).f) === 'SUM(B2:C5)' && await value('H5') === 282 && await outlines() === '', 'Enter: =SUM(B2:C5) is 282 (B2 77 and B3 10 by now), the outlines gone', String(await value('H5')));
await page.evaluate(() => window.oxe.select('H5'));
await page.focus('.oxe-grid');
await page.keyboard.press('F2');
await page.keyboard.press('ArrowLeft');
check(await editorText() === '=SUM(B2:C5)' && await page.evaluate(() => document.querySelector('.oxe-editor').selectionStart) === 10, 'F2 (Edit mode): Left moves the caret, nothing pointed at');
await page.keyboard.press('Escape');

// The ribbon's Home tab: Bold, Fill Color and a border preset on G20:H20, Ctrl+B, undo, F6, and features turned off.
const looksOf = (ref) => page.evaluate((ref) => {
  const wb = window.oxe.workbook;
  const cell = wb.sheet(window.oxe.sheetIndex).cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65);
  const xf = wb.styles.cellXfs[cell?.s || 0];
  return { b: wb.styles.fonts[xf.fontId].b, fill: wb.styles.fills[xf.fillId].fg?.rgb ?? null, bottom: wb.styles.borders[xf.borderId].bottom.style };
}, ref);
await page.evaluate(() => window.oxe.select('G20'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowRight');
await page.click('.oxe-toolbar .oxe-b');
check((await looksOf('G20')).b && (await looksOf('H20')).b && await page.getAttribute('.oxe-toolbar .oxe-b', 'aria-pressed') === 'true', 'Bold on G20:H20, the button pressed');
await page.keyboard.press('Control+b');
check(!(await looksOf('G20')).b && !(await looksOf('H20')).b, 'Ctrl+B in the sheet (focus back there after a click): bold off');
await page.click('.oxe-toolbar .oxe-fillcolor');
check((await looksOf('H20')).fill === 'FFFFFF00', 'Fill Color: yellow, as Excel starts');
await page.click('.oxe-split-more[aria-label="Bottom Border options"]');
await page.click('.oxe-pop >> text=Thick Bottom Border');
check((await looksOf('G20')).bottom === 'medium' && (await looksOf('H20')).bottom === 'medium', 'Borders, Thick Bottom Border: a medium line under both');
await page.keyboard.press('Control+z');
check((await looksOf('G20')).bottom === null && (await looksOf('H20')).fill === 'FFFFFF00', 'Ctrl+Z takes the border away, the fill stays');
await page.keyboard.press('F6');
check(await page.evaluate(() => document.activeElement.getAttribute('role')) === 'tab', 'F6: to the ribbon tab');
await page.keyboard.press('Escape');
check(await page.evaluate(() => document.activeElement.classList.contains('oxe-grid')), 'Esc: back to the sheet');
await page.evaluate(() => window.oxe.setFeatures({ font: false }));
await page.keyboard.press('Control+b');
check(await page.isHidden('.oxe-group[data-feature="font"]') && !(await looksOf('G20')).b, 'features { font: false }: the Font group hidden, Ctrl+B does nothing');
await page.evaluate(() => window.oxe.setFeatures({ font: true }));
// Merge & Center on G22:H22, both filled in: Excel's question, OK, one value kept; Ctrl+Z puts both back.
await page.evaluate(() => { window.oxe.enter('G22', 'left'); window.oxe.enter('H22', 'right'); window.oxe.select('G22'); });
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowRight');
let question;
await ribbon('Merge & Center', async () => { question = await page.textContent('.oxe-dialog'); await page.click('.oxe-dialog-button >> text=OK'); });
await page.waitForFunction(() => window.oxe.workbook.sheet(0).merges.some((m) => m.r0 === 21 && m.c0 === 6));
check(/only keeps the upper-left value/.test(question) && await value('G22') === 'left' && await value('H22') === null, 'Merge & Center: Excel asks first, then G22:H22 merged with the value of G22', question);
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
check(await value('H22') === 'right' && !(await page.evaluate(() => window.oxe.workbook.sheet(0).merges.some((m) => m.r0 === 21))), 'Ctrl+Z: unmerged, H22 back');

// Copy and paste: C2 (=B2*2) to C30, its reference moved; Esc lets the copy go. Cut J32 and paste at L33: the formula
// that reads it follows; Ctrl+Z puts it back.
const formulaAt = (ref) => page.evaluate((ref) => window.oxe.workbook.sheet(0).cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65)?.f ?? null, ref);
await page.evaluate(() => window.oxe.select('C2'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+c');
check(await page.isVisible('.oxe-clip'), 'Ctrl+C: the copied cell outlined');
await page.evaluate(() => window.oxe.select('C30'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+v');
await page.waitForFunction(() => window.oxe.workbook.sheet(0).cells.get(29)?.get(2)?.f != null);
check(await formulaAt('C30') === 'B30*2' && await page.isVisible('.oxe-clip'), 'Ctrl+V: =B30*2 in C30, the copy kept for another paste', await formulaAt('C30'));
await page.keyboard.press('Escape');
check(!(await page.isVisible('.oxe-clip')), 'Esc: the copy let go');
await page.evaluate(() => { window.oxe.enter('J32', 'moved'); window.oxe.enter('K32', '=J32&"!"'); window.oxe.select('J32'); });
await page.focus('.oxe-grid');
await page.keyboard.press('Control+x');
await page.evaluate(() => window.oxe.select('L33'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+v');
await page.waitForFunction(() => window.oxe.workbook.sheet(0).cells.get(32)?.get(11)?.v === 'moved');
check(await value('J32') === null && await formulaAt('K32') === 'L33&"!"' && !(await page.isVisible('.oxe-clip')), 'Ctrl+X, Ctrl+V: moved to L33, the formula that read it follows, the cut let go', await formulaAt('K32'));
await page.keyboard.press('Control+z');
check(await value('J32') === 'moved' && await value('L33') === null && await formulaAt('K32') === 'J32&"!"', 'Ctrl+Z: back in J32, the formula too');

// Rows and columns: row 5's heading clicked, Ctrl++ inserts a row there; Ctrl+Z takes it away. Column E's heading, then
// Home, Delete, Delete Sheet Columns.
const headAt = (axis, i) => page.evaluate(([axis, i]) => {
  const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); const rc = g.cellRect(axis === 'row' ? i : 0, axis === 'row' ? 0 : i);
  return axis === 'row' ? { x: v.x + geo.rowHeadW / 2, y: v.y + geo.colHeadH + rc.y - g.el.scrollTop + rc.h / 2 } : { x: v.x + geo.rowHeadW + rc.x - g.el.scrollLeft + rc.w / 2, y: v.y + geo.colHeadH / 2 };
}, [axis, i]);
await page.evaluate(() => { window.oxe.select('A1'); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
const a5 = await value('A5');
let h = await headAt('row', 4);
await page.mouse.click(h.x, h.y);
check((await sel()).r0 === 4 && (await sel()).c1 === 16383, 'the heading of row 5: the whole row selected', JSON.stringify(await sel()));
await page.keyboard.press('Control+Shift+Equal');
await page.waitForFunction((a5) => window.oxe.workbook.sheet(0).cells.get(5)?.get(0)?.v === a5, a5);
check(await value('A5') === null && await value('A6') === a5, 'Ctrl++: a row inserted at 5, the rows below moved down');
await page.keyboard.press('Control+z');
check(await value('A5') === a5, 'Ctrl+Z: the row taken away again');
const f11 = await value('F11');
h = await headAt('col', 4);
await page.mouse.click(h.x, h.y);
check((await sel()).c0 === 4 && (await sel()).c1 === 4 && (await sel()).r1 === 1048575, 'the heading of column E: the whole column selected', JSON.stringify(await sel()));
await menu('Delete', 'Delete Sheet Columns');
check(f11 != null && await value('E11') === f11, 'Delete Sheet Columns: column E gone, F moved into its place', String(await value('E11')));
await page.keyboard.press('Control+z');
check(await value('F11') === f11, 'Ctrl+Z: column E back');

// Cells: B3:C4 selected, Ctrl++ asks with Excel's Insert box, Shift cells down chosen as Excel would; Enter shifts the
// cells below down. Ctrl+Z; B3:B5, Ctrl+-, Shift cells left chosen first, Shift cells up picked and OK.
await page.evaluate(() => { window.oxe.select('B3'); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowRight');
await page.keyboard.press('Shift+ArrowDown');
const [b3, b6, d3] = [await value('B3'), await value('B6'), await value('D3')];
await page.keyboard.press('Control+Shift+Equal');
await page.waitForSelector('dialog.oxe-dialog');
check(await page.isChecked('dialog.oxe-dialog >> label:text-is("Shift cells down") >> xpath=preceding-sibling::input'), 'Ctrl++ on B3:C4: Excel’s Insert box, Shift cells down chosen');
await page.keyboard.press('Enter');
// (A12:C13 is merged across the columns shifted: Excel's question first, OK to unmerge.)
await page.waitForSelector('dialog.oxe-dialog >> text=merged cells to unmerge');
check(await page.isVisible('dialog.oxe-dialog >> text=merged cells to unmerge'), 'Enter: Excel’s question, as A12:C13 is merged across the cells shifted');
await page.keyboard.press('Enter');
await page.waitForFunction((b3) => window.oxe.workbook.sheet(0).cells.get(4)?.get(1)?.v === b3, b3);
check(await value('B3') === null && await value('D3') === d3 && await value('B8') === b6 && await page.evaluate(() => window.oxe.workbook.sheet(0).merges.length) === 0, 'OK: B3:C4 inserted, the cells below moved down two, column D left as it was, A12:C13 unmerged', JSON.stringify(await page.evaluate(() => window.oxe.workbook.sheet(0).merges)));
await page.keyboard.press('Control+z');
check(await value('B3') === b3 && await value('B6') === b6 && await page.evaluate(() => window.oxe.workbook.sheet(0).merges.length) === 1, 'Ctrl+Z: the cells back where they were, merged again');
await page.evaluate(() => window.oxe.select('B3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Control+Minus');
await page.waitForSelector('dialog.oxe-dialog');
check(await page.isChecked('dialog.oxe-dialog >> label:text-is("Shift cells left") >> xpath=preceding-sibling::input'), 'Ctrl+- on B3:B5 (taller than wide): Excel’s Delete box, Shift cells left chosen');
await page.click('dialog.oxe-dialog >> label:text-is("Shift cells up")');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.waitForSelector('dialog.oxe-dialog >> text=merged cells to unmerge');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.waitForFunction((b6) => window.oxe.workbook.sheet(0).cells.get(2)?.get(1)?.v === b6, b6);
check(await value('B3') === b6 && await value('D3') === d3, 'Shift cells up: B6 moved up into B3, column D left as it was');
await page.keyboard.press('Control+z');
check(await value('B3') === b3, 'Ctrl+Z: back');

// Paste Special: B3:C4 copied, F20 right-clicked: the cell's menu, Paste Special...; Values and Transpose chosen, OK:
// the values across instead of down. Ctrl+Z; then the ribbon's Paste ▾ > Paste Link at F20.
const c3 = await value('C3');
const b4 = await value('B4');
await page.evaluate(() => window.oxe.select('B3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowRight');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Control+c');
await page.evaluate(() => window.oxe.select('F20'));
const f20 = await pointAt(19, 5);
await page.mouse.click(f20.x, f20.y, { button: 'right' });
await page.waitForSelector('.oxe-menu >> text=Paste Special...');
check(await page.isVisible('.oxe-menu >> text=Insert...') && await page.isVisible('.oxe-menu >> text=Clear Contents'), 'F20 right-clicked: the cell’s menu (Cut, Copy, Paste, Paste Special..., Insert..., Delete..., Clear Contents)');
await page.click('.oxe-menu >> text=Paste Special...');
await page.waitForSelector('dialog.oxe-dialog');
check(await page.isChecked('dialog.oxe-dialog >> label:text-is("All") >> xpath=preceding-sibling::input') && await page.isChecked('dialog.oxe-dialog >> label:text-is("None") >> xpath=preceding-sibling::input'), 'Paste Special...: Excel’s box, All and None chosen first');
await page.click('dialog.oxe-dialog >> label:text-is("Values")');
await page.click('dialog.oxe-dialog >> label:text-is("Transpose")');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.waitForFunction((b3) => window.oxe.workbook.sheet(0).cells.get(19)?.get(5)?.v === b3, b3);
check(await value('G20') === b4 && await value('F21') === c3, 'Values, Transpose: B3 in F20, B4 across in G20, C3 down in F21', `${await value('G20')} ${await value('F21')}`);
await page.keyboard.press('Control+z');
check(await value('F20') === null, 'Ctrl+Z: F20 empty again');
await menu('Paste (Ctrl+V) options', 'Paste Link');
await page.waitForFunction(() => window.oxe.workbook.sheet(0).cells.get(19)?.get(5)?.f === 'B3');
check(await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(20)?.get(6)?.f) === 'C4' && await value('F20') === b3, 'Paste ▾ > Paste Link: =B3 in F20 to =C4 in G21, worked out', String(await value('F20')));
await page.keyboard.press('Control+z');
await page.keyboard.press('Escape');

// Fill: "Item 1" typed in J16, its fill handle dragged down to J19: Item 2 to Item 4. Ctrl+Z. J16:J18 selected,
// Ctrl+D: Item 1 copied down. Ctrl+Z; Home > Fill > Right on J16:L16.
await page.evaluate(() => { window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
await page.evaluate(() => window.oxe.select('J16'));
await page.focus('.oxe-grid');
await page.keyboard.type('Item 1');
await page.keyboard.press('Enter');
await page.evaluate(() => window.oxe.select('J16'));
const corner = await page.evaluate(() => {
  const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); const rc = g.cellRect(15, 9);
  return { x: v.x + geo.rowHeadW + rc.x + rc.w - g.el.scrollLeft, y: v.y + geo.colHeadH + rc.y + rc.h - g.el.scrollTop };
});
check(await page.isVisible('.oxe-fill-handle'), 'the fill handle on the selection’s corner');
const j19 = await pointAt(18, 9);
await page.mouse.move(corner.x - 1, corner.y - 1);
await page.mouse.down();
await page.mouse.move(j19.x, j19.y, { steps: 4 });
check(await page.isVisible('.oxe-fill-drag'), 'dragged: the cells to fill outlined');
await page.mouse.up();
await page.waitForFunction(() => window.oxe.workbook.sheet(0).cells.get(18)?.get(9)?.v === 'Item 4');
check(await value('J17') === 'Item 2' && await value('J19') === 'Item 4' && (await sel()).r1 === 18, 'let go at J19: Item 2 to Item 4, as Excel fills, J16:J19 selected');
await page.keyboard.press('Control+z');
check(await value('J17') === null, 'Ctrl+Z: the filled cells empty again');
await page.evaluate(() => window.oxe.select('J16'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Control+d');
check(await value('J18') === 'Item 1' && await value('J17') === 'Item 1', 'Ctrl+D on J16:J18: Item 1 copied down');
await page.keyboard.press('Control+z');
await page.evaluate(() => window.oxe.select('J16'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowRight');
await page.keyboard.press('Shift+ArrowRight');
await menu('Fill', 'Right');
check(await value('L16') === 'Item 1', 'Home > Fill > Right on J16:L16: Item 1 copied across');
await page.keyboard.press('Control+z');
await page.keyboard.press('Control+z');
check(await value('J16') === null, 'Ctrl+Z twice: J16 empty again');

// The status bar: B2:B4 selected: Excel's Average, Count and Sum beside the sheet tabs; one cell, nothing.
await page.evaluate(() => window.oxe.select('B2'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Shift+ArrowDown');
const statusText = await page.textContent('.oxe-status');
const nums = [await value('B2'), await value('B3'), await value('B4')].filter((v) => typeof v === 'number');
const total = nums.reduce((x, y) => x + y, 0);
check(statusText.replace(/\s+/g, ' ') === `Average: ${total / nums.length} Count: ${nums.length} Sum: ${total}`, 'B2:B4 selected: Excel’s Average, Count and Sum of its numbers in the status bar', statusText);
await page.keyboard.press('ArrowDown');
check(!(await page.isVisible('.oxe-status')), 'one cell selected: the status bar says nothing');

// Clear: B11 (bold, a total) cleared with Home > Clear > Clear All: empty and plain; Ctrl+Z. B12:B13 (part of the merged
// A12:C13) and Delete: Excel's refusal.
await page.evaluate(() => window.oxe.select('B11'));
await page.focus('.oxe-grid');
const b11 = await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(10).get(1));
await menu('Clear', 'Clear All');
check(b11.s && await page.evaluate(() => { const c = window.oxe.workbook.sheet(0).cells.get(10)?.get(1); return !c || (c.v == null && !c.s); }), 'Clear All on B11: its value and format gone');
await page.keyboard.press('Control+z');
check(await page.evaluate((s) => window.oxe.workbook.sheet(0).cells.get(10).get(1).s === s, b11.s), 'Ctrl+Z: B11 back with its format');
await page.evaluate(() => window.oxe.select('B12'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Delete');
await page.waitForSelector('dialog.oxe-dialog');
check(await page.textContent('dialog.oxe-dialog') === "We can't do that to a merged cell.OK", 'Delete on B12:B13, part of the merged A12:C13: Excel’s refusal', await page.textContent('dialog.oxe-dialog'));
await page.keyboard.press('Enter');

// Sort: E2:E10 (Yes and No) selected, Sort & Filter > Sort Z to A: Excel's Sort Warning (data beside), Sort with the
// selection expanded: the Yes rows first, each row kept whole, the total under them left; Ctrl+Z. A1:F10, Custom
// Sort...: the headers guessed, Choice Z to A then Amount largest first; OK; Ctrl+Z.
const rowsOf = () => page.evaluate(() => Array.from({ length: 9 }, (_, k) => { const l = window.oxe.workbook.sheet(0).cells.get(k + 1); return `${l?.get(0)?.v}|${l?.get(1)?.v}|${l?.get(4)?.v}`; }));
const rowsWere = await rowsOf();
const totalLabel = await value('A11');
await page.evaluate(() => window.oxe.select('E2'));
await page.focus('.oxe-grid');
for (let k = 0; k < 8; k++) await page.keyboard.press('Shift+ArrowDown');
let warning = '';
await ribbon('Sort & Filter', async () => {
  await page.click('.oxe-pop >> text=Sort Z to A');
  await page.waitForSelector('dialog.oxe-dialog');
  warning = await page.textContent('dialog.oxe-dialog .oxe-dialog-text');
  await page.click('dialog.oxe-dialog >> button:text-is("Sort")');
});
let rowsNow = await rowsOf();
const yesFirst = rowsNow.findIndex((x) => x.endsWith('|No')) > rowsNow.findLastIndex((x) => x.endsWith('|Yes'));
check(/^There is data next to your selection/.test(warning) && yesFirst && [...rowsNow].sort().join() === [...rowsWere].sort().join() && await value('A11') === totalLabel,
  'Sort Z to A on E2:E10: Excel’s Sort Warning, the selection expanded, Yes rows first, rows whole, the total left', rowsNow.join(' '));
await page.keyboard.press('Control+z');
check((await rowsOf()).join() === rowsWere.join(), 'Ctrl+Z: the rows back in their order');
await page.evaluate(() => window.oxe.select('A1'));
await page.focus('.oxe-grid');
for (let k = 0; k < 5; k++) await page.keyboard.press('Shift+ArrowRight');
for (let k = 0; k < 9; k++) await page.keyboard.press('Shift+ArrowDown');
let headed = false;
await ribbon('Sort & Filter', async () => {
  await page.click('.oxe-pop >> text=Custom Sort...');
  await page.waitForSelector('dialog.oxe-sort-box');
  headed = await page.isChecked('dialog.oxe-sort-box >> input[type=checkbox]');
  await page.selectOption('.oxe-sort-level >> nth=0 >> select >> nth=0', { label: 'Choice' });
  await page.selectOption('.oxe-sort-level >> nth=0 >> select >> nth=1', { label: 'Z to A' });
  await page.click('dialog.oxe-sort-box >> button:text-is("Add Level")');
  await page.selectOption('.oxe-sort-level >> nth=1 >> select >> nth=0', { label: 'Amount' });
  await page.selectOption('.oxe-sort-level >> nth=1 >> select >> nth=1', { label: 'Largest to Smallest' });
  await page.click('dialog.oxe-sort-box >> button:text-is("OK")');
});
rowsNow = await rowsOf();
const byBoth = rowsNow.every((x, k) => k === 0 || (() => { const [, a, e] = rowsNow[k - 1].split('|'); const [, b, f] = x.split('|'); return e > f || (e === f && Number(a) >= Number(b)); })());
check(headed && byBoth && await value('A1') === 'Item', 'Custom Sort... on A1:F10: headers guessed, Choice Z to A then Amount largest first, the header row left', rowsNow.join(' '));
await page.keyboard.press('Control+z');
check((await rowsOf()).join() === rowsWere.join(), 'Ctrl+Z: as before');

// AutoSum: 5 and 7 typed in Q40:Q41; on Q42 Alt+= opens the cell with Excel's =SUM(Q40:Q41), Enter enters it.
// Q40:Q42 selected, AutoSum > Average: =AVERAGE(Q40:Q41) in the blank Q42. Ctrl+Z each.
await page.evaluate(() => window.oxe.select('Q40'));
await page.focus('.oxe-grid');
for (const n of ['5', '7']) { await page.keyboard.type(n); await page.keyboard.press('Enter'); }
await page.keyboard.press('Alt+Equal');
await page.waitForSelector('.oxe-editor:visible');
const proposed = await page.evaluate(() => document.querySelector('.oxe-editor').value);
await page.keyboard.press('Enter');
const q42 = () => page.evaluate(() => { const c = window.oxe.workbook.sheet(0).cells.get(41)?.get(16); return c && (c.f != null || c.v != null) ? `${c.f}=${c.v}` : ''; });
check(proposed === '=SUM(Q40:Q41)' && await q42() === 'SUM(Q40:Q41)=12', 'Alt+= under 5 and 7: Excel’s =SUM(Q40:Q41) proposed, Enter: 12', `${proposed} ${await q42()}`);
await page.keyboard.press('Control+z');
await page.evaluate(() => window.oxe.select('Q40'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Shift+ArrowDown');
await menu('AutoSum (Alt+=) options', 'Average');
check(await q42() === 'AVERAGE(Q40:Q41)=6', 'Q40:Q42, AutoSum > Average: =AVERAGE(Q40:Q41) in the blank Q42', await q42());
for (let k = 0; k < 3; k++) await page.keyboard.press('Control+z');
check(await q42() === '' && await value('Q40') === null, 'Ctrl+Z: Q40:Q42 empty again', `${await q42()} ${await value('Q40')} ${await value('Q41')}`);

// The Name Box: C3:D4 typed and Enter: that range selected, C3 active, back on the sheet; "1abc": Excel's words.
await page.click('.oxe-namebox');
await page.keyboard.type('c3:d4');
await page.keyboard.press('Enter');
const went = await page.evaluate(() => ({ sel: window.oxe.selection, box: document.querySelector('.oxe-namebox').value, grid: document.activeElement === window.oxe.grid.el }));
check(went.sel.r0 === 2 && went.sel.c0 === 2 && went.sel.r1 === 3 && went.sel.c1 === 3 && went.sel.ar === 2 && went.box === 'C3' && went.grid, 'Name Box, c3:d4 and Enter: C3:D4 selected, C3 active, back on the sheet', JSON.stringify(went));
await page.click('.oxe-namebox');
await page.keyboard.type('1abc');
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-dialog');
check(/^You must enter a valid reference/.test(await page.textContent('dialog.oxe-dialog')), 'Name Box, 1abc: Excel’s words', await page.textContent('dialog.oxe-dialog'));
await page.keyboard.press('Enter');
await page.click('.oxe-namebox');
await page.keyboard.press('Escape');
check(await page.evaluate(() => document.querySelector('.oxe-namebox').value) === 'C3', 'Esc: the Name Box shows the active cell again');

// Sizes: column B's edge dragged 40px wider, stored as Excel stores it; column A's edge double-clicked (AutoFit);
// Ctrl+Z; Home, Format, Row Height... 30 for row 3; Ctrl+9 hides it, Ctrl+Shift+9 shows it again.
const colPx = (c) => page.evaluate((c) => window.oxe.grid.geometry().cols.size(c), c);
const colOf = (c) => page.evaluate((c) => window.oxe.workbook.sheet(0).cols.find((x) => x.min <= c && x.max >= c) || null, c);
await page.evaluate(() => { window.oxe.select('A1'); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
const edgeX = (c) => page.evaluate((c) => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); return { x: v.left + geo.rowHeadW + geo.cols.start(c + 1) - 1, y: v.top + geo.colHeadH / 2 }; }, c);
const bWas = await colPx(1);
let edge = await edgeX(1);
await page.mouse.move(edge.x, edge.y);
await page.mouse.down();
await page.mouse.move(edge.x + 20, edge.y, { steps: 3 });
const tip = await page.textContent('.oxe-guide-tip');
await page.mouse.move(edge.x + 40, edge.y, { steps: 3 });
await page.mouse.up();
const bNow = await colPx(1);
check(bNow === bWas + 40 && (await colOf(1)).width === Math.trunc((bNow / 7) * 256) / 256 && /^Width: [\d.]+ \(\d+ pixels\)$/.test(tip), 'column B dragged 40px wider: its width stored in Excel’s units, Excel’s tip while dragging', `${bWas} -> ${bNow} ${tip}`);
await page.keyboard.press('Control+z');
check(await colPx(1) === bWas, 'Ctrl+Z: column B as it was');
edge = await edgeX(0);
await page.mouse.dblclick(edge.x, edge.y);
const colA = await colOf(0);
check(colA && colA.width > 0 && await colPx(0) > 30, 'column A’s edge double-clicked: AutoFit to its text', JSON.stringify(colA));
await page.keyboard.press('Control+z');
await page.evaluate(() => window.oxe.select('A3'));
await page.focus('.oxe-grid');
await ribbon('Format', async () => { await page.click('.oxe-pop >> text=Row Height...'); await page.fill('.oxe-dialog-field', '30'); await page.keyboard.press('Enter'); });
await page.waitForFunction(() => window.oxe.workbook.sheet(0).rows.get(2)?.ht === 30);
check(await page.evaluate(() => window.oxe.grid.geometry().rows.size(2)) === 40, 'Format, Row Height... 30: row 3 is 30pt (40 pixels)');
await page.focus('.oxe-grid');
await page.keyboard.press('Control+9');
check(await page.evaluate(() => window.oxe.workbook.sheet(0).rows.get(2)?.hidden === true && window.oxe.grid.geometry().rows.size(2) === 0), 'Ctrl+9: row 3 hidden');
await page.keyboard.press('Control+Shift+9');
check(await page.evaluate(() => window.oxe.grid.geometry().rows.size(2)) === 40, 'Ctrl+Shift+9: shown again, still 30pt');

// Find and Replace: Ctrl+F, "Item", Enter goes from A1 to the next match, then the next; Find All counts them. Ctrl+H,
// "Item" with "Thing", Replace All; Esc back to the sheet, Ctrl+Z puts them back.
await page.evaluate(() => window.oxe.select('A1'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+f');
await page.waitForSelector('.oxe-findbar:not([hidden])');
await page.keyboard.type('Item');
await page.keyboard.press('Enter');
const firstHit = await sel();
await page.keyboard.press('Enter');
const secondHit = await sel();
check(firstHit.ar === 1 && firstHit.ac === 0 && secondHit.ar === 2 && await page.evaluate(() => document.activeElement.classList.contains('oxe-findbox')), 'Ctrl+F, Item, Enter twice: A2 then A3, the box keeping focus', JSON.stringify([firstHit, secondHit]));
await page.click('.oxe-findbar button[title="Find All"]');
const found = await page.textContent('.oxe-findinfo');
check(/^\d+ cell\(s\) found$/.test(found), 'Find All: the cells found counted', found);
await page.keyboard.press('Control+h');
await page.fill('.oxe-replrow .oxe-findbox', 'Thing');
await page.click('.oxe-findbar button[title="Replace All"]');
const done = await page.textContent('.oxe-findinfo');
check(/^All done\. We made \d+ replacements?\./.test(done) && /^Thing/.test(String(await value('A3'))), 'Ctrl+H, Replace All: Item made Thing, said as Excel says it', done);
await page.keyboard.press('Escape');
check(await page.isHidden('.oxe-findbar') && await page.evaluate(() => document.activeElement.classList.contains('oxe-grid')), 'Esc: the bar closed, back in the sheet');
await page.keyboard.press('Control+z');
check(/^Item/.test(String(await value('A3'))), 'Ctrl+Z: every replacement undone at once');

// View: Freeze Panes at C3 from the View tab; scrolled far down and right, a click where A1 shows still selects A1.
// Ctrl+wheel zooms by 10%. Unfreeze Panes, and zoom back to 100%.
await page.evaluate(() => { window.oxe.select('C3'); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
await page.click('.oxe-rtab >> text=View');
await menu('Freeze Panes', 'Freeze Panes');
check(JSON.stringify(await page.evaluate(() => window.oxe.grid.view().freeze)) === '{"rows":2,"cols":2}' && await page.isVisible('.oxe-pane-top') && await page.isVisible('.oxe-pane-corner'), 'View, Freeze Panes at C3: rows 1-2 and columns A-B frozen');
await page.evaluate(() => { window.oxe.grid.el.scrollTop = 2000; window.oxe.grid.el.scrollLeft = 600; });
await page.waitForTimeout(100);
const a1 = await page.evaluate(() => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); return { x: v.left + geo.rowHeadW + 10, y: v.top + geo.colHeadH + 6 }; });
await page.mouse.click(a1.x, a1.y);
check((await sel()).ar === 0 && (await sel()).ac === 0, 'scrolled away, a click on the frozen corner selects A1', JSON.stringify(await sel()));
await page.mouse.move(a1.x + 200, a1.y + 200);
await page.keyboard.down('Control');
await page.mouse.wheel(0, -100);
await page.keyboard.up('Control');
await page.waitForFunction(() => window.oxe.grid.view().zoom === 110);
check(await page.evaluate(() => window.oxe.workbook.sheet(0).zoom) === 110, 'Ctrl+wheel: zoom 110%, kept for the file');
await menu('Freeze Panes', 'Unfreeze Panes');
await ribbon('100%');
check(JSON.stringify(await page.evaluate(() => window.oxe.grid.view().freeze)) === '{"rows":0,"cols":0}' && await page.evaluate(() => window.oxe.grid.view().zoom) === 100 && await page.isHidden('.oxe-pane-top'), 'Unfreeze Panes and 100%: as before');
await page.click('.oxe-rtab >> text=Home');

// Entries as Excel reads them: a date typed into G40 is a date, shown and in the formula bar as one; F2 and Enter keep
// it; $1,234.50 is currency; Ctrl+Z takes the value and its format away together.
const codeAt = (ref) => page.evaluate((ref) => { const wb = window.oxe.workbook; const cell = wb.sheet(0).cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65); const xf = wb.styles.cellXfs[cell?.s || 0]; return wb.styles.numFmts[xf.numFmtId] ?? ({ 14: 'm/d/yyyy', 0: 'General', 8: 'currency' })[xf.numFmtId] ?? String(xf.numFmtId); }, ref);
await page.evaluate(() => window.oxe.select('G40'));
await page.focus('.oxe-grid');
await page.keyboard.type('1/2/2024');
await page.keyboard.press('Enter');
check(await value('G40') === 45293 && await shown(39, 6) === '1/2/2024' && await codeAt('G40') === 'm/d/yyyy', 'typed 1/2/2024: the date 45293, shown 1/2/2024', `${await value('G40')} ${await shown(39, 6)}`);
await page.evaluate(() => window.oxe.select('G40'));
check(await bar() === '1/2/2024', 'the formula bar shows it as a date', await bar());
await page.focus('.oxe-grid');
await page.keyboard.press('F2');
await page.keyboard.press('Enter');
check(await value('G40') === 45293, 'F2 and Enter: still the same date');
await page.evaluate(() => window.oxe.select('G41'));
await page.focus('.oxe-grid');
await page.keyboard.type('$1,234.50');
await page.keyboard.press('Enter');
check(await value('G41') === 1234.5 && /\$"?#,##0\.00/.test(await codeAt('G41')), 'typed $1,234.50: 1234.5 as currency (shown ######## in a column this narrow, as in Excel)', await codeAt('G41'));
await page.keyboard.press('Control+z');
check(await value('G41') === null && await codeAt('G41') === 'General', 'Ctrl+Z: the value and its format gone together');

// Read-only: nothing changes.
await page.evaluate(() => window.oxe.setReadOnly(true));
await clickCell(1, 0);
await page.keyboard.type('nope');
await page.keyboard.press('Delete');
check(await value('A2') === 'Item 1 x' && await page.evaluate(() => document.querySelector('.oxe-grid').getAttribute('aria-readonly')) === 'true', 'read-only: typing and Delete change nothing');
check(await page.isHidden('.oxe-toolbar'), 'read-only: no ribbon');
await page.evaluate(() => window.oxe.setReadOnly(false));

// The saved workbook has the edits, and only its sheet changed.
const bytes = new Uint8Array(await page.evaluate(() => Array.from(window.oxe.save())));
fs.writeFileSync(`${outDir}/basic.typed.xlsx`, bytes);
const wb = Workbook.fromBytes(bytes);
const sh = wb.sheet(0);
const v = (r, c) => sh.cells.get(r)?.get(c)?.v ?? null;
check(v(1, 1) === 77 && v(1, 0) === 'Item 1 x' && v(15, 0) === 0.5 && v(15, 2) === '007' && v(15, 3) === true && v(1, 4) === 'No', 'saved: the edits are in the file');
check(sh.cells.get(15).get(4).f === 'SUM(B2:B3)' && v(15, 4) === 87 && v(2, 2) === 20 && v(10, 1) === 129, 'saved: the formula and the results worked out');
check(wb.styles.fills[wb.styles.cellXfs[sh.cells.get(19).get(7).s].fillId].fg.rgb === 'FFFFFF00', 'saved: H20 filled yellow');
check(!wb.pkg.has('xl/calcChain.xml'), 'saved: calcChain.xml left out (Excel works the order out again)');
check(wb.sheet(1).cells.get(1).get(1).v === -3.5, 'saved: the change on the second sheet undone');
check(wb.sheets.map((s) => `${s.name}:${s.state}`).join() === 'Data:visible,Notes:visible,Hidden:hidden' && !Object.keys(wb.pkg.files).some((p) => /sheet[45]\.xml/.test(p)), 'saved: the sheets as left (Summary and Sheet2 gone with their parts)', wb.sheets.map((s) => s.name).join());

check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
