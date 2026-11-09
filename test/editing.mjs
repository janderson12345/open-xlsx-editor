// Browser check of editing: selection by keys and mouse, typing over a cell, F2, Esc, Tab, Delete, undo and redo,
// typed numbers and text, formulas typed and worked out again, sheet tabs, read-only; the saved workbook read back.
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

// Read-only: nothing changes.
await page.evaluate(() => window.oxe.setReadOnly(true));
await clickCell(1, 0);
await page.keyboard.type('nope');
await page.keyboard.press('Delete');
check(await value('A2') === 'Item 1 x' && await page.evaluate(() => document.querySelector('.oxe-grid').getAttribute('aria-readonly')) === 'true', 'read-only: typing and Delete change nothing');
await page.evaluate(() => window.oxe.setReadOnly(false));

// The saved workbook has the edits, and only its sheet changed.
const bytes = new Uint8Array(await page.evaluate(() => Array.from(window.oxe.save())));
fs.writeFileSync(`${outDir}/basic.typed.xlsx`, bytes);
const wb = Workbook.fromBytes(bytes);
const sh = wb.sheet(0);
const v = (r, c) => sh.cells.get(r)?.get(c)?.v ?? null;
check(v(1, 1) === 77 && v(1, 0) === 'Item 1 x' && v(15, 0) === 0.5 && v(15, 2) === '007' && v(15, 3) === true && v(1, 4) === 'No', 'saved: the edits are in the file');
check(sh.cells.get(15).get(4).f === 'SUM(B2:B3)' && v(15, 4) === 87 && v(2, 2) === 20 && v(10, 1) === 129, 'saved: the formula and the results worked out');
check(!wb.pkg.has('xl/calcChain.xml'), 'saved: calcChain.xml left out (Excel works the order out again)');
check(wb.sheet(1).cells.get(1).get(1).v === -3.5, 'saved: the change on the second sheet undone');

check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
