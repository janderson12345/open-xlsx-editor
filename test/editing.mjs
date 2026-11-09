// Browser check of editing: selection by keys and mouse, typing over a cell, F2, Esc, Tab, Delete, undo and redo,
// typed numbers and text, formula cells refused, read-only; the saved workbook read back. Standards mode asserted.
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

// Formula cells are refused until there is a calculation engine.
await clickCell(2, 2); // C3, a shared formula's cell
await page.keyboard.type('5');
await page.waitForTimeout(250);
check(await value('C3') === 4 && /formula/.test(await live()), 'typing over a formula cell is refused and said so', await live());

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

check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
