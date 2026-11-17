// Browser check of where the items drawn over the cells sit: each filter button inside its heading cell's right edge,
// each note's red mark in its cell's top right corner, a list rule's button just right of its cell, also with panes
// frozen (whose left panes count from column A without the row headings). They were all a row heading's width too far
// left before (a workbook with Cambria's wider columns showed the filter buttons over the headings' text).
// Run after `npm run build`: node test/overlay-places.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish } = await editingPage();

/** For each element matching `sel`, how far its edge (`side`: 'left' or 'right') is from the nearest column boundary,
 *  in screen pixels, signed: 0 on a boundary, -1 a pixel left of one. */
const offsets = (sel, side) => page.evaluate(({ sel, side }) => {
  const g = window.oxe.grid;
  const box = g.el.getBoundingClientRect();
  const { cols, rowHeadW } = g.geometry();
  const bounds = Array.from({ length: 40 }, (_, c) => box.left + rowHeadW + cols.start(c) - g.el.scrollLeft);
  return [...document.querySelectorAll(sel)].filter((e) => e.getClientRects().length).map((e) => {
    const x = e.getBoundingClientRect()[side];
    const near = bounds.reduce((a, b) => (Math.abs(x - b) < Math.abs(x - a) ? b : a));
    return Math.round(x - near);
  });
}, { sel, side });

// basic.xlsx: a filter on its list, and the list rule at E3.
await openFixture(page, 'basic.xlsx');
await page.evaluate(() => { window.oxe.select('A1'); window.oxe.command('filterToggle'); });
await page.waitForSelector('.oxe-filter-btn');
const filters = await offsets('.oxe-grid > .oxe-sheet .oxe-filters .oxe-filter-btn', 'right');
check(filters.length > 0 && filters.every((d) => d === -1), 'filter buttons: each ends a pixel inside its heading cell\'s right edge', JSON.stringify(filters));
await page.evaluate(() => window.oxe.select('E3'));
await page.focus('.oxe-grid');
await page.waitForSelector('.oxe-dv-btn');
const list = await offsets('.oxe-dv-btn', 'left');
check(list.length === 1 && list[0] === 1, 'a list rule\'s button: a pixel right of its cell', JSON.stringify(list));

// Frozen at C3: the heading's A and B in the corner pane, the rest in the top pane, each button still at its edge.
await page.evaluate(() => { window.oxe.select('C3'); window.oxe.command('freeze', 'panes'); });
await page.waitForTimeout(200);
const frozen = await offsets('.oxe-filter-btn', 'right');
check(frozen.length > 0 && frozen.every((d) => d === -1), 'with panes frozen at C3: every filter button, in each pane, at its cell\'s right edge', JSON.stringify(frozen));
await page.evaluate(() => window.oxe.command('freeze', 'none'));

// notes.xlsx: notes on B2, C3 and D4, each mark's right side on its cell's right edge.
await openFixture(page, 'notes.xlsx');
await page.waitForSelector('.oxe-note-mark');
const marks = await offsets('.oxe-note-mark', 'right');
check(marks.length === 3 && marks.every((d) => d === 0), 'note marks: in their cells\' top right corners (B2, C3, D4)', JSON.stringify(marks));

await finish();
