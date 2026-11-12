// Formulas > Watch Window: Excel's list of watched cells (ooxml/watches.js), each with its Book, Sheet, Name (a defined
// name for just that cell), Cell, Value (as the cell shows it) and Formula, kept up to date as the workbook changes;
// Add Watch watches the selected cells (each cell of a range, as Excel's does), Delete Watch the watch picked; a double
// click goes to its cell. Docked under the ribbon here, as Excel's can be.
import { el } from './ribbonkit.js';
import { readWatches, addWatches, deleteWatches } from '../ooxml/watches.js';
import { refName } from '../ooxml/refs.js';
import { cellText } from '../ooxml/numfmt.js';
import { formatCodeOf } from '../ooxml/styles.js';
import { areasOf } from './selection.js';

const COLUMNS = ['Book', 'Sheet', 'Name', 'Cell', 'Value', 'Formula'];
/** A defined name standing for just cell (r, c) of sheet `name`. */
function nameOf(wb, sheetName, r, c) {
  const want = `${sheetName}!${refName(r, c)}`.toUpperCase();
  const n = wb.definedNames.find((d) => d.text.replace(/\$/g, '').replace(/^'(.*)'!/, '$1!').toUpperCase() === want);
  return n ? n.name : '';
}

/** The Watch Window. `ctx` = { workbook(), bookName(), sheetIndex(), selection(), go({ si, r, c }), readOnly(),
 *  changed() (a watch added or deleted), home() (focus back to the grid) }. */
export function createWatchWindow(ctx) {
  const body = el('tbody');
  const table = el('table', { class: 'oxe-watch-table', 'aria-label': 'Watched cells' },
    el('thead', {}, el('tr', {}, ...COLUMNS.map((h) => el('th', { scope: 'col' }, h)))), body);
  let picked = -1;
  let watches = [];
  const add = el('button', { type: 'button', class: 'oxe-watch-btn' }, 'Add Watch');
  const del = el('button', { type: 'button', class: 'oxe-watch-btn' }, 'Delete Watch');
  const close = el('button', { type: 'button', class: 'oxe-watch-btn oxe-find-close', 'aria-label': 'Close the Watch Window', title: 'Close' }, '✕');
  const bar = el('section', { class: 'oxe-watchbar', 'aria-label': 'Watch Window', hidden: '' },
    el('div', { class: 'oxe-watch-head' }, el('strong', {}, 'Watch Window'), add, del, close), el('div', { class: 'oxe-watch-scroll' }, table));
  const refresh = () => {
    if (bar.hidden) return;
    const wb = ctx.workbook();
    if (!wb) return;
    watches = readWatches(wb);
    if (picked >= watches.length) picked = watches.length - 1;
    body.replaceChildren(...watches.map((w, k) => {
      const sheet = wb.sheets[w.si].name;
      const cell = wb.sheet(w.si).cells.get(w.r)?.get(w.c);
      const value = cell ? cellText(cell, Infinity, formatCodeOf(wb.styles, cell.s || 0), wb.date1904) : '';
      const formula = cell && cell.f != null && cell.f !== '' ? `=${cell.f}` : '';
      const row = el('tr', { tabindex: '0', 'aria-selected': String(k === picked) },
        ...[ctx.bookName(), sheet, nameOf(wb, sheet, w.r, w.c), refName(w.r, w.c), value, formula].map((t) => el('td', {}, t)));
      row.addEventListener('click', () => { picked = k; refresh(); body.children[k]?.focus(); });
      row.addEventListener('dblclick', () => ctx.go(w));
      row.addEventListener('keydown', (e) => { if (e.key === 'Enter') ctx.go(w); else if (e.key === 'Delete') del.click(); });
      return row;
    }));
    del.disabled = picked < 0 || ctx.readOnly();
    add.disabled = ctx.readOnly();
  };
  add.addEventListener('click', () => {
    const wb = ctx.workbook();
    const si = ctx.sheetIndex();
    const cells = [];
    for (const g of areasOf(ctx.selection())) {
      if ((g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1) > 1000) continue;
      for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) cells.push({ r, c });
    }
    if (addWatches(wb, si, cells)) ctx.changed();
    refresh();
  });
  del.addEventListener('click', () => {
    if (picked < 0 || !watches[picked]) return;
    deleteWatches(ctx.workbook(), [watches[picked]]);
    ctx.changed();
    refresh();
  });
  close.addEventListener('click', () => { bar.hidden = true; ctx.home(); });
  return {
    el: bar,
    /** Shown, or hidden again (Formulas > Watch Window turns it over). */
    toggle() { bar.hidden = !bar.hidden; picked = -1; refresh(); if (!bar.hidden) add.focus(); else ctx.home(); },
    refresh,
    get isOpen() { return !bar.hidden; },
  };
}
