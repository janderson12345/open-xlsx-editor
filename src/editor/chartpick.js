// Charts picked out on the sheet, as in Excel: a click on a chart selects it (outlined, with eight handles); dragged it
// moves, a handle dragged sizes it, both kept to the anchor's cells and offsets on letting go; Delete deletes it, Esc
// or a click on a cell lets it go, and its menu (right click, Shift+F10 or the menu key) has Edit Chart..., Change
// Chart Type... and Delete. From the keyboard, Ctrl+Alt+5 picks the sheet's first chart and Tab the next; the arrow
// keys move a picked chart a column or row at a time. Each change is one step of the history (acts.js chartAnchor,
// chartDelete).
import { announce } from './live.js';
import { EMU_PER_PX } from '../ooxml/units.js';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];

/** The picker for a grid. `grid` is grid.js's; `charts()` the sheet's charts as drawn (with `index`, readCharts's);
 *  `act(name, ...args)` runs a command; `menu(point, k)` opens a chart's menu; `readOnly()`. */
export function chartPicker({ grid, charts, act, menu, readOnly }) {
  const view = grid.el;
  let picked = null; // the chart's index (readCharts's), or null
  let drag = null;
  const nameOf = (k) => charts().find((c) => c.index === k)?.name || 'Chart';
  const boxOf = (k) => view.querySelector(`.oxe-chart[data-oxe-chart="${k}"]`);
  /** The picked chart's outline and handles (again after a redraw). */
  const paint = () => {
    for (const el of view.querySelectorAll('.oxe-chart.oxe-chart-on')) { el.classList.remove('oxe-chart-on'); el.querySelectorAll('.oxe-chart-h').forEach((h) => h.remove()); }
    if (picked == null) return;
    const box = boxOf(picked);
    if (!box) return;
    box.classList.add('oxe-chart-on');
    for (const h of HANDLES) box.append(Object.assign(document.createElement('span'), { className: `oxe-chart-h oxe-chart-h-${h}` }));
    box.querySelectorAll('.oxe-chart-h').forEach((h, i) => { h.dataset.h = HANDLES[i]; });
  };
  const pick = (k, say = true) => {
    picked = k;
    paint();
    if (k != null && say) announce(`${nameOf(k)} selected. Delete deletes it, the arrow keys move it, Shift+F10 opens its menu, Esc goes back to the cells.`);
  };
  /** A chart's rectangle on the sheet, in pixels. */
  const rectOf = (ch) => {
    const { rows, cols } = grid.geometry();
    const x = cols.start(ch.from.c) + ch.from.dx;
    const y = rows.start(ch.from.r) + ch.from.dy;
    return { x, y, w: cols.start(ch.to.c) + ch.to.dx - x, h: rows.start(ch.to.r) + ch.to.dy - y };
  };
  /** A point on the sheet as an anchor corner: its cell and the offset into it (EMU). */
  const corner = (x, y) => {
    const { rows, cols } = grid.geometry();
    const c = cols.indexAt(Math.max(0, x));
    const r = rows.indexAt(Math.max(0, y));
    return { col: c, colOff: Math.round((Math.max(0, x) - cols.start(c)) * EMU_PER_PX), row: r, rowOff: Math.round((Math.max(0, y) - rows.start(r)) * EMU_PER_PX) };
  };
  const place = (k, rect) => { act('chartAnchor', k, corner(rect.x, rect.y), corner(rect.x + rect.w, rect.y + rect.h)); pick(k, false); };

  view.addEventListener('mousedown', (e) => {
    const box = e.target.closest?.('.oxe-chart');
    if (!box) { if (picked != null) pick(null, false); return; }
    e.preventDefault();
    e.stopPropagation();
    const k = Number(box.dataset.oxeChart);
    if (picked !== k) pick(k);
    view.focus({ preventScroll: true });
    if (e.button === 2) { menu({ x: e.clientX, y: e.clientY }, k); return; }
    if (readOnly()) return;
    const ch = charts().find((c) => c.index === k);
    drag = { k, handle: e.target.dataset?.h || null, x: e.clientX, y: e.clientY, rect: rectOf(ch), box, moved: false };
  }, true);
  // (The right click's own contextmenu event, on Windows after the button is let go: the chart's menu is open already,
  // so neither the browser's menu nor the grid's cell menu opens over it.)
  view.addEventListener('contextmenu', (e) => {
    if (!e.target.closest?.('.oxe-chart')) return;
    e.preventDefault();
    e.stopImmediatePropagation();
  }, true);
  const onMove = (e) => {
    if (!drag) return;
    const z = (grid.view().zoom || 100) / 100;
    const dx = (e.clientX - drag.x) / z;
    const dy = (e.clientY - drag.y) / z;
    if (Math.abs(dx) + Math.abs(dy) > 2) drag.moved = true;
    const r = { ...drag.rect };
    const h = drag.handle;
    if (!h) { r.x += dx; r.y += dy; }
    else {
      if (h.includes('w')) { r.x += Math.min(dx, r.w - 10); r.w -= Math.min(dx, r.w - 10); }
      if (h.includes('e')) r.w = Math.max(10, r.w + dx);
      if (h.includes('n')) { r.y += Math.min(dy, r.h - 10); r.h -= Math.min(dy, r.h - 10); }
      if (h.includes('s')) r.h = Math.max(10, r.h + dy);
    }
    drag.now = r;
    // (Shown while dragging: the chart moved, or an outline of its new size.)
    drag.box.style.transform = h ? '' : `translate(${(r.x - drag.rect.x)}px, ${(r.y - drag.rect.y)}px)`;
    drag.box.style.outlineOffset = '';
    if (h) { drag.box.style.setProperty('--oxe-ghost-w', `${r.w}px`); drag.box.style.setProperty('--oxe-ghost-h', `${r.h}px`); drag.box.style.setProperty('--oxe-ghost-x', `${r.x - drag.rect.x}px`); drag.box.style.setProperty('--oxe-ghost-y', `${r.y - drag.rect.y}px`); drag.box.classList.add('oxe-chart-sizing'); }
  };
  const onUp = () => {
    if (!drag) return;
    const d = drag;
    drag = null;
    d.box.style.transform = '';
    d.box.classList.remove('oxe-chart-sizing');
    if (d.moved && d.now) place(d.k, d.now);
  };
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp);

  /** A key while the grid has focus: true when it was the picker's. */
  const onKey = (e) => {
    const ctrl = e.ctrlKey || e.metaKey;
    const list = charts();
    if (ctrl && e.altKey && (e.key === '5' || e.code === 'Digit5')) { if (list.length) { pick(list[0].index); return true; } announce('No charts on this sheet'); return true; }
    if (picked == null) return false;
    const ch = list.find((c) => c.index === picked);
    if (!ch) { picked = null; return false; }
    if (e.key === 'Escape') { pick(null, false); announce('Back to the cells'); return true; }
    if (e.key === 'Tab') { const i = list.indexOf(ch); pick(list[(i + (e.shiftKey ? list.length - 1 : 1)) % list.length].index); return true; }
    if ((e.key === 'Delete' || e.key === 'Backspace') && !readOnly()) { const k = picked; pick(null, false); act('chartDelete', k); announce(`${ch.name} deleted`); return true; }
    if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') { const b = boxOf(picked)?.getBoundingClientRect(); menu(b ? { x: b.left + 20, y: b.top + 20 } : { x: 100, y: 100 }, picked, true); return true; }
    const step = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (step && !readOnly()) {
      const { rows, cols } = grid.geometry();
      const r = rectOf(ch);
      const dx = step[0] ? (step[0] > 0 ? cols.size(ch.from.c) : -cols.size(Math.max(0, ch.from.c - 1))) : 0;
      const dy = step[1] ? (step[1] > 0 ? rows.size(ch.from.r) : -rows.size(Math.max(0, ch.from.r - 1))) : 0;
      if (r.x + dx < 0 || r.y + dy < 0) return true;
      place(picked, { ...r, x: r.x + dx, y: r.y + dy });
      return true;
    }
    if (e.key === 'Enter') { act('chartBox', picked); return true; }
    return !e.key.startsWith('F') && e.key.length === 1; // (typing doesn't reach the cells while a chart is picked)
  };
  return {
    onKey,
    /** The charts were drawn again: the picked one's outline put back, or let go when it's gone. */
    redrawn() { if (picked != null && !charts().some((c) => c.index === picked)) picked = null; paint(); },
    get picked() { return picked; },
    pick,
    destroy() { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); },
  };
}
