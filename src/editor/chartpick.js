// Charts, pictures, shapes and groups picked out on the sheet, as in Excel: a click on one selects it (outlined, with
// eight handles), Ctrl+click or Shift+click adds it or takes it away; dragged they move, a handle dragged sizes that
// one, kept to the anchor's cells and offsets on letting go; Delete deletes them, Esc or a click on a cell lets them
// go. A chart's menu (right click, Shift+F10 or the menu key) has Edit Chart..., Change Chart Type... and Delete; an
// object's has Delete. From the keyboard, Ctrl+Alt+5 picks the sheet's first object and Tab the next; the arrow keys
// move what is picked a column or row at a time. Home > Find & Select > Go To Special > Objects picks them all
// (pickAll). Each change is one step of the history (acts-more.js chartAnchor, chartDelete, drawingsPlace,
// drawingsDelete). Objects are known by their place among the sheet's anchors (`ord`: drawobjects.js readObjects's k;
// a chart's is set on it by index.js); a chart's own commands go by its chart index.
import { announce } from './live.js';
import { EMU_PER_PX } from '../ooxml/units.js';
import { anchorRect } from '../ooxml/drawobjects.js';

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const BOXES = '.oxe-chart, .oxe-obj';

/** The picker for a grid. `grid` is grid.js's; `charts()` the sheet's charts as drawn (with `index`, readCharts's,
 *  and `ord`); `objects()` its other drawing objects (readObjects's, drawn); `act(name, ...args)` runs a command;
 *  `menu(point, k)` opens a chart's menu, `objectMenu(point, ords)` an object's; `readOnly()`. */
export function chartPicker({ grid, charts, objects = () => [], act, menu, objectMenu = null, readOnly }) {
  const view = grid.el;
  let picked = []; // the objects' ords, in the order picked
  let drag = null;
  /** Everything that can be picked, in the drawing's order: { ord, chart (its index, for a chart), name, item }. */
  const items = () => [
    ...charts().map((c) => ({ ord: c.ord ?? -1 - c.index, chart: c.index, name: c.name || 'Chart', c })),
    ...objects().map((o) => ({ ord: o.k, chart: null, name: o.name || 'Object', o })),
  ].sort((a, b) => a.ord - b.ord);
  const itemOf = (ord) => items().find((x) => x.ord === ord) || null;
  const boxesOf = (ord) => [...view.querySelectorAll(`.oxe-chart[data-oxe-ord="${ord}"], .oxe-obj[data-oxe-ord="${ord}"]`)];
  /** The picked objects' outlines and handles (again after a redraw). */
  const paint = () => {
    for (const el of view.querySelectorAll('.oxe-chart-on')) { el.classList.remove('oxe-chart-on'); el.querySelectorAll('.oxe-chart-h').forEach((h) => h.remove()); }
    for (const ord of picked) {
      for (const box of boxesOf(ord)) {
        box.classList.add('oxe-chart-on');
        for (const h of HANDLES) box.append(Object.assign(document.createElement('span'), { className: `oxe-chart-h oxe-chart-h-${h}` }));
        box.querySelectorAll('.oxe-chart-h').forEach((h, i) => { h.dataset.h = HANDLES[i]; });
      }
    }
  };
  const say = () => {
    if (!picked.length) return;
    if (picked.length > 1) { announce(`${picked.length} objects selected. Delete deletes them, the arrow keys move them, Esc goes back to the cells.`); return; }
    announce(`${itemOf(picked[0])?.name || 'Object'} selected. Delete deletes it, the arrow keys move it, Shift+F10 opens its menu, Esc goes back to the cells.`);
  };
  const setPicked = (ords, tell = true) => { picked = [...new Set(ords)]; paint(); if (tell) say(); };
  /** Chart k (its chart index) picked, or nothing (null). */
  const pick = (k, tell = true) => {
    if (k == null) { setPicked([], false); return; }
    const it = items().find((x) => x.chart === k);
    setPicked(it ? [it.ord] : [], tell);
  };
  /** An object's rectangle on the sheet, in pixels. */
  const rectOf = (it) => {
    const { rows, cols } = grid.geometry();
    if (it.o) return anchorRect(it.o.anchor, rows, cols, it.o.item);
    const ch = it.c;
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
  /** Objects put at new rectangles ([{ it, rect }]), as one step. */
  const place = (list) => {
    const keep = [...picked];
    if (list.length === 1 && list[0].it.chart != null) act('chartAnchor', list[0].it.chart, corner(list[0].rect.x, list[0].rect.y), corner(list[0].rect.x + list[0].rect.w, list[0].rect.y + list[0].rect.h));
    else {
      act('drawingsPlace', list.map(({ it, rect }) => ({
        k: it.ord, chart: it.chart, from: corner(rect.x, rect.y), to: corner(rect.x + rect.w, rect.y + rect.h),
        rect: { x: Math.max(0, rect.x) * EMU_PER_PX, y: Math.max(0, rect.y) * EMU_PER_PX, w: rect.w * EMU_PER_PX, h: rect.h * EMU_PER_PX },
      })));
    }
    setPicked(keep, false);
  };
  const remove = () => {
    const list = picked.map(itemOf).filter(Boolean);
    if (!list.length || readOnly()) return;
    setPicked([], false);
    if (list.length === 1 && list[0].chart != null) act('chartDelete', list[0].chart);
    else act('drawingsDelete', list.map((it) => it.ord));
    announce(list.length > 1 ? `${list.length} objects deleted` : `${list[0].name} deleted`);
  };

  view.addEventListener('mousedown', (e) => {
    const box = e.target.closest?.(BOXES);
    if (!box) { if (picked.length) setPicked([], false); return; }
    e.preventDefault();
    e.stopPropagation();
    const ord = Number(box.dataset.oxeOrd);
    if (e.ctrlKey || e.metaKey || e.shiftKey) {
      setPicked(picked.includes(ord) ? picked.filter((x) => x !== ord) : [...picked, ord]);
      view.focus({ preventScroll: true });
      return;
    }
    if (!picked.includes(ord)) setPicked([ord]);
    view.focus({ preventScroll: true });
    const it = itemOf(ord);
    if (e.button === 2) {
      if (it?.chart != null && picked.length === 1) menu({ x: e.clientX, y: e.clientY }, it.chart);
      else objectMenu?.({ x: e.clientX, y: e.clientY }, [...picked]);
      return;
    }
    if (readOnly() || !it) return;
    const handle = e.target.dataset?.h || null;
    // (A handle sizes its own object; the object dragged moves all those picked with it.)
    const moving = handle ? [it] : picked.map(itemOf).filter(Boolean);
    drag = { handle, x: e.clientX, y: e.clientY, list: moving.map((m) => ({ it: m, rect: rectOf(m), boxes: boxesOf(m.ord) })), box, moved: false };
  }, true);
  // (The right click's own contextmenu event, on Windows after the button is let go: the object's menu is open
  // already, so neither the browser's menu nor the grid's cell menu opens over it.)
  view.addEventListener('contextmenu', (e) => {
    if (!e.target.closest?.(BOXES)) return;
    e.preventDefault();
    e.stopImmediatePropagation();
  }, true);
  const onMove = (e) => {
    if (!drag) return;
    const z = (grid.view().zoom || 100) / 100;
    const dx = (e.clientX - drag.x) / z;
    const dy = (e.clientY - drag.y) / z;
    if (Math.abs(dx) + Math.abs(dy) > 2) drag.moved = true;
    const h = drag.handle;
    for (const d of drag.list) {
      const r = { ...d.rect };
      if (!h) { r.x += dx; r.y += dy; }
      else {
        if (h.includes('w')) { r.x += Math.min(dx, r.w - 10); r.w -= Math.min(dx, r.w - 10); }
        if (h.includes('e')) r.w = Math.max(10, r.w + dx);
        if (h.includes('n')) { r.y += Math.min(dy, r.h - 10); r.h -= Math.min(dy, r.h - 10); }
        if (h.includes('s')) r.h = Math.max(10, r.h + dy);
      }
      d.now = r;
      // (Shown while dragging: the objects moved, or an outline of the new size.)
      for (const b of d.boxes) {
        b.style.transform = h ? '' : `translate(${r.x - d.rect.x}px, ${r.y - d.rect.y}px)`;
        if (h) { b.style.setProperty('--oxe-ghost-w', `${r.w}px`); b.style.setProperty('--oxe-ghost-h', `${r.h}px`); b.style.setProperty('--oxe-ghost-x', `${r.x - d.rect.x}px`); b.style.setProperty('--oxe-ghost-y', `${r.y - d.rect.y}px`); b.classList.add('oxe-chart-sizing'); }
      }
    }
  };
  const onUp = () => {
    if (!drag) return;
    const d = drag;
    drag = null;
    for (const x of d.list) for (const b of x.boxes) { b.style.transform = ''; b.classList.remove('oxe-chart-sizing'); }
    if (d.moved && d.list.every((x) => x.now)) place(d.list.map((x) => ({ it: x.it, rect: x.now })));
  };
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp);

  /** A key while the grid has focus: true when it was the picker's. */
  const onKey = (e) => {
    const ctrl = e.ctrlKey || e.metaKey;
    const list = items();
    if (ctrl && e.altKey && (e.key === '5' || e.code === 'Digit5')) { if (list.length) { setPicked([list[0].ord]); return true; } announce('No objects on this sheet'); return true; }
    picked = picked.filter((ord) => list.some((x) => x.ord === ord));
    if (!picked.length) return false;
    const first = list.find((x) => x.ord === picked[0]);
    if (e.key === 'Escape') { setPicked([], false); announce('Back to the cells'); return true; }
    if (e.key === 'Tab') { const i = list.indexOf(first); setPicked([list[(i + (e.shiftKey ? list.length - 1 : 1)) % list.length].ord]); return true; }
    if ((e.key === 'Delete' || e.key === 'Backspace') && !readOnly()) { remove(); return true; }
    if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') {
      const b = boxesOf(first.ord)[0]?.getBoundingClientRect();
      const point = b ? { x: b.left + 20, y: b.top + 20 } : { x: 100, y: 100 };
      if (first.chart != null && picked.length === 1) menu(point, first.chart, true); else objectMenu?.(point, [...picked], true);
      return true;
    }
    const step = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (step && !readOnly()) {
      const { rows, cols } = grid.geometry();
      const moves = picked.map((ord) => list.find((x) => x.ord === ord)).filter(Boolean).map((it) => {
        const r = rectOf(it);
        const c = cols.indexAt(Math.max(0, r.x));
        const rr = rows.indexAt(Math.max(0, r.y));
        const dx = step[0] ? (step[0] > 0 ? cols.size(c) : -cols.size(Math.max(0, c - 1))) : 0;
        const dy = step[1] ? (step[1] > 0 ? rows.size(rr) : -rows.size(Math.max(0, rr - 1))) : 0;
        return { it, rect: { ...r, x: r.x + dx, y: r.y + dy } };
      });
      if (moves.some((m) => m.rect.x < 0 || m.rect.y < 0)) return true;
      place(moves);
      return true;
    }
    if (first.chart != null && picked.length === 1) {
      if (e.key === 'Enter') { act('chartBox', first.chart); return true; }
      // (Ctrl+1, Excel's Format pane for what is picked: the Format Chart box.)
      if (e.key === '1' && ctrl) { act('chartFormatBox', first.chart); return true; }
    }
    return !e.key.startsWith('F') && e.key.length === 1; // (typing doesn't reach the cells while an object is picked)
  };
  return {
    onKey,
    /** The drawings were drawn again: the picked ones' outlines put back, those gone let go. */
    redrawn() { const list = items(); picked = picked.filter((ord) => list.some((x) => x.ord === ord)); paint(); },
    /** The chart picked (its chart index), when one chart alone is. */
    get picked() { const it = picked.length === 1 ? itemOf(picked[0]) : null; return it?.chart ?? null; },
    /** What is picked: the objects' ords. */
    get pickedOrds() { return [...picked]; },
    pick,
    /** Go To Special > Objects: every chart, picture, shape and group on the sheet picked; false when there are none. */
    pickAll() { const list = items(); if (!list.length) return false; setPicked(list.map((x) => x.ord)); return true; },
    destroy() { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); },
  };
}
