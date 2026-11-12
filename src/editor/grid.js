// The sheet grid, virtualised: only the rows and columns in view (and a few around them) are in the DOM, redrawn from
// the sparse cell store as the view scrolls. The scroll position is the one source of truth for what is drawn.
//
// Layout: .oxe-grid scrolls; inside it .oxe-sheet is as large as the whole grid. The column headings are one row stuck
// to the top (position: sticky), each row's heading is stuck to the left inside its row, and cells sit at their sheet
// coordinates. Geometry goes through custom properties (--oxe-x, --oxe-y, --oxe-w, --oxe-h), never through styles
// written for state.
import { setVars, div } from './dom.js';
import { sheetGeometry, MDW, COLHEAD_H, OVERSCAN } from './geometry.js';
import { drawColumnHeadings, drawPanes, drawGridlines, drawRows } from './gridlayers.js';
import { columnFitWidth } from './cells.js';
import { decorations } from './decor.js';
import { autoRowHeights, rowDescents } from './rowheights.js';
import { drawSelection, drawMarks, rangeBox, handleCorner } from './selection.js';
import { arrowLayer, drawArrows } from './arrows.js';
import { drawPeers } from './peers.js';
import { drawFilterButtons } from './filterbuttons.js';
import { drawNoteMarks } from './notemarks.js';
import { drawValidation } from './dvmarks.js';
import { drawCharts } from './chartview.js';
import { cellDisplay } from '../ooxml/numfmt.js';
import { colChars } from '../ooxml/units.js';


/** The range of rows and columns in view (with overscan), from the scroll position. */
function visibleWindow(st) {
  const { view, rows, cols, rowHeadW } = st;
  const z = (st.zoom || 100) / 100;
  const top = view.scrollTop / z;
  const left = view.scrollLeft / z;
  const w = Math.max(0, view.clientWidth / z - rowHeadW);
  const h = Math.max(0, view.clientHeight / z - COLHEAD_H);
  return {
    r0: Math.max(0, rows.indexAt(top) - OVERSCAN),
    r1: Math.min(rows.count - 1, rows.indexAt(top + h) + OVERSCAN),
    c0: Math.max(0, cols.indexAt(left) - OVERSCAN),
    c1: Math.min(cols.count - 1, cols.indexAt(left + w) + OVERSCAN),
  };
}

function draw(st, force = false) {
  const win = visibleWindow(st);
  const last = st.win;
  if (!force && last && last.r0 === win.r0 && last.r1 === win.r1 && last.c0 === win.c0 && last.c1 === win.c1) return;
  st.win = win;
  drawColumnHeadings(st, win);
  drawGridlines(st, win);
  const { fills, edges } = decorations(st, win);
  st.fills.replaceChildren(...fills);
  drawRows(st, win);
  st.edges.replaceChildren(...edges);
  drawPeers(st);
  drawSelection(st);
  drawMarks(st);
  drawArrows(st);
  drawFilterButtons(st, win);
  drawErrorButton(st);
  drawNoteMarks(st, win);
  drawCharts(st, win);
  drawValidation(st);
  if (st.panes) drawPanes(st, win);
  const active = st.sel && st.idBase ? document.getElementById(`${st.idBase}-${st.sel.ar}-${st.sel.ac}`) : null;
  if (active) st.view.setAttribute('aria-activedescendant', active.id);
  else st.view.removeAttribute('aria-activedescendant');
}

/** Excel's error button, beside the active cell when it has an error indicator: left of it, or right in column A. */
function drawErrorButton(st) {
  const layer = st.errLayer;
  if (!layer) return;
  const s = st.sel;
  if (!s || !st.errors?.has(`${s.ar},${s.ac}`) || !st.rows.size(s.ar)) { layer.replaceChildren(); return; }
  const b = document.createElement('button');
  b.type = 'button';
  b.tabIndex = -1;
  b.className = 'oxe-err-btn';
  b.textContent = '!';
  b.title = 'Error options (Alt+Shift+F10)';
  b.setAttribute('aria-label', 'Error options');
  const x = s.ac > 0 ? st.cols.start(s.ac) - 20 : st.cols.start(s.ac + 1) + 2;
  setVars(b, { '--oxe-x': st.rowHeadW + x, '--oxe-y': st.rows.start(s.ar) });
  b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
  b.addEventListener('click', (e) => { e.stopPropagation(); const box = b.getBoundingClientRect(); st.onErrorButton?.({ x: box.left, y: box.bottom }); });
  layer.replaceChildren(b);
}

let gridCount = 0;

/** The outline of the cells copied or cut. */
function drawClip(st) {
  st.clipLayer.replaceChildren(...(st.clip ? [rangeBox(st, st.clip.r0, st.clip.c0, st.clip.r1, st.clip.c1, 'oxe-clip')] : []));
}

/** A grid of one sheet in `host`. Returns the handle the editor and the tests use. */
export function createGrid(host, { sheet, label = 'Sheet', looks = {}, readOnly = true }) {
  const view = div('oxe-grid', 'grid');
  view.tabIndex = 0;
  view.setAttribute('aria-readonly', String(readOnly));
  view.setAttribute('aria-label', label);
  const canvas = div('oxe-sheet');
  const colheads = div('oxe-colheads', 'row');
  colheads.setAttribute('aria-rowindex', '1');
  const corner = div('oxe-corner');
  corner.setAttribute('aria-hidden', 'true');
  const lines = div('oxe-lines');
  lines.setAttribute('aria-hidden', 'true');
  // (Fills over the gridlines and under the text, borders over the text.)
  const fills = div('oxe-fills');
  fills.setAttribute('aria-hidden', 'true');
  const body = div('oxe-rows');
  const edges = div('oxe-edges');
  edges.setAttribute('aria-hidden', 'true');
  // (The selection overlay sits over the borders and under the sticky headings.)
  const overlay = div('oxe-overlay');
  overlay.setAttribute('aria-hidden', 'true');
  // (Layers of their own for the selection and a formula's ranges: redrawing them leaves an in-cell editor alone. Other
  // people's cursors in a co-edited workbook go under this editor's selection.)
  const peerLayer = div('oxe-layer');
  const selLayer = div('oxe-layer');
  const markLayer = div('oxe-layer');
  const clipLayer = div('oxe-layer');
  // (Formulas > Trace Precedents and Trace Dependents: their arrows, arrows.js.)
  const traceLayer = arrowLayer();
  overlay.append(peerLayer, selLayer, markLayer, clipLayer, traceLayer);
  // (Frozen panes: the rows above the freeze stuck under the column headings, the columns left of it beside the row
  // headings, the corner where they meet; each its own set of layers, hidden from assistive technology: the sheet's
  // own cells are under them.)
  const pane = (cls) => {
    const wrap = div(`oxe-pane ${cls}`);
    wrap.setAttribute('aria-hidden', 'true');
    const inner = div('oxe-pane-inner');
    const layers = { lines: div('oxe-lines'), fills: div('oxe-fills'), body: div('oxe-rows'), edges: div('oxe-edges'), selLayer: div('oxe-layer'), filterLayer: div('oxe-filters'), noteLayer: div('oxe-notes') };
    const over = div('oxe-overlay');
    over.append(layers.selLayer);
    inner.append(layers.lines, layers.fills, layers.body, layers.edges, over, layers.noteLayer, layers.filterLayer);
    wrap.append(inner);
    return { wrap, inner, ...layers };
  };
  const panes = { top: pane('oxe-pane-top'), left: pane('oxe-pane-left'), corner: pane('oxe-pane-corner') };
  panes.top.inner.prepend(panes.corner.wrap);
  const pinHeads = div('oxe-colheads-pin');
  pinHeads.append(div('oxe-colheads-pin-inner'));
  // (The AutoFilter's buttons: over the cells, clickable, out of the tab order and hidden from assistive technology.)
  const filterLayer = div('oxe-filters');
  // (Excel's error button beside the active cell: errormenu.js.)
  const errLayer = div('oxe-filters');
  filterLayer.setAttribute('aria-hidden', 'true');
  // (Notes: their marks and boxes in a layer hidden from assistive technology, which hears each note as its cell's
  // description from the hidden texts.)
  const noteLayer = div('oxe-notes');
  noteLayer.setAttribute('aria-hidden', 'true');
  // (A data validation list's button and input message: the same.)
  // (Charts float over the cells; a screen reader is given each as a figure outside the grid instead.)
  const chartLayer = div('oxe-charts');
  chartLayer.setAttribute('aria-hidden', 'true');
  const dvLayer = div('oxe-dv');
  dvLayer.setAttribute('aria-hidden', 'true');
  const noteTexts = div('oxe-note-texts');
  noteTexts.hidden = true;
  canvas.append(colheads, panes.top.wrap, panes.left.wrap, lines, fills, body, edges, overlay, chartLayer, noteLayer, filterLayer, errLayer, dvLayer, noteTexts);
  view.appendChild(canvas);
  host.appendChild(view);
  const st = { view, canvas, colheads, corner, lines, fills, body, edges, overlay, peerLayer, selLayer, markLayer, clipLayer, arrowLayer: traceLayer, filterLayer, errLayer, noteLayer, noteTexts, dvLayer, chartLayer, panes, pinHeads, freeze: { rows: 0, cols: 0 }, zoom: 100, paneW: 0, paneH: 0, idBase: `oxe-g${++gridCount}`, sel: null, sheet: null, win: null, frame: 0, classOf: looks.classOf, deco: looks.deco, date1904: looks.date1904, color: looks.color };
  view.id = st.idBase; // (the sheet tabs' aria-controls; cell ids add to it)

  /** Show another sheet (`nextLooks`: its workbook's cell format classes and decorations, from cellstyles.js). */
  const measure = (next, nextLooks) => {
    const auto = autoRowHeights(next, nextLooks.deco, sheetGeometry(next).cols);
    Object.assign(st, { sheet: next, win: null, classOf: nextLooks.classOf, deco: nextLooks.deco, date1904: nextLooks.date1904, color: nextLooks.color, rowDesc: rowDescents(next, nextLooks.deco) }, sheetGeometry(next, { auto }));
    view.setAttribute('aria-rowcount', String(st.rows.count + 1));
    view.setAttribute('aria-colcount', String(st.cols.count + 1));
    setVars(view, { '--oxe-rowhead-w': st.rowHeadW, '--oxe-colhead-h': COLHEAD_H });
    setVars(canvas, { '--oxe-sheet-w': st.rowHeadW + st.cols.end, '--oxe-sheet-h': COLHEAD_H + st.rows.end });
    applyView(next.freeze, next.zoom);
  };
  /** Frozen rows and columns and the zoom (percent) applied: the panes sized, the sheet scaled. */
  const applyView = (freeze = { rows: 0, cols: 0 }, zoom = 100) => {
    st.freeze = { rows: freeze.rows || 0, cols: freeze.cols || 0 };
    st.zoom = zoom || 100;
    st.paneH = st.freeze.rows ? st.rows.start(st.freeze.rows) : 0;
    st.paneW = st.freeze.cols ? st.cols.start(st.freeze.cols) : 0;
    view.classList.toggle('oxe-frozen-rows', !!st.freeze.rows);
    view.classList.toggle('oxe-frozen-cols', !!st.freeze.cols);
    setVars(view, { '--oxe-pane-h': st.paneH, '--oxe-pane-w': st.paneW });
    canvas.style.zoom = st.zoom === 100 ? '' : String(st.zoom / 100);
  };
  /** A point of the view in the sheet's own coordinates: x past the row headings and y past the column headings
   *  (negative over them), at the zoom, frozen panes not scrolled. */
  const toContent = (lx, ly) => {
    const z = st.zoom / 100;
    let x = lx / z - st.rowHeadW;
    let y = ly / z - COLHEAD_H;
    if (x >= st.paneW) x += view.scrollLeft / z;
    if (y >= st.paneH) y += view.scrollTop / z;
    return { x, y };
  };
  const setSheet = (next, nextLabel = label, nextLooks = { classOf: st.classOf, deco: st.deco, date1904: st.date1904, color: st.color }) => {
    measure(next, nextLooks);
    view.setAttribute('aria-label', nextLabel);
    view.scrollTop = 0;
    view.scrollLeft = 0;
    draw(st, true);
  };
  const onScroll = () => {
    if (st.frame) return;
    st.frame = requestAnimationFrame(() => { st.frame = 0; draw(st); });
  };
  view.addEventListener('scroll', onScroll, { passive: true });
  const resize = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => draw(st)) : null;
  resize?.observe(view);
  setSheet(sheet, label, looks);

  // (A note's box opens while the pointer rests on its cell, as Excel's does; Shift+F2 keeps it open.)
  let rest = 0;
  const showNote = (key, pinned = false) => {
    if (st.noteOpen === key && st.notePinned === pinned) return;
    st.noteOpen = key;
    st.notePinned = pinned;
    drawNoteMarks(st, st.win);
  };
  view.addEventListener('mousemove', (e) => {
    const rect = view.getBoundingClientRect();
    const { x, y } = toContent(e.clientX - rect.left, e.clientY - rect.top);
    const key = x < 0 || y < 0 ? null : `${st.rows.indexAt(y)},${st.cols.indexAt(x)}`;
    if (key === st.noteHover) return;
    st.noteHover = key;
    clearTimeout(rest);
    if (key && st.notes?.has(key)) rest = setTimeout(() => showNote(key), 250);
    else if (st.noteOpen && !st.notePinned) showNote(null);
  });
  view.addEventListener('mouseleave', () => { clearTimeout(rest); st.noteHover = null; if (st.noteOpen && !st.notePinned) showNote(null); });

  return {
    el: view,
    setSheet,
    /** The rows and columns drawn now ({ r0, r1, c0, c1 }, 0-based, overscan included). */
    window: () => ({ ...st.win }),
    /** Scroll so that cell (r, c) is at the top left, and draw at once. */
    scrollToCell(r, c) {
      const z = st.zoom / 100;
      view.scrollTop = Math.max(0, st.rows.start(r) - st.paneH) * z;
      view.scrollLeft = Math.max(0, st.cols.start(c) - st.paneW) * z;
      draw(st);
    },
    /** Frozen rows and columns ({ rows, cols }) and the zoom (percent), as the sheet's view has them. */
    setView({ freeze, zoom } = {}) {
      applyView(freeze ?? st.freeze, zoom ?? st.zoom);
      draw(st, true);
      drawClip(st);
    },
    /** The view as shown: { freeze, zoom, topLeft } (topLeft: the first cell of the scrolled part). */
    view: () => {
      const z = st.zoom / 100;
      return { freeze: { ...st.freeze }, zoom: st.zoom, showFormulas: !!st.sheet?.showFormulas, topLeft: { r: st.rows.indexAt(view.scrollTop / z + st.paneH), c: st.cols.indexAt(view.scrollLeft / z + st.paneW) } };
    },
    /** Outline the cells copied or cut, dashed, until they are pasted or let go ({ r0, c0, r1, c1 }; null for none). */
    setClip(rect) {
      st.clip = rect;
      drawClip(st);
    },
    /** The tracer arrows (arrows.js drawArrows; [] for none). */
    setArrows(arrows) {
      st.arrows = arrows;
      drawArrows(st);
    },
    /** Outline a formula's ranges (selection.js drawMarks; [] for none). */
    setMarks(marks) {
      st.marks = marks;
      drawMarks(st);
    },
    /** Other people's cursors on this sheet (peers.js; [] for none), drawn. */
    setPeers(peers) {
      st.peers = peers;
      drawPeers(st);
    },
    /** Show a selection (selection.js; null for none) and draw it. */
    setSelection(sel) {
      st.sel = sel;
      draw(st, true);
    },
    /** What the sheet's conditional formats give its cells (ooxml/condformat.js conditionalLooks), drawn. */
    setConditional(looks) {
      st.cond = looks;
      draw(st, true);
    },
    /** The cells with an error indicator (ooxml/errorcheck.js errorFlags: "r,c" => [rules]), marked. */
    setErrors(flags) {
      st.errors = flags;
      draw(st, true);
    },
    /** The sheet's notes and comment threads (ooxml/notes.js readNotes), marked and described. */
    setNotes(notes) {
      st.notes = notes;
      st.noteOpen = null;
      draw(st, true);
    },
    /** Shift+F2: cell (r, c)'s note shown beside it, or hidden again; false when it has none. */
    toggleNote(r, c) {
      const key = `${r},${c}`;
      if (!st.notes?.has(key)) { if (st.noteOpen) showNote(null); return false; }
      showNote(st.noteOpen === key && st.notePinned ? null : key, true);
      return true;
    },
    /** What runs after the charts are drawn (chartpick.js: the picked one's outline put back). */
    onChartsDrawn(fn) { st.onChartsDrawn = fn; },
    /** The sheet's charts (ooxml/charts.js readCharts), drawn over the cells in the workbook's theme. */
    setCharts(charts, theme) {
      st.charts = charts;
      st.theme = theme;
      st.chartCache = new Map();
      draw(st, true);
    },
    /** The active cell's data validation ({ r, c, list, prompt: { title, text } }, or null): a list's button beside
     *  it, the input message under it. */
    setValidation(v) {
      st.validation = v;
      drawValidation(st);
    },
    /** What a click on a list's button does: `fn(button)`. */
    onDropdown(fn) { st.onDropdown = fn; },
    /** What a click on a filter button does: `fn(c, button)`. */
    onFilterButton(fn) { st.onFilter = fn; },
    /** Excel's error button clicked: fn({ x, y }) opens its menu there. */
    onErrorButton(fn) { st.onErrorButton = fn; },
    /** Where the error button is, if it is shown (Alt+Shift+F10 opens its menu there). */
    errorButtonAt() { const b = st.errLayer?.firstChild; if (!b) return null; const box = b.getBoundingClientRect(); return { x: box.left, y: box.bottom }; },
    /** Show the fill handle on the selection (not on a read-only workbook, or with filling turned off). */
    setFillHandle(on) {
      st.fillHandle = !!on;
      draw(st, true);
    },
    /** Whether a point of the page is on the fill handle (within 4 pixels of the selection's bottom-right corner). */
    fillHandleAt(clientX, clientY) {
      const corner = handleCorner(st);
      if (!corner) return false;
      const rect = view.getBoundingClientRect();
      const { x, y } = toContent(clientX - rect.left, clientY - rect.top);
      const near = 4 / (st.zoom / 100);
      return Math.abs(x - st.cols.start(corner.c + 1)) <= near && Math.abs(y - st.rows.start(corner.r + 1)) <= near;
    },
    /** While the fill handle is dragged: the cells it will fill outlined (null: none). */
    showFill(range) {
      st.filling = range;
      draw(st, true);
    },
    /** The cell under a point of the page: { r, c }, -1 for a heading (or null outside the grid). */
    hitTest(clientX, clientY) {
      const rect = view.getBoundingClientRect();
      const lx = clientX - rect.left;
      const ly = clientY - rect.top;
      if (lx < 0 || ly < 0 || lx > view.clientWidth || ly > view.clientHeight) return null;
      const { x, y } = toContent(lx, ly);
      return { r: y < 0 ? -1 : st.rows.indexAt(y), c: x < 0 ? -1 : st.cols.indexAt(x) };
    },
    /** A cell's box in the sheet's coordinates (the overlay's): { x, y, w, h }. */
    cellRect: (r, c) => ({ x: st.cols.start(c), y: st.rows.start(r), w: st.cols.size(c), h: st.rows.size(r) }),
    /** Scroll as little as needed for cell (r, c) to be in full view, past the headings. */
    scrollIntoView(r, c) {
      // (In the sheet's coordinates at the zoom; a cell in a frozen pane is always in view that way, the others are
      // brought out from under the panes.)
      const z = st.zoom / 100;
      const x = st.cols.start(c);
      const y = st.rows.start(r);
      const w = view.clientWidth / z - st.rowHeadW;
      const h = view.clientHeight / z - COLHEAD_H;
      const left = view.scrollLeft / z;
      const top = view.scrollTop / z;
      if (c >= st.freeze.cols) {
        if (x - st.paneW < left) view.scrollLeft = (x - st.paneW) * z;
        else if (x + st.cols.size(c) > left + w) view.scrollLeft = (x + st.cols.size(c) - w) * z;
      }
      if (r >= st.freeze.rows) {
        if (y - st.paneH < top) view.scrollTop = (y - st.paneH) * z;
        else if (y + st.rows.size(r) > top + h) view.scrollTop = (y + st.rows.size(r) - h) * z;
      }
      draw(st, true);
    },
    /** Rows in view, for paging. */
    rowsInView: () => { const z = st.zoom / 100; return Math.max(1, st.rows.indexAt(view.scrollTop / z + view.clientHeight / z - COLHEAD_H) - st.rows.indexAt(view.scrollTop / z + st.paneH)); },
    /** After cells changed: sizes worked out again (row heights follow their fonts), drawn where the view is. */
    refresh() {
      measure(st.sheet, { classOf: st.classOf, deco: st.deco, date1904: st.date1904, color: st.color });
      draw(st, true);
      drawClip(st);
    },
    /** The heading edge under a point, which a drag resizes: { axis: 'col' | 'row', index, size } (the line before the
     *  edge), or null. Within 4 pixels of the edge, as Excel's cursor changes. */
    edgeAt(clientX, clientY) {
      const rect = view.getBoundingClientRect();
      const lx = clientX - rect.left;
      const ly = clientY - rect.top;
      const near = 4 / (st.zoom / 100);
      const at = toContent(lx, ly);
      const find = (axis, pos) => {
        let i = axis.indexAt(pos);
        if (pos - axis.start(i) <= near && i > 0) i -= 1;
        else if (axis.start(i + 1) - pos > near) return null;
        return i;
      };
      if (ly >= 0 && at.y < 0 && at.x >= 0 && lx <= view.clientWidth) {
        const c = find(st.cols, at.x);
        return c == null ? null : { axis: 'col', index: c, size: st.cols.size(c) };
      }
      if (lx >= 0 && at.x < 0 && at.y >= 0 && ly <= view.clientHeight) {
        const r = find(st.rows, at.y);
        return r == null ? null : { axis: 'row', index: r, size: st.rows.size(r) };
      }
      return null;
    },
    /** While a heading edge is dragged: a line where the edge will be and Excel's tip ("Width: 20.00 (145 pixels)",
     *  "Height: 30.00 (40 pixels)"); no axis takes them away. */
    showGuide(axis, index, size) {
      st.canvas.querySelectorAll(':scope > .oxe-guide, :scope > .oxe-guide-tip').forEach((el) => el.remove());
      if (axis == null) return;
      const line = div('oxe-guide');
      const tip = div('oxe-guide-tip');
      const z = st.zoom / 100;
      const top = view.scrollTop / z;
      const left = view.scrollLeft / z;
      if (axis === 'col') {
        // (A frozen column's edge doesn't scroll.)
        const x = st.rowHeadW + st.cols.start(index) + size + (index < st.freeze.cols ? left : 0);
        setVars(line, { '--oxe-x': x, '--oxe-y': top, '--oxe-w': 1, '--oxe-h': view.clientHeight / z });
        setVars(tip, { '--oxe-x': x + 4, '--oxe-y': top + 2 });
        tip.textContent = `Width: ${(Math.max(0, size - 5) / MDW).toFixed(2)} (${Math.round(size)} pixels)`;
      } else {
        const y = COLHEAD_H + st.rows.start(index) + size + (index < st.freeze.rows ? top : 0);
        setVars(line, { '--oxe-x': left, '--oxe-y': y, '--oxe-w': view.clientWidth / z, '--oxe-h': 1 });
        setVars(tip, { '--oxe-x': left + st.rowHeadW + 4, '--oxe-y': y + 4 });
        tip.textContent = `Height: ${(Math.round(size) * 0.75).toFixed(2)} (${Math.round(size)} pixels)`;
      }
      st.canvas.append(line, tip);
    },
    /** The pixels AutoFit gives column c: its widest text as shown (a General number in full, the longest line of
     *  wrapped text), measured in each cell's font, plus 9 of padding (docs/SIZES.md); null when the column is empty. */
    fitWidth(c) { return columnFitWidth(st, c); },
    /** The sheet read again (rows or columns inserted or deleted, or undone): drawn where the view is. */
    setModel(next) {
      measure(next, { classOf: st.classOf, deco: st.deco, date1904: st.date1904, color: st.color });
      draw(st, true);
      drawClip(st);
    },
    /** The workbook's cell formats changed (new ones added): drawn again with them, where the view is. */
    setLooks(nextLooks) {
      measure(st.sheet, nextLooks);
      draw(st, true);
      drawClip(st);
    },
    /** A cell's text as the grid shows it (a General number to the width of its column). */
    shownText(r, c) {
      const cell = st.sheet.cells.get(r)?.get(c);
      if (!cell) return null;
      const look = st.deco ? st.deco(cell.s || 0) : null;
      return cellDisplay(cell, colChars(st.cols.size(c), MDW), look ? look.code : 'General', st.date1904).text;
    },
    setReadOnly: (flag) => view.setAttribute('aria-readonly', String(flag)),
    /** The overlay element, which an in-cell editor joins. */
    overlay,
    /** Where a column and a row start and how large they are, in pixels (for tests and, later, layout checks). */
    geometry: () => ({ rows: st.rows, cols: st.cols, rowHeadW: st.rowHeadW, colHeadH: COLHEAD_H }),
    destroy() {
      resize?.disconnect();
      if (st.frame) cancelAnimationFrame(st.frame);
      view.remove();
    },
  };
}

