// The sheet grid, virtualised: only the rows and columns in view (and a few around them) are in the DOM, redrawn from
// the sparse cell store as the view scrolls. The scroll position is the one source of truth for what is drawn.
//
// Layout: .oxe-grid scrolls; inside it .oxe-sheet is as large as the whole grid. The column headings are one row stuck
// to the top (position: sticky), each row's heading is stuck to the left inside its row, and cells sit at their sheet
// coordinates. Geometry goes through custom properties (--oxe-x, --oxe-y, --oxe-w, --oxe-h), never through styles
// written for state.
import { Axis } from './axis.js';
import { decorations } from './decor.js';
import { autoRowHeights, rowDescents } from './rowheights.js';
import { fontDescentPx } from './fonts.js';
import { drawSelection, drawMarks, rangeBox, handleCorner } from './selection.js';
import { drawFilterButtons } from './filterbuttons.js';
import { colName, MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { cellDisplay } from '../ooxml/numfmt.js';
import { colWidthToPx, defaultColPx, rowHeightToPx, printColPx, printRowPx, colChars } from '../ooxml/units.js';

/** Calibri 11's maximum digit width at 96 dpi; it will come from the Normal style's font once styles are read. */
const MDW = 7;
/** The column headings' height and the room each row heading's digits need (view chrome, not workbook content). */
const COLHEAD_H = 20;
/** How far past the used range the grid reaches (Excel's goes on to row 1,048,576; this one grows with the content). */
const PAD_ROWS = 100;
const PAD_COLS = 10;
/** Rows and columns drawn beyond the edges of the view, so a small scroll shows no gap before the redraw. */
const OVERSCAN = 3;
/** How many columns back to look for text that runs on into the view from the left. */
const SPILL_BACK = 64;

const px = (n) => `${n}px`;
export const setVars = (el, vars) => { for (const k in vars) el.style.setProperty(k, px(vars[k])); };
export const div = (cls, role) => {
  const el = document.createElement('div');
  el.className = cls;
  if (role) el.setAttribute('role', role);
  return el;
};

/** The grid's rows and columns for a sheet: sizes from its rows, cols and defaults, hidden ones of no size. With
 *  `print`, columns and rows at their printed sizes (units.js printColPx, printRowPx, which needs the Normal style's
 *  typeface) instead of their sizes on screen. With `auto` (rowheights.js), rows without customHeight take Excel's
 *  own heights instead of their stored `ht`. */
export function sheetGeometry(sheet, { print = false, normalFamily, auto = null } = {}) {
  const colPx = print ? printColPx : (px) => px;
  const rowPx = print ? (pt) => printRowPx(pt, normalFamily) : rowHeightToPx;
  const rowCount = Math.min(MAX_ROWS, Math.max(sheet.maxRow + 1 + PAD_ROWS, 200));
  const colCount = Math.min(MAX_COLS, Math.max(sheet.maxCol + 1 + PAD_COLS, 26));
  const rowSizes = new Map();
  for (const [r, info] of sheet.rows) {
    if (r >= rowCount) continue;
    if (info.hidden) rowSizes.set(r, 0);
    else if (info.ht != null && (!auto || info.customHeight)) rowSizes.set(r, rowPx(info.ht));
  }
  if (auto) for (const [r, pt] of auto) if (r < rowCount && !rowSizes.has(r)) rowSizes.set(r, rowPx(pt));
  const colSizes = new Map();
  for (const col of sheet.cols) {
    for (let c = col.min; c <= Math.min(col.max, colCount - 1); c++) {
      if (col.hidden) colSizes.set(c, 0);
      else if (col.width != null) colSizes.set(c, colPx(colWidthToPx(col.width, MDW)));
    }
  }
  const rows = new Axis(rowCount, rowPx(sheet.format.defaultRowHeight), rowSizes);
  const cols = new Axis(colCount, colPx(defaultColPx(sheet.format, MDW)), colSizes);
  return { rows, cols, rowHeadW: Math.max(3, String(rowCount).length) * MDW + 12 };
}

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

function drawColumnHeadings(st, win) {
  const heading = (c, role) => {
    const el = div('oxe-colhead', role);
    if (role) el.setAttribute('aria-colindex', String(c + 2));
    el.textContent = colName(c);
    setVars(el, { '--oxe-x': st.cols.start(c), '--oxe-w': st.cols.size(c) });
    return el;
  };
  const kids = [st.corner];
  // (Frozen columns' headings stay where they are, in a pinned strip of their own.)
  const fc = st.freeze?.cols || 0;
  if (fc && st.pinHeads) {
    const inner = st.pinHeads.firstChild;
    inner.replaceChildren(...Array.from({ length: fc }, (_, c) => c).filter((c) => st.cols.size(c)).map((c) => heading(c, null)));
    kids.push(st.pinHeads);
  }
  for (let c = win.c0; c <= win.c1; c++) if (st.cols.size(c)) kids.push(heading(c, 'columnheader'));
  st.colheads.replaceChildren(...kids);
}

/** The frozen panes (rows above and columns left of the freeze, and the corner where they meet), each drawn with the
 *  layers the sheet has (gridlines, fills, cells, borders, selection) over the part of the sheet they keep in place. */
function drawPanes(st, win) {
  const fr = st.freeze?.rows || 0;
  const fc = st.freeze?.cols || 0;
  const paint = (pane, w) => {
    if (!w) { for (const k of ['lines', 'fills', 'body', 'edges', 'selLayer', 'filterLayer']) pane[k].replaceChildren(); return; }
    // (The pane's own layers; no cell ids, which belong to the sheet's own cells.)
    const sub = Object.assign(Object.create(st), { lines: pane.lines, fills: pane.fills, body: pane.body, edges: pane.edges, selLayer: pane.selLayer, filterLayer: pane.filterLayer, idBase: null });
    drawGridlines(sub, w);
    const { fills, edges } = decorations(sub, w);
    pane.fills.replaceChildren(...fills);
    drawRows(sub, w);
    pane.edges.replaceChildren(...edges);
    drawSelection(sub);
    drawFilterButtons(sub, w);
  };
  paint(st.panes.top, fr ? { r0: 0, r1: fr - 1, c0: win.c0, c1: win.c1 } : null);
  paint(st.panes.left, fc ? { r0: win.r0, r1: win.r1, c0: 0, c1: fc - 1 } : null);
  paint(st.panes.corner, fr && fc ? { r0: 0, r1: fr - 1, c0: 0, c1: fc - 1 } : null);
}

function drawGridlines(st, win) {
  if (!st.sheet.showGridLines) return st.lines.replaceChildren();
  const top = st.rows.start(win.r0);
  const left = st.cols.start(win.c0);
  const height = st.rows.start(win.r1 + 1) - top;
  const width = st.cols.start(win.c1 + 1) - left;
  const kids = [];
  // (A cell's gridlines are its right and bottom pixel, inside its own width and height.)
  for (let c = win.c0; c <= win.c1; c++) {
    if (!st.cols.size(c)) continue;
    const el = div('oxe-vline');
    setVars(el, { '--oxe-x': st.cols.start(c + 1) - 1, '--oxe-y': top, '--oxe-h': height });
    kids.push(el);
  }
  for (let r = win.r0; r <= win.r1; r++) {
    if (!st.rows.size(r)) continue;
    const el = div('oxe-hline');
    setVars(el, { '--oxe-x': left, '--oxe-y': st.rows.start(r + 1) - 1, '--oxe-w': width });
    kids.push(el);
  }
  st.lines.replaceChildren(...kids);
}

/** The merged range a cell is inside, if any. */
export const mergeAt = (st, r, c) => st.sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) || null;

const NO_ALIGN = { horizontal: 'general', vertical: 'bottom', wrap: false, indent: 0 };
/** An indent level in px: 10.08 in print (Excel's PDF of test/fixtures/align.xlsx: one, two and three levels from
 *  either side), 9 on screen (not yet measured). */
const INDENT_PX = { screen: 9, print: 10.08 };

/** Where a cell's text sits across: general alignment puts numbers right, booleans and errors in the middle and text
 *  left; fill, justify and distributed are drawn left for now, centre-across-selection as centre. */
function horizontalOf(align, cell) {
  const h = align.horizontal;
  if (h === 'general') return cell.t === 'n' ? 'right' : cell.t === 'b' || cell.t === 'e' ? 'center' : 'left';
  if (h === 'center' || h === 'centerContinuous') return 'center';
  return h === 'right' ? 'right' : 'left';
}

/** The columns either side of c that text may run on into: up to the nearest cell with something in it, a merged
 *  range or the edge of what is drawn. Returns [first, last + 1). */
function runOn(st, r, c, line, win) {
  let end = c + 1;
  while (end <= win.c1 + 1 && end < st.cols.count && !(line.get(end)?.v != null) && !mergeAt(st, r, end)) end++;
  let begin = c;
  while (begin - 1 >= Math.max(0, win.c0 - 1) && !(line.get(begin - 1)?.v != null) && !mergeAt(st, r, begin - 1)) begin--;
  return [begin, end];
}

/** One cell's element: its text in an inner span (which carries the line height and the baseline shifts), placed by
 *  the cell format's alignment and indent; text runs on over empty neighbours (left text to the right, right text to
 *  the left, centred text both ways, as far as the nearer side allows), wrapped text and numbers never do. */
export function cellElement(st, r, c, cell, line, win, rowTop) {
  const merge = mergeAt(st, r, c);
  const look = st.deco ? st.deco(cell.s || 0) : null;
  const align = (look && look.align) || NO_ALIGN;
  const isText = cell.t === 's' || cell.t === 'str';
  const wrap = align.wrap && isText;
  // (A General number gets the characters that fit its cell on screen, in print too: st.screenCols when printing.)
  const screen = st.screenCols || st.cols;
  const screenW = merge ? screen.start(merge.c1 + 1) - screen.start(c) : screen.size(c);
  const shown = cellDisplay(cell, colChars(screenW, MDW), look ? look.code : 'General', st.date1904);
  // (A cell full of #s is centred, whatever its alignment: Excel's PDF of numfmt.xlsx, negative dates.)
  const h = shown.hashes ? 'center' : shown.textFormat && align.horizontal === 'general' ? 'left' : horizontalOf(align, cell);
  // (Without wrapping, a line break shows as nothing; wrapped, it breaks the line.)
  const clean = (t) => (wrap ? t.replace(/\r/g, '') : t.replace(/[\r\n]/g, ''));
  const el = div(`oxe-cell oxe-h-${h} oxe-v-${align.vertical === 'top' ? 'top' : align.vertical === 'center' ? 'center' : 'bottom'}${wrap ? ' oxe-wrap' : ''}`, 'gridcell');
  if (st.classOf) el.classList.add(st.classOf(cell.s || 0));
  el.setAttribute('aria-colindex', String(c + 2));
  if (st.idBase) el.id = `${st.idBase}-${r}-${c}`;
  if (st.sel && r >= st.sel.r0 && r <= st.sel.r1 && c >= st.sel.c0 && c <= st.sel.c1) el.setAttribute('aria-selected', 'true');
  let x = st.cols.start(c);
  let w = st.cols.size(c);
  let hgt = st.rows.size(r);
  if (merge) {
    el.classList.add('oxe-merged');
    el.setAttribute('aria-colspan', String(merge.c1 - merge.c0 + 1));
    el.setAttribute('aria-rowspan', String(merge.r1 - merge.r0 + 1));
    w = st.cols.start(merge.c1 + 1) - x;
    hgt = st.rows.start(merge.r1 + 1) - st.rows.start(r);
  } else if (isText && !wrap) {
    const [begin, end] = runOn(st, r, c, line, win);
    const left = st.cols.start(begin);
    const right = st.cols.start(end);
    if (h === 'left') w = right - x;
    else if (h === 'right') { w = x + w - left; x = left; }
    else {
      const centre = x + w / 2;
      const half = Math.min(centre - left, right - centre);
      x = centre - half;
      w = 2 * half;
    }
  }
  const vars = { '--oxe-x': x, '--oxe-y': st.rows.start(r) - rowTop, '--oxe-w': w, '--oxe-h': hgt };
  // Bottom-aligned text sits on its row's baseline, which the row's tallest font sets (rowheights.js rowDescents).
  if (!wrap && align.vertical !== 'top' && align.vertical !== 'center' && look && st.rowDesc) {
    const lift = (st.rowDesc.get(r) ?? 0) - fontDescentPx(look.font.family, look.font.sizePx);
    if (lift > 0.01) vars['--oxe-row-shift'] = -lift;
  }
  if (align.indent && h !== 'center') {
    el.classList.add(h === 'right' ? 'oxe-indent-r' : 'oxe-indent-l');
    vars['--oxe-indent'] = align.indent * (st.print ? INDENT_PX.print : INDENT_PX.screen);
  }
  el.appendChild(textSpan(shown, clean));
  const color = shown.color && st.color ? st.color(shown.color) : null;
  if (color) {
    el.classList.add('oxe-nf-color');
    el.style.setProperty('--oxe-nf-color', color);
  }
  // (A formatted number wider than its cell shows #s instead; fitNumbers checks once the cell is in the page.)
  if (shown.fit) el.dataset.oxeFit = '1';
  setVars(el, vars);
  return el;
}

/** The text span: plain text, or for formats with padding (_x: a space as wide as x, drawn as a hidden x) and fill
 *  (*x: x repeated across the room left in the cell) the pieces in a row. */
function textSpan(shown, clean) {
  const span = document.createElement('span');
  span.className = 'oxe-tx';
  if (shown.parts.every((p) => typeof p === 'string')) {
    span.textContent = clean(shown.text);
    return span;
  }
  span.classList.add('oxe-parts');
  if (shown.parts.some((p) => p && p.fill)) span.classList.add('oxe-parts-fill');
  for (const p of shown.parts) {
    if (typeof p === 'string') span.append(clean(p));
    else {
      const piece = document.createElement('span');
      piece.className = p.pad ? 'oxe-pad' : 'oxe-repeat';
      piece.textContent = p.pad || p.fill.repeat(200);
      span.appendChild(piece);
    }
  }
  return span;
}

let measure = null;
/** The canvas AutoFit measures text with. */
let fitCtx = null;
/** Replace formatted numbers wider than their cells with #s, as Excel does (as many as fit). Run once the cells are in
 *  the page, so one layout serves them all; the width is measured in the cell's own font. */
export function fitNumbers(container) {
  const cells = container.querySelectorAll('.oxe-cell[data-oxe-fit]');
  if (!cells.length) return;
  measure = measure || document.createElement('canvas').getContext('2d');
  const jobs = [];
  for (const el of cells) {
    const span = el.firstElementChild;
    const cs = getComputedStyle(span);
    const box = getComputedStyle(el);
    const room = parseFloat(el.style.getPropertyValue('--oxe-w')) - parseFloat(box.paddingLeft) - parseFloat(box.paddingRight);
    measure.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    // (Padding pieces count; a fill shrinks to nothing.)
    const text = [...span.childNodes].map((n) => (n.classList && n.classList.contains('oxe-repeat') ? '' : n.textContent)).join('');
    if (measure.measureText(text).width > room + 0.5) jobs.push([span, Math.max(1, Math.floor(room / measure.measureText('#').width))]);
  }
  for (const [span, n] of jobs) {
    span.classList.remove('oxe-parts', 'oxe-parts-fill');
    span.textContent = '#'.repeat(n);
    span.parentElement.classList.remove('oxe-h-left', 'oxe-h-right');
    span.parentElement.classList.add('oxe-h-center');
  }
}

function drawRows(st, win) {
  const kids = [];
  const sheet = st.sheet;
  // Merged ranges that start above the view are drawn with the first row in view.
  const pending = sheet.merges.filter((m) => m.r0 < win.r0 && m.r1 >= win.r0 && m.c1 >= win.c0 && m.c0 <= win.c1);
  for (let r = win.r0; r <= win.r1; r++) {
    const h = st.rows.size(r);
    if (!h) continue;
    const row = div('oxe-row', 'row');
    row.setAttribute('aria-rowindex', String(r + 2));
    const top = st.rows.start(r);
    setVars(row, { '--oxe-y': top, '--oxe-h': h });
    const head = div('oxe-rowhead', 'rowheader');
    head.textContent = String(r + 1);
    row.appendChild(head);
    const line = sheet.cells.get(r);
    const cells = [];
    for (const m of pending.splice(0)) cells.push([m.r0, m.c0]);
    if (line) {
      // Text in a cell left of the view may run on into it.
      for (let c = win.c0 - 1; c >= Math.max(0, win.c0 - SPILL_BACK); c--) {
        const cell = line.get(c);
        if (cell?.v == null) continue;
        if (cell.t === 's' || cell.t === 'str') cells.push([r, c]);
        break;
      }
      for (let c = win.c0; c <= win.c1; c++) if (line.get(c) && st.cols.size(c)) cells.push([r, c]);
    }
    // (A merged range is drawn even when its top-left cell is empty: it hides the gridlines inside it.)
    for (const m of sheet.merges) if (m.r0 === r && m.c1 >= win.c0 && m.c0 <= win.c1) cells.push([m.r0, m.c0]);
    const seen = new Set();
    for (const [cr, cc] of cells) {
      if (seen.has(`${cr},${cc}`)) continue;
      seen.add(`${cr},${cc}`);
      const merge = mergeAt(st, cr, cc);
      // (Inside a merged range only the top-left cell shows; the others are kept in the file but not drawn.)
      if (merge && (merge.r0 !== cr || merge.c0 !== cc)) continue;
      const cell = sheet.cells.get(cr)?.get(cc) || { t: 'n', v: null };
      if (cell.v == null && !merge && !(st.sel && st.sel.ar === cr && st.sel.ac === cc)) continue;
      row.appendChild(cellElement(st, cr, cc, cell, sheet.cells.get(cr) || new Map(), win, top));
    }
    // (The active cell is drawn even when it is empty, so that the grid can name it as its active descendant.)
    const sel = st.sel;
    if (sel && sel.ar === r && sel.ac >= win.c0 && sel.ac <= win.c1 && !seen.has(`${r},${sel.ac}`) && !mergeAt(st, r, sel.ac)) {
      row.appendChild(cellElement(st, r, sel.ac, { t: 'n', v: null }, line || new Map(), win, top));
    }
    kids.push(row);
  }
  st.body.replaceChildren(...kids);
  fitNumbers(st.body);
}

/** Redraw if the window in view changed (or `force`). */
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
  drawSelection(st);
  drawMarks(st);
  drawFilterButtons(st, win);
  if (st.panes) drawPanes(st, win);
  const active = st.sel && st.idBase ? document.getElementById(`${st.idBase}-${st.sel.ar}-${st.sel.ac}`) : null;
  if (active) st.view.setAttribute('aria-activedescendant', active.id);
  else st.view.removeAttribute('aria-activedescendant');
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
  // (Layers of their own for the selection and a formula's ranges: redrawing them leaves an in-cell editor alone.)
  const selLayer = div('oxe-layer');
  const markLayer = div('oxe-layer');
  const clipLayer = div('oxe-layer');
  overlay.append(selLayer, markLayer, clipLayer);
  // (Frozen panes: the rows above the freeze stuck under the column headings, the columns left of it beside the row
  // headings, the corner where they meet; each its own set of layers, hidden from assistive technology: the sheet's
  // own cells are under them.)
  const pane = (cls) => {
    const wrap = div(`oxe-pane ${cls}`);
    wrap.setAttribute('aria-hidden', 'true');
    const inner = div('oxe-pane-inner');
    const layers = { lines: div('oxe-lines'), fills: div('oxe-fills'), body: div('oxe-rows'), edges: div('oxe-edges'), selLayer: div('oxe-layer'), filterLayer: div('oxe-filters') };
    const over = div('oxe-overlay');
    over.append(layers.selLayer);
    inner.append(layers.lines, layers.fills, layers.body, layers.edges, over, layers.filterLayer);
    wrap.append(inner);
    return { wrap, inner, ...layers };
  };
  const panes = { top: pane('oxe-pane-top'), left: pane('oxe-pane-left'), corner: pane('oxe-pane-corner') };
  panes.top.inner.prepend(panes.corner.wrap);
  const pinHeads = div('oxe-colheads-pin');
  pinHeads.append(div('oxe-colheads-pin-inner'));
  // (The AutoFilter's buttons: over the cells, clickable, out of the tab order and hidden from assistive technology.)
  const filterLayer = div('oxe-filters');
  filterLayer.setAttribute('aria-hidden', 'true');
  canvas.append(colheads, panes.top.wrap, panes.left.wrap, lines, fills, body, edges, overlay, filterLayer);
  view.appendChild(canvas);
  host.appendChild(view);
  const st = { view, canvas, colheads, corner, lines, fills, body, edges, overlay, selLayer, markLayer, clipLayer, filterLayer, panes, pinHeads, freeze: { rows: 0, cols: 0 }, zoom: 100, paneW: 0, paneH: 0, idBase: `oxe-g${++gridCount}`, sel: null, sheet: null, win: null, frame: 0, classOf: looks.classOf, deco: looks.deco, date1904: looks.date1904, color: looks.color };
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
      return { freeze: { ...st.freeze }, zoom: st.zoom, topLeft: { r: st.rows.indexAt(view.scrollTop / z + st.paneH), c: st.cols.indexAt(view.scrollLeft / z + st.paneW) } };
    },
    /** Outline the cells copied or cut, dashed, until they are pasted or let go ({ r0, c0, r1, c1 }; null for none). */
    setClip(rect) {
      st.clip = rect;
      drawClip(st);
    },
    /** Outline a formula's ranges (selection.js drawMarks; [] for none). */
    setMarks(marks) {
      st.marks = marks;
      drawMarks(st);
    },
    /** Show a selection (selection.js; null for none) and draw it. */
    setSelection(sel) {
      st.sel = sel;
      draw(st, true);
    },
    /** What a click on a filter button does: `fn(c, button)`. */
    onFilterButton(fn) { st.onFilter = fn; },
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
    fitWidth(c) {
      if (!fitCtx) fitCtx = document.createElement('canvas').getContext('2d');
      let widest = 0;
      for (const [r, line] of st.sheet.cells) {
        const cell = line.get(c);
        if (!cell || cell.v == null || mergeAt(st, r, c)) continue;
        const look = st.deco ? st.deco(cell.s || 0) : null;
        const text = cellDisplay(cell, Infinity, look ? look.code : 'General', st.date1904).text;
        if (!text) continue;
        fitCtx.font = look ? look.font.css : '14.6667px sans-serif';
        for (const part of String(text).split('\n')) widest = Math.max(widest, fitCtx.measureText(part.replace(/\r/g, '')).width);
      }
      return widest ? Math.ceil(widest) + 9 : null;
    },
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
