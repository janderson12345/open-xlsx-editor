// The sheet grid, virtualised: only the rows and columns in view (and a few around them) are in the DOM, redrawn from
// the sparse cell store as the view scrolls. The scroll position is the one source of truth for what is drawn.
//
// Layout: .oxe-grid scrolls; inside it .oxe-sheet is as large as the whole grid. The column headings are one row stuck
// to the top (position: sticky), each row's heading is stuck to the left inside its row, and cells sit at their sheet
// coordinates. Geometry goes through custom properties (--oxe-x, --oxe-y, --oxe-w, --oxe-h), never through styles
// written for state.
import { Axis } from './axis.js';
import { decorations } from './decor.js';
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
 *  typeface) instead of their sizes on screen. */
export function sheetGeometry(sheet, { print = false, normalFamily } = {}) {
  const colPx = print ? printColPx : (px) => px;
  const rowPx = print ? (pt) => printRowPx(pt, normalFamily) : rowHeightToPx;
  const rowCount = Math.min(MAX_ROWS, Math.max(sheet.maxRow + 1 + PAD_ROWS, 200));
  const colCount = Math.min(MAX_COLS, Math.max(sheet.maxCol + 1 + PAD_COLS, 26));
  const rowSizes = new Map();
  for (const [r, info] of sheet.rows) {
    if (r >= rowCount) continue;
    if (info.hidden) rowSizes.set(r, 0);
    else if (info.ht != null) rowSizes.set(r, rowPx(info.ht));
  }
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
  const w = Math.max(0, view.clientWidth - rowHeadW);
  const h = Math.max(0, view.clientHeight - COLHEAD_H);
  return {
    r0: Math.max(0, rows.indexAt(view.scrollTop) - OVERSCAN),
    r1: Math.min(rows.count - 1, rows.indexAt(view.scrollTop + h) + OVERSCAN),
    c0: Math.max(0, cols.indexAt(view.scrollLeft) - OVERSCAN),
    c1: Math.min(cols.count - 1, cols.indexAt(view.scrollLeft + w) + OVERSCAN),
  };
}

function drawColumnHeadings(st, win) {
  const kids = [st.corner];
  for (let c = win.c0; c <= win.c1; c++) {
    const w = st.cols.size(c);
    if (!w) continue;
    const el = div('oxe-colhead', 'columnheader');
    el.setAttribute('aria-colindex', String(c + 2));
    el.textContent = colName(c);
    setVars(el, { '--oxe-x': st.cols.start(c), '--oxe-w': w });
    kids.push(el);
  }
  st.colheads.replaceChildren(...kids);
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
      if (cell.v == null && !merge) continue;
      row.appendChild(cellElement(st, cr, cc, cell, sheet.cells.get(cr) || new Map(), win, top));
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
}

/** A read-only grid of one sheet in `host`. Returns the handle the editor and the tests use. */
export function createGrid(host, { sheet, label = 'Sheet', looks = {} }) {
  const view = div('oxe-grid', 'grid');
  view.tabIndex = 0;
  view.setAttribute('aria-readonly', 'true');
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
  canvas.append(colheads, lines, fills, body, edges);
  view.appendChild(canvas);
  host.appendChild(view);
  const st = { view, canvas, colheads, corner, lines, fills, body, edges, sheet: null, win: null, frame: 0, classOf: looks.classOf, deco: looks.deco, date1904: looks.date1904, color: looks.color };

  /** Show another sheet (`nextLooks`: its workbook's cell format classes and decorations, from cellstyles.js). */
  const setSheet = (next, nextLabel = label, nextLooks = { classOf: st.classOf, deco: st.deco, date1904: st.date1904, color: st.color }) => {
    Object.assign(st, { sheet: next, win: null, classOf: nextLooks.classOf, deco: nextLooks.deco, date1904: nextLooks.date1904, color: nextLooks.color }, sheetGeometry(next));
    view.setAttribute('aria-label', nextLabel);
    view.setAttribute('aria-rowcount', String(st.rows.count + 1));
    view.setAttribute('aria-colcount', String(st.cols.count + 1));
    setVars(view, { '--oxe-rowhead-w': st.rowHeadW, '--oxe-colhead-h': COLHEAD_H });
    setVars(canvas, { '--oxe-sheet-w': st.rowHeadW + st.cols.end, '--oxe-sheet-h': COLHEAD_H + st.rows.end });
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
      view.scrollTop = st.rows.start(r);
      view.scrollLeft = st.cols.start(c);
      draw(st);
    },
    /** Where a column and a row start and how large they are, in pixels (for tests and, later, layout checks). */
    geometry: () => ({ rows: st.rows, cols: st.cols, rowHeadW: st.rowHeadW, colHeadH: COLHEAD_H }),
    destroy() {
      resize?.disconnect();
      if (st.frame) cancelAnimationFrame(st.frame);
      view.remove();
    },
  };
}
