// The printed sheet, as Excel prints it (measured against its PDFs, docs/LAYOUT.md and docs/PRINT.md): its print area
// (or used range) split into pages (columns that fit across the printable width, rows that fit down its height, at its
// manual breaks, pages in the sheet's page order), the print title rows and columns repeated on every page, all at
// the sheet's scale or the scale that fits it on the pages asked for, each page at its paper size with the sheet's
// margins. Built only for printing (and for PDF export and tools/excel-compare); hidden on screen.
import { sheetGeometry } from './geometry.js';
import { cellElement, mergeAt, fitNumbers, lineSegments } from './cells.js';
import { setVars, div } from './dom.js';
import { decorations } from './decor.js';
import { autoRowHeights, rowDescents } from './rowheights.js';
import { setSheet } from './sheets.js';
import { drawHeaderFooter } from './printhf.js';
import { colName } from '../ooxml/refs.js';

const DPI = 96;
/** The room Excel paginates into: the printable height, or width, less a point (measured: Calibri rows of 121 dots at
 *  600 dpi, 47 to a 9.5in page; at 60% rows of 73 dots, 77 to it, where 78 would fill 5,694 of its 5,700 dots; at 90%
 *  columns of 387 dots, 10 across 7.1in where 11 would fill 4,257 of 4,260). */
const ROOM_LESS = 96 / 72;

/** Split an axis's indexes `from`..`last` into runs that each fit in `room` pixels (a run always takes at least one),
 *  starting a new run at each of `breaks` (indexes a page starts at). `roomAt(start)` gives the room of a run that
 *  starts at `start`, when it differs (titles repeated on it). */
export function breakAxis(axis, last, room, { from = 0, breaks = [], roomAt = null } = {}) {
  const runs = [];
  const forced = new Set(breaks);
  let start = from;
  let used = 0;
  let fits = roomAt ? roomAt(start) : room;
  for (let i = from; i <= last; i++) {
    const size = axis.size(i);
    if (i > start && (forced.has(i) || used + size > fits)) {
      runs.push([start, i - 1]);
      start = i;
      used = 0;
      fits = roomAt ? roomAt(start) : room;
    }
    used += size;
  }
  runs.push([start, last]);
  return runs;
}

/** The size of lines a..b on an axis. */
const span = (axis, a, b) => (b < a ? 0 : axis.start(b) + axis.size(b) - axis.start(a));

/**
 * The pages of a sheet: { pages: [{ r0, r1, c0, c1 }] in printing order, paper and printable area in pixels, the
 * scale, and the titles }. `geometry` is the sheet's printed geometry at that scale (sizes before the scale: renderPrint
 * draws the page's cells zoomed by it); `print`: { area, titles } (printarea.js printSetup), `scale` (a fraction).
 */
export function paginate(sheet, geometry = sheetGeometry(sheet, { print: true }), { area: printArea = null, titles = { rows: null, cols: null }, scale = 1 } = {}) {
  const { page } = sheet;
  const [pw, ph] = page.landscape ? [page.paper[1], page.paper[0]] : page.paper;
  const m = page.margins;
  const area = { x: m.left * DPI, y: m.top * DPI, w: (pw - m.left - m.right) * DPI, h: (ph - m.top - m.bottom) * DPI };
  const range = printArea || { r0: 0, c0: 0, r1: Math.max(0, sheet.maxRow), c1: Math.max(0, sheet.maxCol) };
  const { rows, cols } = geometry;
  // (Title rows take room on every page that starts after them; a page that starts at or before them shows them in
  // their place.)
  const titleRows = titles.rows ? span(rows, titles.rows[0], titles.rows[1]) : 0;
  const titleCols = titles.cols ? span(cols, titles.cols[0], titles.cols[1]) : 0;
  const breaks = page.breaks || { rows: [], cols: [] };
  // (Headings: Print takes the row numbers' column and the column letters' row from every page's room.)
  const heads = page.headings ? { w: headingWidth(range.r1), h: rows.base } : null;
  const roomDown = (area.h - ROOM_LESS) / scale - (heads ? heads.h : 0);
  const roomAcross = (area.w - ROOM_LESS) / scale - (heads ? heads.w : 0);
  const rowRuns = breakAxis(rows, range.r1, roomDown, { from: range.r0, breaks: breaks.rows, roomAt: (start) => roomDown - (titles.rows && start > titles.rows[1] ? titleRows : 0) });
  const colRuns = breakAxis(cols, range.c1, roomAcross, { from: range.c0, breaks: breaks.cols, roomAt: (start) => roomAcross - (titles.cols && start > titles.cols[1] ? titleCols : 0) });
  const pages = [];
  const add = ([r0, r1], [c0, c1]) => pages.push({ r0, r1, c0, c1 });
  if (page.pageOrder === 'overThenDown') for (const rr of rowRuns) for (const cr of colRuns) add(rr, cr);
  else for (const cr of colRuns) for (const rr of rowRuns) add(rr, cr);
  return { pages, paper: { w: pw * DPI, h: ph * DPI }, area, scale, titles, center: { h: !!page.centerH, v: !!page.centerV }, gridLines: !!page.gridLines, heads };
}

/** Where Excel's printed headings sit against the cells (its PDFs): every rule 1.28px (a hairline drawn over it), the
 *  row numbers' ones and the one under the letters 1.31px in from their start, the letters' separators from 1.15px
 *  down, the vertical one from just under the rule below the letters to the foot of the last; every label 1.4px
 *  right of and 1.1px below the middle of its box. */
const HEAVY_IN = 1.31;
const LABEL_SHIFT = { x: 1.4, y: 1.1 };

/** The printed row numbers' column, as wide as Excel's (its PDFs of workbooks made in code: rows to 999 27.52px, rows
 *  past 999 35.04): 4.96px and a printed digit's 7.52 for each digit of the last row's number, three at least. */
export const headingWidth = (lastRow) => 4.96 + 7.52 * Math.max(3, String(lastRow + 1).length);

/**
 * The scale a sheet prints at: its own (`scale`%), or with fit to page the largest whole percentage (10 to 100) at
 * which its columns fit `fitToWidth` pages across and its rows `fitToHeight` down (0: as many as it takes), as Excel
 * chooses it (print-fit.xlsx: 24 columns on one page, 41%). `geometry` at 100%.
 */
export function printScale(sheet, geometry, range) {
  const { page } = sheet;
  if (!page.fitToPage) return Math.min(400, Math.max(10, page.scale || 100)) / 100;
  const [pw, ph] = page.landscape ? [page.paper[1], page.paper[0]] : page.paper;
  const m = page.margins;
  const w = (pw - m.left - m.right) * DPI;
  const h = (ph - m.top - m.bottom) * DPI;
  const across = span(geometry.cols, range.c0, range.c1);
  const down = span(geometry.rows, range.r0, range.r1);
  let s = 1;
  if (page.fitToWidth > 0 && across > 0) s = Math.min(s, (w * page.fitToWidth) / across);
  if (page.fitToHeight > 0 && down > 0) s = Math.min(s, (h * page.fitToHeight) / down);
  return Math.max(10, Math.floor(s * 100 + 1e-9)) / 100;
}

/** One block of a page: cells r0..r1 by c0..c1 at (x, y) of the zoomed box, clipped to their size, or (`rest`, the
 *  page's own cells) to the rest of the printable area, as text running on past them prints. */
function drawBlock(st, block, x, y, layout, rest = false) {
  // (Which gridlines text running on crosses, from the cells as they are made: cells.js runGaps.)
  const gaps = new Map();
  st.gaps = gaps;
  const w = rest ? layout.area.w / layout.scale - x : span(st.cols, block.c0, block.c1);
  const h = rest ? layout.area.h / layout.scale - y : span(st.rows, block.r0, block.r1);
  const clip = div('oxe-page-block');
  setVars(clip, { '--oxe-bx': x, '--oxe-by': y, '--oxe-bw': w, '--oxe-bh': h });
  const layer = div('oxe-page-cells');
  setVars(layer, { '--oxe-ox': st.cols.start(block.c0), '--oxe-oy': st.rows.start(block.r0) });
  for (let r = block.r0; r <= block.r1; r++) {
    if (!st.rows.size(r)) continue;
    const line = st.sheet.cells.get(r);
    for (let c = block.c0; c <= block.c1; c++) {
      if (!st.cols.size(c)) continue;
      const merge = mergeAt(st, r, c);
      if (merge && (merge.r0 !== r || merge.c0 !== c)) continue;
      const cell = line?.get(c);
      if (cell?.v == null && !merge) continue;
      layer.appendChild(cellElement(st, r, c, cell || { t: 'n', v: null }, line || new Map(), block, 0));
    }
  }
  // (Dashes are laid along the page: where the sheet's 0 falls on it, in the page's own pixels.)
  const s = layout.scale;
  const { fills, edges } = decorations({ ...st, pageOrigin: { x: layout.area.x + (x - st.cols.start(block.c0)) * s, y: layout.area.y + (y - st.rows.start(block.r0)) * s } }, block);
  layer.prepend(...fills);
  layer.append(...edges);
  clip.appendChild(layer);
  // (Page Layout > Sheet Options > Gridlines: Print: a hairline at each row and column edge of the block's cells.)
  if (layout.gridLines) {
    const cw = span(st.cols, block.c0, block.c1);
    const ch = span(st.rows, block.r0, block.r1);
    const line = (cls, vars) => { const e = div(`oxe-print-grid ${cls}`); setVars(e, vars); clip.appendChild(e); };
    for (let r = block.r0; r <= block.r1 + 1; r++) if (r === block.r1 + 1 || st.rows.size(r)) line('oxe-print-grid-h', { '--oxe-gy': st.rows.start(r) - st.rows.start(block.r0), '--oxe-gl': cw });
    // (Text running on over empty cells leaves out the lines it crosses, inside the block: Excel's PDF, runGaps.)
    const y0 = st.rows.start(block.r0);
    for (let c = block.c0; c <= block.c1 + 1; c++) {
      if (c !== block.c1 + 1 && !st.cols.size(c)) continue;
      const inner = c > block.c0 && c <= block.c1 && [...gaps.values()].some((s) => s.has(c));
      const pieces = inner ? lineSegments(st.rows, block.r0, block.r1, (r) => gaps.get(r)?.has(c), y0, 0) : [[y0, ch]];
      for (const [py, ph] of pieces) line('oxe-print-grid-v', { '--oxe-gx': st.cols.start(c) - st.cols.start(block.c0), '--oxe-gt': py - y0, '--oxe-gl': ph });
    }
  }
  return clip;
}

/** One page: its title rows and columns (when the page starts after them), then its cells, in the zoomed box; its
 *  header and footer. */
/** Headings: Print's row numbers and column letters for a page, as Excel prints them (its PDFs of workbooks made in
 *  code): the letters in a row of the default height above the columns `colRuns` ([c0, c1] each, in order), the
 *  numbers in a column `heads.w` wide left of the rows `rowRuns`, each centred and on the cells' baseline (drawn as cells
 *  of the Normal style); rules round them, heavier in the numbers' column, and an outline round it all. */
function drawHeadings(st, heads, colRuns, rowRuns) {
  const out = div('oxe-print-heads');
  const label = (text, x, y, w, h) => {
    const el = div(`oxe-cell oxe-h-center oxe-v-bottom ${st.classOf ? st.classOf(0) : ''}`);
    setVars(el, { '--oxe-x': x + LABEL_SHIFT.x, '--oxe-y': y + LABEL_SHIFT.y, '--oxe-w': w, '--oxe-h': h });
    const t = document.createElement('span');
    t.className = 'oxe-tx';
    t.textContent = text;
    el.appendChild(t);
    out.appendChild(el);
  };
  const rule = (cls, x, y, w, h) => { const e = div(`oxe-print-head-rule ${cls}`); setVars(e, { '--oxe-rx': x, '--oxe-ry': y, '--oxe-rw': w, '--oxe-rh': h }); out.appendChild(e); };
  let x = heads.w;
  for (const [c0, c1] of colRuns) {
    for (let c = c0; c <= c1; c++) {
      const w = st.cols.size(c);
      if (!w) continue;
      label(colName(c), x, 0, w, heads.h);
      rule('oxe-heavy', x, 1.15, 0, heads.h);
      x += w;
    }
  }
  rule('oxe-heavy', x, 1.15, 0, heads.h);
  let y = heads.h;
  for (const [r0, r1] of rowRuns) {
    for (let r = r0; r <= r1; r++) {
      const h = st.rows.size(r);
      if (!h) continue;
      label(String(r + 1), 0, y, heads.w, h);
      rule('oxe-heavy', HEAVY_IN, y, heads.w, 0);
      y += h;
    }
  }
  rule('oxe-heavy', HEAVY_IN, y, heads.w, 0);
  rule('oxe-heavy', heads.w, heads.h + 1.15, 0, y - heads.h + 1.28 - 1.15);
  rule('oxe-heavy', heads.w + HEAVY_IN, heads.h, x - heads.w, 0);
  // (The outline round headings and cells.)
  // (Its sides run on by a rule's thickness, so that the corners close.)
  rule('oxe-heavy', 0, 0, x + 1.28, 0);
  rule('oxe-heavy', 0, 0, 0, y + 1.28);
  rule('oxe-heavy', x, 0, 0, y + 1.28);
  rule('oxe-heavy', 0, y, x + 1.28, 0);
  return out;
}

function drawPage(st, block, layout, k = 0, hf = {}) {
  const pageEl = div('oxe-page');
  setVars(pageEl, { '--oxe-page-w': layout.paper.w, '--oxe-page-h': layout.paper.h });
  const area = div('oxe-page-area');
  // (Center on page, Horizontally and Vertically: each page's cells in the middle of the room between its margins, a
  // point short of it, as Excel's PDFs have them: 1, 3 and 8 columns, portrait and landscape, each 1.2 to 1.3px left of
  // the middle, and as much above it.)
  const { rows: tr0, cols: tc0 } = layout.titles;
  const wide = ((tc0 && block.c0 > tc0[1] ? span(st.cols, tc0[0], tc0[1]) : 0) + span(st.cols, block.c0, block.c1) + (layout.heads ? layout.heads.w : 0)) * layout.scale;
  const tall = ((tr0 && block.r0 > tr0[1] ? span(st.rows, tr0[0], tr0[1]) : 0) + span(st.rows, block.r0, block.r1) + (layout.heads ? layout.heads.h : 0)) * layout.scale;
  const dx = layout.center?.h ? Math.max(0, (layout.area.w - wide) / 2 - ROOM_LESS) : 0;
  const dy = layout.center?.v ? Math.max(0, (layout.area.h - tall) / 2 - ROOM_LESS) : 0;
  setVars(area, { '--oxe-ml': layout.area.x + dx, '--oxe-mt': layout.area.y + dy, '--oxe-aw': layout.area.w - dx, '--oxe-ah': layout.area.h - dy });
  const box = div('oxe-page-zoom');
  setVars(box, { '--oxe-zw': layout.area.w / layout.scale, '--oxe-zh': layout.area.h / layout.scale });
  box.style.setProperty('--oxe-zoom', String(layout.scale)); // (a number: setVars gives pixels)
  // (Scaled, Excel's text sits lower in its row than the scale alone puts it, 1.2px x (1 - scale) on the page, and
  // further into its cell, 1.4px x (1 - scale) (left-aligned text: its 3.2px padding at 100% is 1.4 + 1.8 x scale),
  // measured at 40% to 90%; before the zoom here. Half a pixel more down: zoomed, the browser draws text baselines up to
  // a pixel higher than placed, half a pixel on average.)
  if (layout.scale !== 1) setVars(box, { '--oxe-pshift': (1.2 * (1 - layout.scale) + 0.5) / layout.scale, '--oxe-pshift-x': (1.4 * (1 - layout.scale)) / layout.scale });
  const { rows: tr, cols: tc } = layout.titles;
  const withRows = tr && block.r0 > tr[1];
  const withCols = tc && block.c0 > tc[1];
  const hw = layout.heads ? layout.heads.w : 0;
  const hh = layout.heads ? layout.heads.h : 0;
  const th = (withRows ? span(st.rows, tr[0], tr[1]) : 0) + hh;
  const tw = (withCols ? span(st.cols, tc[0], tc[1]) : 0) + hw;
  if (withRows && withCols) box.appendChild(drawBlock(st, { r0: tr[0], r1: tr[1], c0: tc[0], c1: tc[1] }, hw, hh, layout));
  if (withRows) box.appendChild(drawBlock(st, { r0: tr[0], r1: tr[1], c0: block.c0, c1: block.c1 }, tw, hh, layout));
  if (withCols) box.appendChild(drawBlock(st, { r0: block.r0, r1: block.r1, c0: tc[0], c1: tc[1] }, hw, th, layout));
  box.appendChild(drawBlock(st, block, tw, th, layout, true));
  if (layout.heads) box.appendChild(drawHeadings(st, layout.heads, [...(withCols ? [[tc[0], tc[1]]] : []), [block.c0, block.c1]], [...(withRows ? [[tr[0], tr[1]]] : []), [block.r0, block.r1]]));
  area.appendChild(box);
  pageEl.appendChild(area);
  drawHeaderFooter(pageEl, 'header', k, layout, hf);
  drawHeaderFooter(pageEl, 'footer', k, layout, hf);
  return pageEl;
}

/** Build the printed pages of a sheet into `root` (replacing any from before) and set the paper size for printing;
 *  `looks` gives cell format classes and decorations (cellstyles.js), `normalFamily` the Normal style's typeface (printed
 *  row heights depend on it), `print` its print area and titles (printarea.js printSetup) and, for headers and footers,
 *  `sheet` and `file` (names), `sizePt` (the Normal style's size) and `pictures` (hfpictures.js headerPictures), and
 *  `sparks` (ooxml/sparklines.js sparklineLooks), `tables` (ooxml/tablestyle.js tableLooks). Returns the number of pages. */
export function renderPrint(root, sheet, looks = {}, normalFamily = undefined, print = {}) {
  const screen = sheetGeometry(sheet);
  const auto = autoRowHeights(sheet, looks.deco, screen.cols);
  const full = sheetGeometry(sheet, { print: true, normalFamily, auto });
  // (Cells holding only a sparkline are printed too, as Excel prints them: sparklines.xlsx's column O.)
  const keys = [...(print.sparks?.keys() || [])].map((k) => k.split(',').map(Number));
  const range = print.area || { r0: 0, c0: 0, r1: Math.max(0, sheet.maxRow, ...keys.map((k) => k[0])), c1: Math.max(0, sheet.maxCol, ...keys.map((k) => k[1])) };
  const scale = printScale(sheet, full, range);
  const geometry = scale === 1 ? full : sheetGeometry(sheet, { print: true, normalFamily, auto, scale });
  const layout = { ...paginate(sheet, geometry, { ...print, area: range, scale }), sheet };
  // (A print area of several ranges: each on pages of its own, in the order listed, as Excel prints them (its PDFs:
  // two ranges down or across, three out of order, with print titles over each); the scale is the first range's.)
  if (print.areas?.length > 1) layout.pages = print.areas.flatMap((a) => paginate(sheet, geometry, { ...print, area: a, scale }).pages);
  const st = { sheet, rows: geometry.rows, cols: geometry.cols, screenCols: screen.cols, classOf: looks.classOf, deco: looks.deco, date1904: looks.date1904, color: looks.color, hiddenAt: looks.hiddenAt, runStyle: looks.runStyle, rowDesc: rowDescents(sheet, looks.deco), sparks: print.sparks, tables: print.tables, print: true };
  let el = root.querySelector(':scope > .oxe-print');
  if (!el) {
    el = div('oxe-print');
    el.setAttribute('aria-hidden', 'true');
    root.appendChild(el);
  }
  const hf = { family: normalFamily, sizePt: print.sizePt, sheet: print.sheet, file: print.file, pictures: print.pictures };
  el.replaceChildren(...layout.pages.map((b, k) => drawPage(st, b, layout, k, hf)));
  fitNumbers(el);
  // (The paper size can only be given to the browser in an @page rule.)
  setSheet(document, `@page { size: ${layout.paper.w / DPI}in ${layout.paper.h / DPI}in; margin: 0; }`);
  return layout.pages.length;
}
