// The printed sheet, as Excel prints it (measured against its PDFs, docs/LAYOUT.md and docs/PRINT.md): its print area
// (or used range) split into pages (columns that fit across the printable width, rows that fit down its height, at its
// manual breaks, pages in the sheet's page order), the print title rows and columns repeated on every page, all at
// the sheet's scale or the scale that fits it on the pages asked for, each page at its paper size with the sheet's
// margins. Built only for printing (and for PDF export and tools/excel-compare); hidden on screen.
import { sheetGeometry, cellElement, mergeAt, setVars, div, fitNumbers } from './grid.js';
import { decorations } from './decor.js';
import { autoRowHeights, rowDescents } from './rowheights.js';
import { fontStack } from './fonts.js';
import { parseHeaderFooter, fillFields } from '../ooxml/headerfooter.js';

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
  const roomDown = (area.h - ROOM_LESS) / scale;
  const roomAcross = (area.w - ROOM_LESS) / scale;
  const rowRuns = breakAxis(rows, range.r1, roomDown, { from: range.r0, breaks: breaks.rows, roomAt: (start) => roomDown - (titles.rows && start > titles.rows[1] ? titleRows : 0) });
  const colRuns = breakAxis(cols, range.c1, roomAcross, { from: range.c0, breaks: breaks.cols, roomAt: (start) => roomAcross - (titles.cols && start > titles.cols[1] ? titleCols : 0) });
  const pages = [];
  const add = ([r0, r1], [c0, c1]) => pages.push({ r0, r1, c0, c1 });
  if (page.pageOrder === 'overThenDown') for (const rr of rowRuns) for (const cr of colRuns) add(rr, cr);
  else for (const cr of colRuns) for (const rr of rowRuns) add(rr, cr);
  return { pages, paper: { w: pw * DPI, h: ph * DPI }, area, scale, titles };
}

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
  return clip;
}

/**
 * A page's header or footer (`which`), as Excel prints it (print-headers.xlsx): in the Normal style's font and size
 * (scaled with the sheet unless scaleWithDoc is off), the header's baseline 1.05em below its margin, the footer's 0.436em
 * above its margin; the left section from the left margin and the right one to the right margin, each 1.25px in, the
 * centre centred on the page.
 */
function drawHeaderFooter(pageEl, which, k, layout, hf) {
  const { page } = layout.sheet;
  const setup = page.headerFooter;
  if (!setup) return;
  const set = k === 0 && setup.differentFirst ? setup.first : setup.differentOddEven && k % 2 === 1 ? setup.even : setup.odd;
  const sections = parseHeaderFooter(set[which]);
  if (!sections.left.length && !sections.center.length && !sections.right.length) return;
  const scale = setup.scaleWithDoc ? layout.scale : 1;
  const basePt = hf.sizePt || 11;
  const px = (pt) => (pt * 96) / 72 * scale;
  const em = px(basePt);
  const m = page.margins;
  const y = which === 'header' ? m.header * DPI + 1.05 * em : layout.paper.h - m.footer * DPI - 0.436 * em;
  const ctx = (drawHeaderFooter.canvas ||= document.createElement('canvas')).getContext('2d');
  ctx.font = `${em}px ${fontStack(hf.family || 'Calibri')}`;
  const met = ctx.measureText('Hg');
  const asc = met.fontBoundingBoxAscent ?? em * 0.95;
  const desc = met.fontBoundingBoxDescent ?? em * 0.25;
  const inset = 1.25;
  const left = setup.alignWithMargins ? m.left * DPI + inset : inset;
  const right = layout.paper.w - (setup.alignWithMargins ? m.right * DPI + inset : inset);
  for (const [name, runs] of Object.entries(sections)) {
    if (!runs.length) continue;
    const el = div(`oxe-hf oxe-hf-${name}`);
    setVars(el, { '--oxe-hf-top': y - asc, '--oxe-hf-lh': asc + desc, '--oxe-hf-left': name === 'left' ? left : name === 'right' ? right : layout.paper.w / 2 });
    el.style.setProperty('--oxe-hf-font', fontStack(hf.family || 'Calibri'));
    el.style.setProperty('--oxe-hf-size', `${em}px`);
    for (const r of fillFields(runs, { page: k + 1, pages: layout.pages.length, sheet: hf.sheet, file: hf.file })) {
      const span = document.createElement('span');
      span.textContent = r.text;
      const cls = [r.bold && 'oxe-hf-b', r.italic && 'oxe-hf-i', r.underline && 'oxe-hf-u', r.strike && 'oxe-hf-s'].filter(Boolean);
      if (cls.length) span.className = cls.join(' ');
      if (r.size) span.style.setProperty('--oxe-hf-run-size', `${px(r.size)}px`);
      if (r.font) span.style.setProperty('--oxe-hf-run-font', fontStack(r.font));
      if (r.color) span.style.setProperty('--oxe-hf-run-color', r.color);
      el.appendChild(span);
    }
    pageEl.appendChild(el);
  }
}

/** One page: its title rows and columns (when the page starts after them), then its cells, in the zoomed box; its
 *  header and footer. */
function drawPage(st, block, layout, k = 0, hf = {}) {
  const pageEl = div('oxe-page');
  setVars(pageEl, { '--oxe-page-w': layout.paper.w, '--oxe-page-h': layout.paper.h });
  const area = div('oxe-page-area');
  setVars(area, { '--oxe-ml': layout.area.x, '--oxe-mt': layout.area.y, '--oxe-aw': layout.area.w, '--oxe-ah': layout.area.h });
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
  const th = withRows ? span(st.rows, tr[0], tr[1]) : 0;
  const tw = withCols ? span(st.cols, tc[0], tc[1]) : 0;
  if (withRows && withCols) box.appendChild(drawBlock(st, { r0: tr[0], r1: tr[1], c0: tc[0], c1: tc[1] }, 0, 0, layout));
  if (withRows) box.appendChild(drawBlock(st, { r0: tr[0], r1: tr[1], c0: block.c0, c1: block.c1 }, tw, 0, layout));
  if (withCols) box.appendChild(drawBlock(st, { r0: block.r0, r1: block.r1, c0: tc[0], c1: tc[1] }, 0, th, layout));
  box.appendChild(drawBlock(st, block, tw, th, layout, true));
  area.appendChild(box);
  pageEl.appendChild(area);
  drawHeaderFooter(pageEl, 'header', k, layout, hf);
  drawHeaderFooter(pageEl, 'footer', k, layout, hf);
  return pageEl;
}

/** Build the printed pages of a sheet into `root` (replacing any from before) and set the paper size for printing;
 *  `looks` gives cell format classes and decorations (cellstyles.js), `normalFamily` the Normal style's typeface (printed
 *  row heights depend on it), `print` its print area and titles (printarea.js printSetup) and, for headers and footers,
 *  `sheet` and `file` (names) and `sizePt` (the Normal style's size). Returns the number of pages. */
export function renderPrint(root, sheet, looks = {}, normalFamily = undefined, print = {}) {
  const screen = sheetGeometry(sheet);
  const auto = autoRowHeights(sheet, looks.deco, screen.cols);
  const full = sheetGeometry(sheet, { print: true, normalFamily, auto });
  const range = print.area || { r0: 0, c0: 0, r1: Math.max(0, sheet.maxRow), c1: Math.max(0, sheet.maxCol) };
  const scale = printScale(sheet, full, range);
  const geometry = scale === 1 ? full : sheetGeometry(sheet, { print: true, normalFamily, auto, scale });
  const layout = { ...paginate(sheet, geometry, { ...print, scale }), sheet };
  const st = { sheet, rows: geometry.rows, cols: geometry.cols, screenCols: screen.cols, classOf: looks.classOf, deco: looks.deco, date1904: looks.date1904, color: looks.color, rowDesc: rowDescents(sheet, looks.deco), print: true };
  let el = root.querySelector(':scope > .oxe-print');
  if (!el) {
    el = div('oxe-print');
    el.setAttribute('aria-hidden', 'true');
    root.appendChild(el);
  }
  const hf = { family: normalFamily, sizePt: print.sizePt, sheet: print.sheet, file: print.file };
  el.replaceChildren(...layout.pages.map((b, k) => drawPage(st, b, layout, k, hf)));
  fitNumbers(el);
  // (The paper size can only be given to the browser in an @page rule.)
  let style = document.getElementById('oxe-page-style');
  if (!style) {
    style = document.createElement('style');
    style.id = 'oxe-page-style';
    document.head.appendChild(style);
  }
  style.textContent = `@page { size: ${layout.paper.w / DPI}in ${layout.paper.h / DPI}in; margin: 0; }`;
  return layout.pages.length;
}
