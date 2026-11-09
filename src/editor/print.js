// The printed sheet: the used range split into pages as Excel splits it (columns that fit across the printable width,
// rows that fit down its height, pages in the sheet's page order), each page at its paper size with the sheet's
// margins. Built only for printing (and for tools/excel-compare); hidden on screen. No scaling, fit-to-page, print
// titles, manual breaks, headers or footers yet: those come with the print model. Everything here is to be measured
// against Excel's own PDFs.
import { sheetGeometry, cellElement, mergeAt, setVars, div } from './grid.js';

const DPI = 96;

/** Split an axis's indexes 0..last into runs that each fit in `room` pixels (a run always takes at least one). */
export function breakAxis(axis, last, room) {
  const runs = [];
  let start = 0;
  let used = 0;
  for (let i = 0; i <= last; i++) {
    const size = axis.size(i);
    if (used + size > room && i > start) {
      runs.push([start, i - 1]);
      start = i;
      used = 0;
    }
    used += size;
  }
  runs.push([start, last]);
  return runs;
}

/** The pages of a sheet: each { r0, r1, c0, c1 } in printing order, with the paper and printable area in pixels. */
export function paginate(sheet, geometry = sheetGeometry(sheet)) {
  const { page } = sheet;
  const [pw, ph] = page.landscape ? [page.paper[1], page.paper[0]] : page.paper;
  const m = page.margins;
  const area = { x: m.left * DPI, y: m.top * DPI, w: (pw - m.left - m.right) * DPI, h: (ph - m.top - m.bottom) * DPI };
  const lastRow = Math.max(0, sheet.maxRow);
  const lastCol = Math.max(0, sheet.maxCol);
  const rowRuns = breakAxis(geometry.rows, lastRow, area.h);
  const colRuns = breakAxis(geometry.cols, lastCol, area.w);
  const pages = [];
  const add = ([r0, r1], [c0, c1]) => pages.push({ r0, r1, c0, c1 });
  if (page.pageOrder === 'overThenDown') for (const rr of rowRuns) for (const cr of colRuns) add(rr, cr);
  else for (const cr of colRuns) for (const rr of rowRuns) add(rr, cr);
  return { pages, paper: { w: pw * DPI, h: ph * DPI }, area };
}

/** One page's cells, at their sheet coordinates inside a layer shifted so the page's first cell sits at the margin. */
function drawPage(st, block, layout) {
  const pageEl = div('oxe-page');
  setVars(pageEl, { '--oxe-page-w': layout.paper.w, '--oxe-page-h': layout.paper.h });
  const area = div('oxe-page-area');
  setVars(area, { '--oxe-ml': layout.area.x, '--oxe-mt': layout.area.y, '--oxe-aw': layout.area.w, '--oxe-ah': layout.area.h });
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
  area.appendChild(layer);
  pageEl.appendChild(area);
  return pageEl;
}

/** Build the printed pages of a sheet into `root` (replacing any from before) and set the paper size for printing.
 *  Returns the number of pages. */
export function renderPrint(root, sheet) {
  const geometry = sheetGeometry(sheet);
  const layout = paginate(sheet, geometry);
  const st = { sheet, rows: geometry.rows, cols: geometry.cols };
  let el = root.querySelector(':scope > .oxe-print');
  if (!el) {
    el = div('oxe-print');
    el.setAttribute('aria-hidden', 'true');
    root.appendChild(el);
  }
  el.replaceChildren(...layout.pages.map((b) => drawPage(st, b, layout)));
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
