// Data > Data Validation > Circle Invalid Data: a red oval round each cell that breaks its rule, as Excel draws it.
// Measured over COM (Excel's ovals are shapes there): 3.75pt left of the cell, 1.5pt above it, 7.5pt wider and 3pt
// taller, in any row height or column width; cut off at the sheet's top and left edges (A1's starts at 0, that much
// narrower); a merged area's as wide as the area and as tall as its first row. Line FF0000, 1.35pt, solid, no fill.
// Not saved in the file. Drawn in an SVG layer over the cells (grid.js), in the sheet's coordinates.

const SVG = 'http://www.w3.org/2000/svg';
const PX = 4 / 3;
const LEFT = 3.75 * PX;
const TOP = 1.5 * PX;
const WIDER = 7.5 * PX;
const TALLER = 3 * PX;

/** The SVG layer the circles are drawn in. */
export function circleLayer() {
  const el = document.createElementNS(SVG, 'svg');
  el.setAttribute('class', 'oxe-circles');
  el.setAttribute('aria-hidden', 'true');
  el.setAttribute('overflow', 'visible');
  return el;
}

/** Each circle's box in the sheet's pixels (`rows`, `cols`: the grid's axes; `merges`: the sheet's), null for a
 *  hidden cell: { x, y, w, h }, from the cells' own edge (x 0 the first column's left). */
export function circleBoxes(cells, rows, cols, merges = []) {
  return cells.map(({ r, c }) => {
    const m = merges.find((a) => a.r0 === r && a.c0 === c);
    const c1 = m ? m.c1 : c;
    const w = cols.start(c1 + 1) - cols.start(c);
    const h = rows.size(r);
    if (w <= 0 || h <= 0) return null;
    const x = cols.start(c) - LEFT;
    const y = rows.start(r) - TOP;
    return { x: Math.max(0, x), y: Math.max(0, y), w: w + WIDER - Math.max(0, -x), h: h + TALLER - Math.max(0, -y) };
  });
}

/** Draw `st.circles` ([{ r, c }]) into `st.circleLayer`. */
export function drawCircles(st) {
  const layer = st.circleLayer;
  if (!layer || !st.cols) return;
  const kids = [];
  for (const b of circleBoxes(st.circles || [], st.rows, st.cols, st.sheet?.merges)) {
    if (!b) continue;
    const e = document.createElementNS(SVG, 'ellipse');
    const at = { cx: st.rowHeadW + b.x + b.w / 2, cy: b.y + b.h / 2, rx: b.w / 2, ry: b.h / 2, fill: 'none', stroke: '#ff0000', 'stroke-width': 1.35 * PX };
    for (const [k, v] of Object.entries(at)) e.setAttribute(k, String(v));
    kids.push(e);
  }
  layer.replaceChildren(...kids);
}
