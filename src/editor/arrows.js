// Excel's tracer arrows (Formulas > Trace Precedents, Trace Dependents): from what is read to what reads it, a dot at
// the start and an arrowhead at the end; a range read is boxed, the arrow starting in its top-left cell; blue, red from
// a cell holding an error; what is on another sheet a dashed black arrow from (or to) a worksheet icon beside the cell.
// Drawn in an SVG layer over the cells (grid.js), in the sheet's coordinates; ooxml/trace.js says what to draw.

const SVG = 'http://www.w3.org/2000/svg';
const COLORS = { blue: '#0000ff', red: '#ff0000', black: '#000000' };
const svg = (name, attrs) => {
  const el = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
};

/** The SVG layer the arrows are drawn in. */
export function arrowLayer() {
  const el = svg('svg', { class: 'oxe-arrows', 'aria-hidden': 'true', overflow: 'visible' });
  return el;
}

/** Draw `st.arrows` ([{ from: rect or null (another sheet), to: rect or null, color }]): a rect's top-left cell is
 *  where the arrow starts or ends, a rect of more than one cell boxed. */
export function drawArrows(st) {
  const layer = st.arrowLayer;
  if (!layer || !st.cols) return;
  const kids = [];
  const centre = (r, c) => ({ x: st.rowHeadW + st.cols.start(c) + st.cols.size(c) / 2, y: st.rows.start(r) + st.rows.size(r) / 2 });
  // (The worksheet icon: up and to the left of the cell it stands for, where Excel puts it.)
  const iconBy = (p) => ({ x: Math.max(st.rowHeadW + 10, p.x - 60), y: Math.max(8, p.y - 28) });
  for (const a of st.arrows || []) {
    const color = COLORS[a.color] || COLORS.blue;
    for (const rect of [a.from, a.to]) {
      if (rect && (rect.r1 > rect.r0 || rect.c1 > rect.c0)) {
        const x = st.rowHeadW + st.cols.start(rect.c0);
        const y = st.rows.start(rect.r0);
        kids.push(svg('rect', { x, y, width: st.cols.start(rect.c1 + 1) - st.cols.start(rect.c0), height: st.rows.start(rect.r1 + 1) - y, fill: 'none', stroke: color, 'stroke-width': 1 }));
      }
    }
    const end = a.to ? centre(a.to.r0, a.to.c0) : null;
    const start = a.from ? centre(a.from.r0, a.from.c0) : null;
    const p = start || iconBy(end);
    const q = end || iconBy(start);
    if (!a.from || !a.to) kids.push(sheetIcon(a.from ? q : p));
    const dx = q.x - p.x;
    const dy = q.y - p.y;
    const len = Math.hypot(dx, dy) || 1;
    const ux = dx / len;
    const uy = dy / len;
    const tip = { x: q.x - ux * 2, y: q.y - uy * 2 };
    kids.push(svg('line', { x1: p.x, y1: p.y, x2: tip.x - ux * 6, y2: tip.y - uy * 6, stroke: color, 'stroke-width': 1, 'stroke-dasharray': a.from && a.to ? 'none' : '4 3' }));
    if (a.from) kids.push(svg('circle', { cx: p.x, cy: p.y, r: 2.5, fill: color }));
    const back = { x: tip.x - ux * 9, y: tip.y - uy * 9 };
    kids.push(svg('polygon', { points: `${tip.x},${tip.y} ${back.x - uy * 3.5},${back.y + ux * 3.5} ${back.x + uy * 3.5},${back.y - ux * 3.5}`, fill: color }));
  }
  layer.replaceChildren(...kids);
}

/** Excel's worksheet icon: a small grid, its top band filled. */
function sheetIcon(at) {
  const g = svg('g', { transform: `translate(${at.x - 8},${at.y - 6})` });
  g.append(svg('rect', { width: 16, height: 12, fill: '#fff', stroke: '#000', 'stroke-width': 1 }),
    svg('rect', { width: 16, height: 3, fill: '#000' }),
    svg('line', { x1: 0, y1: 7.5, x2: 16, y2: 7.5, stroke: '#000', 'stroke-width': 0.6 }),
    svg('line', { x1: 5.5, y1: 3, x2: 5.5, y2: 12, stroke: '#000', 'stroke-width': 0.6 }),
    svg('line', { x1: 10.5, y1: 3, x2: 10.5, y2: 12, stroke: '#000', 'stroke-width': 0.6 }));
  return g;
}
