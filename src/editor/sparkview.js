// A sparkline drawn in its cell (or its merged range) as an SVG of the cell's size, under the cell's text, on screen
// and in print alike (decor.js); the PDF export draws the same paths (pdf.js). Shapes from sparkgeom.js.
import { setVars } from './dom.js';
import { sparkShapes, sparkPaths } from './sparkgeom.js';

const SVG = 'http://www.w3.org/2000/svg';

/** The element for sparkline `sp` over cells (r0, c0) to (r1, c1), offset by (ox, oy) as decor.js places fills. */
export function sparkElement(st, m, sp, ox, oy) {
  const x = st.cols.start(m.c0) + ox;
  const y = st.rows.start(m.r0) + oy;
  const w = st.cols.start(m.c1 + 1) - st.cols.start(m.c0);
  const h = st.rows.start(m.r1 + 1) - st.rows.start(m.r0);
  const svg = document.createElementNS(SVG, 'svg');
  svg.setAttribute('class', 'oxe-fill oxe-spark');
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  svg.setAttribute('aria-hidden', 'true');
  setVars(svg, { '--oxe-x': x, '--oxe-y': y, '--oxe-w': w, '--oxe-h': h });
  for (const p of sparkPaths(sparkShapes(sp, w, h))) {
    const path = document.createElementNS(SVG, 'path');
    path.setAttribute('d', p.d);
    if (p.fill) path.setAttribute('fill', p.fill);
    else {
      path.setAttribute('fill', 'none');
      path.setAttribute('stroke', p.stroke);
      path.setAttribute('stroke-width', String(p.width));
      path.setAttribute('stroke-linejoin', 'round');
    }
    svg.appendChild(path);
  }
  return svg;
}
