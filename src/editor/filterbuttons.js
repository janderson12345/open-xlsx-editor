// The AutoFilter's buttons, one in the bottom-right corner of each heading cell of the filter's range, as Excel draws
// them; a column with a filter set shows it. A layer of their own over the cells: clicked with the mouse, and reached
// from the keyboard with Alt+Down on a heading cell (as in Excel), so they stay out of the tab order and away from
// assistive technology, which hears the cell instead.
import { setVars } from './dom.js';

const SIZE = 16;

/** Draw the filter buttons of `st.sheet.autoFilter` that fall in window `win` into `st.filterLayer`; a click calls
 *  `st.onFilter(c, button)`. */
export function drawFilterButtons(st, win) {
  const layer = st.filterLayer;
  if (!layer) return;
  const f = st.sheet?.autoFilter;
  if (!f || !win || f.r0 < win.r0 || f.r0 > win.r1) { layer.replaceChildren(); return; }
  const out = [];
  const bottom = st.rows.start(f.r0 + 1);
  if (st.rows.size(f.r0) < 4) { layer.replaceChildren(); return; }
  for (let c = Math.max(f.c0, win.c0); c <= Math.min(f.c1, win.c1); c++) {
    if (st.cols.size(c) < 4) continue;
    const set = f.columns.has(c);
    const b = document.createElement('button');
    b.type = 'button';
    b.tabIndex = -1;
    b.dataset.col = String(c);
    b.className = set ? 'oxe-filter-btn oxe-on' : 'oxe-filter-btn';
    b.textContent = set ? '▼' : '▾';
    b.title = set ? 'Filtered' : '(Showing All)';
    setVars(b, { '--oxe-x': Math.max(st.cols.start(c), st.cols.start(c + 1) - SIZE - 1), '--oxe-y': Math.max(st.rows.start(f.r0), bottom - SIZE - 1) });
    b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
    b.addEventListener('click', (e) => { e.stopPropagation(); st.onFilter?.(c, b); });
    out.push(b);
  }
  layer.replaceChildren(...out);
}
