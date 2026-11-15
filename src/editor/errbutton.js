// The grid's button layer beside the active cell (grid.js): Excel's error button, with Auto Fill Options after a fill.
import { setVars } from './dom.js';
import { fillOptionsButton } from './fillopts.js';

/** Excel's error button, beside the active cell when it has an error indicator: left of it, or right in column A. */
export function drawErrorButton(st) {
  const layer = st.errLayer;
  if (!layer) return;
  const s = st.sel;
  // (Auto Fill Options after a fill, in the same layer: fillopts.js.)
  const fillButton = fillOptionsButton(st);
  if (!s || !st.errors?.has(`${s.ar},${s.ac}`) || !st.rows.size(s.ar)) { layer.replaceChildren(...(fillButton ? [fillButton] : [])); return; }
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
  layer.replaceChildren(b, ...(fillButton ? [fillButton] : []));
}
