// Excel's Auto Fill Options button: shown at the bottom-right of the cells a fill handle has just filled, until the
// next change or the selection moves off them; clicked, its menu (index.js fillOptionsMenu) fills them again another
// way: Copy Cells, Fill Series, Fill Formatting Only, Fill Without Formatting, and for dates Fill Days ... Fill Years.
import { setVars } from './dom.js';

/** The button for the filled area `st.fillOptions` ({ r0, c0, r1, c1 }), or null when there is none or it isn't in
 *  view. Clicked: st.onFillOptions({ x, y }). */
export function fillOptionsButton(st) {
  const a = st.fillOptions;
  if (!a || !st.rows.size(a.r1) || !st.cols.size(a.c1)) return null;
  const b = document.createElement('button');
  b.type = 'button';
  b.tabIndex = -1;
  b.className = 'oxe-err-btn oxe-fillopt-btn';
  b.textContent = '▤';
  b.title = 'Auto Fill Options';
  b.setAttribute('aria-label', 'Auto Fill Options');
  setVars(b, { '--oxe-x': st.rowHeadW + st.cols.start(a.c1 + 1) + 2, '--oxe-y': st.rows.start(a.r1 + 1) + 2 });
  b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
  b.addEventListener('click', (e) => { e.stopPropagation(); const box = b.getBoundingClientRect(); st.onFillOptions?.({ x: box.left, y: box.bottom }); });
  return b;
}

/** The menu's choices for cells filled from `cells` (dates or not): [{ type, label }], Excel's words. */
export function fillOptionChoices(dated) {
  return [['copy', 'Copy Cells'], ['series', 'Fill Series'], ['formats', 'Fill Formatting Only'], ['values', 'Fill Without Formatting'],
    ...(dated ? [['days', 'Fill Days'], ['weekdays', 'Fill Weekdays'], ['months', 'Fill Months'], ['years', 'Fill Years']] : [])]
    .map(([type, label]) => ({ type, label }));
}
