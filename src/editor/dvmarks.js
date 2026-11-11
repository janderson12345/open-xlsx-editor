// A data validation list's dropdown button, beside the active cell's right edge as Excel draws it, and the rule's
// input message in a box under the cell. Clicked with the mouse; Alt+Down opens the list from the keyboard, and the
// message is read with the cell, so both stay out of the tab order and away from assistive technology.
import { div, setVars } from './dom.js';

/** Draw `st.validation` ({ r, c, list, prompt }) into `st.dvLayer`; the button calls `st.onDropdown(button)`. */
export function drawValidation(st) {
  const layer = st.dvLayer;
  if (!layer) return;
  const v = st.validation;
  if (!v || !st.rows || st.rows.size(v.r) < 1 || st.cols.size(v.c) < 1) { layer.replaceChildren(); return; }
  const out = [];
  const right = st.cols.start(v.c + 1);
  const top = st.rows.start(v.r);
  const h = st.rows.size(v.r);
  if (v.list) {
    const b = document.createElement('button');
    b.type = 'button';
    b.tabIndex = -1;
    b.className = 'oxe-dv-btn';
    b.textContent = '▾';
    setVars(b, { '--oxe-x': right + 1, '--oxe-y': top, '--oxe-h': h });
    b.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
    b.addEventListener('click', (e) => { e.stopPropagation(); st.onDropdown?.(b); });
    out.push(b);
  }
  if (v.prompt) {
    const box = div('oxe-dv-prompt');
    setVars(box, { '--oxe-x': st.cols.start(v.c) + 12, '--oxe-y': top + h + 4 });
    if (v.prompt.title) { const t = document.createElement('strong'); t.textContent = v.prompt.title; box.append(t); }
    if (v.prompt.text) { const t = div('oxe-dv-prompt-text'); t.textContent = v.prompt.text; box.append(t); }
    out.push(box);
  }
  layer.replaceChildren(...out);
}
