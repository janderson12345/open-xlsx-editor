// The Name Box's drop-down: the workbook's names to go to, under the box. Excel's Name Box has an arrow that lists
// the names, and picking one goes to it (Excel's documented behaviour; COM doesn't reach the Name Box, so which names it
// lists isn't measured: the list here is Go To's, goto.js goToNames, the names standing for ranges and the tables).
// Opened by the arrow or Alt+Down in the box; Up and Down move, Enter or a click goes, Esc closes.
import { div } from './dom.js';

let lists = 0;

/** A drop-down for the Name Box `input`, its arrow button put after it; `names()` the list when it opens, `onPick`
 *  the name chosen. The box becomes a combobox that controls the list. */
export function attachNameList(input, { names, onPick }) {
  const id = `oxe-namelist-${++lists}`;
  const list = div('oxe-namelist', 'listbox');
  list.id = id;
  list.hidden = true;
  list.setAttribute('aria-label', 'Names');
  const arrow = document.createElement('button');
  arrow.type = 'button';
  arrow.className = 'oxe-namebox-arrow';
  arrow.tabIndex = -1; // (as Excel's arrow, reached with Alt+Down from the box, not with Tab)
  arrow.setAttribute('aria-label', 'Names');
  arrow.textContent = '▾';
  input.setAttribute('role', 'combobox');
  input.setAttribute('aria-controls', id);
  input.setAttribute('aria-expanded', 'false');
  input.setAttribute('aria-autocomplete', 'none');
  let items = [];
  let at = -1;

  const mark = (k) => {
    at = k;
    items.forEach((el, i) => el.setAttribute('aria-selected', String(i === k)));
    if (k >= 0) { input.setAttribute('aria-activedescendant', items[k].id); items[k].scrollIntoView?.({ block: 'nearest' }); }
    else input.removeAttribute('aria-activedescendant');
  };
  const close = () => {
    list.hidden = true;
    input.setAttribute('aria-expanded', 'false');
    mark(-1);
  };
  const open = () => {
    items = names().map((n, k) => {
      const el = div('oxe-namelist-item', 'option');
      el.id = `${id}-${k}`;
      el.textContent = n;
      return el;
    });
    list.replaceChildren(...items);
    if (!items.length) { close(); return; }
    list.hidden = false;
    input.setAttribute('aria-expanded', 'true');
    mark(Math.max(0, items.findIndex((el) => el.textContent.toLowerCase() === input.value.toLowerCase())));
  };
  const pick = (k) => {
    const name = items[k]?.textContent;
    close();
    if (name != null) { input.value = name; onPick(name); }
  };

  // (The box keeps the focus while the list is clicked.)
  list.addEventListener('mousedown', (e) => e.preventDefault());
  list.addEventListener('click', (e) => { const k = items.indexOf(e.target.closest('.oxe-namelist-item')); if (k >= 0) pick(k); });
  arrow.addEventListener('mousedown', (e) => e.preventDefault());
  arrow.addEventListener('click', () => { if (list.hidden) { input.focus(); open(); } else close(); });
  input.addEventListener('blur', close);
  input.after(arrow, list);
  return {
    get open() { return !list.hidden; },
    /** A key in the box: true when the list took it. */
    key(e) {
      if (e.altKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) { if (list.hidden) open(); else close(); return true; }
      if (list.hidden) return false;
      if (e.key === 'ArrowDown') { mark(Math.min(items.length - 1, at + 1)); return true; }
      if (e.key === 'ArrowUp') { mark(Math.max(0, at - 1)); return true; }
      if (e.key === 'Enter') { pick(at); return true; }
      if (e.key === 'Escape') { close(); return true; }
      return false;
    },
    close,
  };
}
