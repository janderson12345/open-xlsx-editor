// The formula bar: the active cell's name and its contents as they would be edited (a number in full, a formula with
// its = sign). Typing there and pressing Enter enters it into the cell, as typing in the cell does. `onKey` sees each
// key first (pointing.js: F4, and the arrows while a formula goes on on another sheet).
import { div } from './grid.js';

/** A bar at the top of `root`; `onEnter(text)` when Enter is pressed in it, `onCancel()` on Esc, `onKey(e)` (true:
 *  taken) before either, `onFocus()` and `onBlur()`. */
export function createFormulaBar(root, { onEnter, onCancel, onKey = null, onFocus = null, onBlur = null }) {
  const bar = div('oxe-bar');
  const name = document.createElement('output');
  name.className = 'oxe-namebox';
  name.setAttribute('aria-label', 'Active cell');
  const input = document.createElement('input');
  input.className = 'oxe-formula';
  input.type = 'text';
  input.spellcheck = false;
  input.setAttribute('aria-label', 'Formula bar');
  bar.append(name, input);
  root.prepend(bar);
  input.addEventListener('focus', () => onFocus?.());
  input.addEventListener('blur', () => onBlur?.());
  input.addEventListener('keydown', (e) => {
    if (onKey?.(e)) { e.preventDefault(); return; }
    if (e.key === 'Enter') { e.preventDefault(); onEnter(input.value); }
    else if (e.key === 'Escape') { e.preventDefault(); onCancel(); }
  });
  return {
    el: bar,
    /** Show a cell: its name ("B2") and editing text. */
    show(ref, text) {
      name.value = ref;
      if (document.activeElement !== input) input.value = text;
    },
    setReadOnly(flag) { input.readOnly = flag; },
    input,
  };
}
