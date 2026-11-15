// The formula bar: the Name Box (the active cell's name, or the selection's name when it has one; a reference or a name
// typed there and Enter goes to it, as in Excel; its drop-down lists the names to go to, namelist.js) and the cell's contents as they would be edited (a number in full, a formula with its = sign). Typing there
// and pressing Enter enters it into the cell, as typing in the cell does. `onKey` sees each key first (pointing.js: F4,
// and the arrows while a formula goes on on another sheet).
import { autoCorrect } from './autocorrect.js';
import { div } from './dom.js';
import { attachNameList } from './namelist.js';

/** A bar at the top of `root`; `onEnter(text)` when Enter is pressed in it, `onCancel()` on Esc, `onKey(e)` (true:
 *  taken) before either, `onFocus()` and `onBlur()`; `onGo(text)` when Enter is pressed in the Name Box, `onCancel()`
 *  on Esc there too; `names()` the Name Box's drop-down list. */
export function createFormulaBar(root, { onEnter, onCancel, onGo = null, onKey = null, onFocus = null, onBlur = null, names = () => [] }) {
  const bar = div('oxe-bar');
  const name = document.createElement('input');
  name.className = 'oxe-namebox';
  name.type = 'text';
  name.spellcheck = false;
  name.setAttribute('aria-label', 'Name Box');
  let shown = '';
  name.addEventListener('focus', () => name.select());
  const box = document.createElement('span');
  box.className = 'oxe-namebox-wrap';
  box.append(name);
  const dropDown = attachNameList(name, { names, onPick: (text) => onGo?.(text) });
  name.addEventListener('keydown', (e) => {
    if (dropDown.key(e)) { e.preventDefault(); return; }
    if (e.key === 'Enter') { e.preventDefault(); onGo?.(name.value); }
    else if (e.key === 'Escape') { e.preventDefault(); name.value = shown; onCancel(); }
  });
  name.addEventListener('blur', () => { name.value = shown; });
  const input = document.createElement('input');
  input.className = 'oxe-formula';
  input.type = 'text';
  input.spellcheck = false;
  input.setAttribute('aria-label', 'Formula bar');
  const corrector = autoCorrect(input);
  bar.append(box, input);
  root.prepend(bar);
  input.addEventListener('focus', () => onFocus?.());
  input.addEventListener('blur', () => onBlur?.());
  input.addEventListener('keydown', (e) => {
    if (onKey?.(e)) { e.preventDefault(); return; }
    if (e.key === 'Enter') { e.preventDefault(); onEnter(corrector.finish()); }
    else if (e.key === 'Escape') { e.preventDefault(); onCancel(); }
  });
  return {
    el: bar,
    /** Show a cell: its name ("B2", or the selection's name) and editing text. */
    show(ref, text) {
      shown = ref;
      if (document.activeElement !== name) name.value = ref;
      if (document.activeElement !== input) { input.value = text; corrector.refresh(); }
    },
    setReadOnly(flag) { input.readOnly = flag; },
    input,
  };
}
