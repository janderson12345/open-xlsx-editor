// The formula bar: the Name Box (the active cell's name, or the selection's name when it has one; a reference or a name
// typed there and Enter goes to it, as in Excel; its drop-down lists the names to go to, namelist.js) and the cell's contents as they would be edited (a number in full, a formula with its = sign). Typing there
// and pressing Enter enters it into the cell, as typing in the cell does. `onKey` sees each key first (pointing.js: F4,
// and the arrows while a formula goes on on another sheet).
import { autoCorrect } from './autocorrect.js';
import { div } from './dom.js';
import { attachNameList } from './namelist.js';
import { trackRich, keepsEdit } from './richtrack.js';
import { toggleFont, setFont, sizeAt } from './richmarks.js';

/** Ctrl+B, I, U and 5 on part of the text, as in the cell editor. */
const MARK_KEYS = { b: 'b', i: 'i', u: 'u', 5: 'strike' };

/** A bar at the top of `root`; `onEnter(text, runs)` when Enter is pressed in it, `onCancel()` on Esc, `onKey(e)` (true:
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
  // (A cell's text shown plain, as Excel's bar shows it, but its runs kept through the edit, and part of it formatted
  // here as in the cell editor: richtrack.js. `parked`: focus gone to the Font group's font or size list.)
  const track = trackRich(input);
  let parked = false;
  // (The selection kept while a box formats part of the text: focus coming back loses it.)
  let held = null;
  input.addEventListener('beforeinput', () => track.before());
  input.addEventListener('input', (e) => track.sync(e));
  input.addEventListener('blur', (e) => {
    if (!keepsEdit(e.relatedTarget) || input.selectionEnd <= input.selectionStart) return;
    parked = true;
    e.relatedTarget.addEventListener('focusout', () => { parked = false; }, { once: true });
  });
  const change = (fn) => {
    if (held) input.setSelectionRange(held[0], held[1]);
    if ((document.activeElement !== input && !parked) || !track.change(fn)) return false;
    if (parked) { parked = false; input.focus({ preventScroll: true }); }
    return true;
  };
  bar.append(box, input);
  root.prepend(bar);
  input.addEventListener('focus', () => onFocus?.());
  input.addEventListener('blur', () => onBlur?.());
  input.addEventListener('keydown', (e) => {
    if (onKey?.(e)) { e.preventDefault(); return; }
    const markKey = (e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && MARK_KEYS[e.key.toLowerCase()];
    if (markKey && change((fonts, a, b) => toggleFont(fonts, a, b, markKey))) { e.preventDefault(); return; }
    if (e.key === 'Enter') { e.preventDefault(); const text = corrector.finish(); onEnter(text, track.runs(text)); }
    else if (e.key === 'Escape') { e.preventDefault(); onCancel(); }
  });
  return {
    el: bar,
    /** Show a cell: its name ("B2", or the selection's name) and editing text, with its runs when it is rich text. */
    show(ref, text, runs = null) {
      shown = ref;
      if (document.activeElement !== name) name.value = ref;
      if (document.activeElement !== input && !parked) { input.value = text; track.reset(text, runs); corrector.refresh(); }
    },
    /** Bold, Italic, Underline or Strikethrough ('b', 'i', 'u', 'strike') on the part of the text selected here. */
    mark: (key) => change((fonts, a, b) => toggleFont(fonts, a, b, key)),
    /** Font properties ({ name, size, color }, or a function of the first selected size, `base` the cell's) on it. */
    markFont: (patch, base = null) => change((fonts, a, b) => setFont(fonts, a, b, typeof patch === 'function' ? patch(sizeAt(fonts, a, base)) : patch)),
    /** The font of the selection's first character here (null: the cell's own), for Format Cells. */
    get selectedFont() { return track.fonts[held ? held[0] : input.selectionStart] ?? null; },
    /** Kept as it is while a box formats part of its text (`on`), focus gone to the box, or no longer. */
    hold(on) {
      if (!on && held) input.setSelectionRange(held[0], held[1]);
      parked = !!on;
      held = on ? [input.selectionStart, input.selectionEnd] : null;
    },
    /** Whether part of the text is selected here while it is edited. */
    get selecting() { return (document.activeElement === input || parked) && (held ? held[1] > held[0] : input.selectionEnd > input.selectionStart) && !input.value.startsWith('='); },
    setReadOnly(flag) { input.readOnly = flag; },
    input,
  };
}
