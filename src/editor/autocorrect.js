// AutoCorrect as a cell's text is typed, Excel's options as it ships them on (Application.AutoCorrect read over COM:
// ReplaceText, CorrectSentenceCap, CapitalizeNamesOfDays, TwoInitialCapitals all on): the replacement list's entries
// (autocorrect-list.js), a sentence's first word capitalised after a full stop, question or exclamation mark (not the
// cell's first word: Excel leaves that), a day's name capitalised, and TWo INitial CApitals put right. A word is
// corrected when it ends (a space or punctuation typed after it) and when the cell is entered. Not in a formula. Each
// correction goes through the text box's own undo, so Ctrl+Z while typing gives back what was typed, as in Excel.
// Also the browser's spelling underlines while text (not a formula) is typed. Typing in Excel can't be driven over COM
// without keystrokes, so when a correction happens is from Microsoft's documentation, not measured.
import { REPLACEMENTS } from './autocorrect-list.js';

const DAYS = new Set(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']);
const ENDS = /[\s.,;:!?)\]"']/;
const own = (k) => (Object.hasOwn(REPLACEMENTS, k) ? REPLACEMENTS[k] : null);

/** A formula's text (=, or + and - as Excel takes them): left alone. */
const isFormula = (text) => /^[=+]/.test(text) || /^-[^\d\s]/.test(text);

/**
 * The correction of the word just before position `end` of `text` (a word ended there): { from, to, text } as offsets
 * into `text`, or null. Exported for the tests.
 */
export function correction(text, end = text.length) {
  if (isFormula(text)) return null;
  const before = text.slice(0, end);
  // (Symbols such as (c) and ... are entries too, ending at a bracket or a dot.)
  for (const k of ['(c)', '(e)', '(r)', '(tm)', '...']) {
    if (before.toLowerCase().endsWith(k) && own(k)) return { from: end - k.length, to: end, text: own(k) };
  }
  const m = /(^|[\s(["'])([^\s(["']+)$/u.exec(before);
  if (!m) return null;
  const word = m[2].replace(/[.,;:!?)\]]+$/, '');
  if (!word || !/\p{L}/u.test(word)) return null;
  const from = before.length - m[2].length;
  const to = from + word.length;
  const lower = word.toLowerCase();
  let fixed = own(word) ?? (word === lower ? null : own(lower) && caseLike(word, own(lower)));
  if (fixed == null && DAYS.has(lower) && word === lower) fixed = word[0].toUpperCase() + word.slice(1);
  // (TWo INitial CApitals: two capitals then lower case, at least three letters.)
  if (fixed == null && /^\p{Lu}{2}\p{Ll}+$/u.test(word) && word.length > 2) fixed = word[0] + word.slice(1).toLowerCase();
  const out = fixed ?? word;
  // (A sentence's first word: after . ? or ! and a space, not the cell's first word.)
  const lead = before.slice(0, from);
  if (/[.?!]\s+$/.test(lead) && /^\p{Ll}/u.test(out)) fixed = out[0].toUpperCase() + out.slice(1);
  return fixed != null && fixed !== word ? { from, to, text: fixed } : null;
}

/** `fixed` with `word`'s capital first letter, as Excel corrects "Teh" to "The". */
function caseLike(word, fixed) {
  return /^\p{Lu}/u.test(word) ? fixed[0].toUpperCase() + fixed.slice(1) : fixed;
}

/** A text box (the cell editor's textarea, the formula bar's input) corrected as it is typed in, and underlined for
 *  spelling while it holds text. Returns `finish()`, which corrects the last word (call before the text is taken). */
export function autoCorrect(box) {
  let busy = false;
  const replace = ({ from, to, text }, caret) => {
    busy = true;
    const [s, e] = [box.selectionStart, box.selectionEnd];
    box.setSelectionRange(from, to);
    // (execCommand keeps the change in the box's own undo; setRangeText where it isn't there.)
    const done = typeof document.execCommand === 'function' && document.execCommand('insertText', false, text);
    if (!done) box.setRangeText(text, from, to, 'end');
    const shift = text.length - (to - from);
    box.setSelectionRange(caret == null ? s + shift : caret + shift, caret == null ? e + shift : caret + shift);
    busy = false;
  };
  box.addEventListener('input', (e) => {
    box.spellcheck = !isFormula(box.value);
    if (busy || e.inputType !== 'insertText' || !e.data || !ENDS.test(e.data.slice(-1))) return;
    const caret = box.selectionStart;
    // (The word before the character just typed; for a bracket or dot the symbol entries end at it.)
    const fix = /[).]/.test(e.data.slice(-1)) ? correction(box.value, caret) || correction(box.value, caret - 1) : correction(box.value, caret - 1);
    if (fix) replace(fix, caret);
  });
  return {
    finish() {
      const fix = correction(box.value, box.value.length);
      if (fix) replace(fix, box.value.length);
      return box.value;
    },
    /** Underlines set for what the box holds now (after its text was set from code). */
    refresh() { box.spellcheck = !!box.value && !isFormula(box.value); },
  };
}
