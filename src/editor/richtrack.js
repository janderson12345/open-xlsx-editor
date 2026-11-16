// One font per character of a text box being edited (richmarks.js), kept in step with what is typed: the cell editor's
// text box and the formula bar's each have one, so part of a cell's text can be formatted in either, and a cell of
// rich text keeps its runs whichever one it is edited in.
import { fontsFromRuns, editAt, applyEdit, runsOf } from './richmarks.js';

/** The tracker of `box` (a textarea or input). */
export function trackRich(box) {
  // (`seen`, the text the fonts were last placed in; `sel`, the selection before the edit being made; `kept`, the
  // fonts of texts the box's own undo can bring back.)
  const t = { fonts: [], seen: '', sel: null, kept: new Map() };
  const keep = () => {
    t.kept.set(t.seen, t.fonts);
    if (t.kept.size > 200) t.kept.delete(t.kept.keys().next().value);
  };
  return {
    get fonts() { return t.fonts; },
    /** Start again on `text`, with its runs when it is a cell of rich text. */
    reset(text, runs = null) {
      t.fonts = runs && runs.map((x) => x.text).join('') === text ? fontsFromRuns(runs) : Array(text.length).fill(null);
      t.seen = text;
      t.sel = null;
      t.kept.clear();
      keep();
    },
    /** The selection as it is before an edit (beforeinput, or a line break put in by hand). */
    before() { t.sel = [box.selectionStart, box.selectionEnd]; },
    /** The fonts brought up to the box's text (after an edit, `e` its input event). */
    sync(e = null) {
      const text = box.value;
      if (text !== t.seen) {
        const undo = e && /^history/.test(e.inputType || '') && t.kept.get(text);
        t.fonts = undo || applyEdit(t.fonts, editAt(t.seen, text, t.sel, box.selectionEnd));
        t.seen = text;
        keep();
      }
      t.sel = null;
    },
    /** `change(fonts, start, end)` on the selection (richmarks.js toggleFont, setFont); false when nothing is
     *  selected or the text is a formula. */
    change(fn) {
      if (box.selectionEnd <= box.selectionStart || box.value.startsWith('=')) return false;
      this.sync();
      t.fonts = fn(t.fonts, box.selectionStart, box.selectionEnd);
      keep();
      return true;
    },
    /** The text's runs, or null (a formula, or nothing formatted). */
    runs(text = box.value) {
      this.sync();
      return text.startsWith('=') ? null : runsOf(text, t.fonts);
    },
  };
}

/** Whether focus is going to a control that formats part of the text being edited (the Font group's font and size
 *  lists, marked `data-oxe-keeps-edit`), so the edit stays open while it is used. */
export const keepsEdit = (to) => !!to?.closest?.('[data-oxe-keeps-edit]');
