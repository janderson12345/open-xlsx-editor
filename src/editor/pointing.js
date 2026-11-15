// Pointing at cells while a formula is typed, as Excel's Point mode does it. Where a reference can go (after =, an
// operator, "(", "," or ":"), the arrow keys (Shift extends, Ctrl jumps over data) and a click or a drag on the grid put
// the reference of a cell or range into the formula, and go on changing it until something else is typed. F4 cycles
// the reference at the caret through $A$1, A$1, $A1 and A1. The formula's references are outlined on the grid, each in
// its colour (Excel's range finder). On another sheet (its tab clicked while the formula is typed) the sheet's name goes
// in at once ('Notes & more'!) and the references pointed at there follow it; the arrows start from that sheet's own
// active cell. A Function Arguments box (attachArgument) is pointed from the same way, its text read as what follows
// "=": a click or the arrows put a reference in it when it is empty, its text selected (as a box is when Tab moves into
// it, the reference replacing it) or the caret is where one can go, and a press on the grid never selects there.
import { tokenize, areaText, sheetPrefix } from '../calc/tokens.js';
import { refName } from '../ooxml/refs.js';
import { step, jump } from './navigate.js';

/** The outlines' colours, in the order references come in the formula. (Excel's own are in no file to measure.) */
export const MARK_COLORS = ['#2f6fbf', '#c0392b', '#7d3c98', '#1e8449', '#b9770e', '#0e8fa8', '#a8327a', '#5d6d7e'];

const BEFORE = '=+-*/^&(,:<>;';
const AFTER = ')+-*/^&,:<>=;';

/** Whether a reference can go at `at` in `text` (a formula): after =, an operator, "(", "," or ":" (spaces passed
 *  over), not inside quotes, and before the end, ")" or an operator. */
export function insertable(text, at) {
  if (!text.startsWith('=')) return false;
  let i = at - 1;
  while (i >= 0 && text[i] === ' ') i--;
  if (i < 0 || !BEFORE.includes(text[i])) return false;
  if ((text.slice(0, at).match(/"/g) || []).length % 2) return false;
  let j = at;
  while (j < text.length && text[j] === ' ') j++;
  return j === text.length || AFTER.includes(text[j]);
}

/** A range's reference text, its corners in any order: "B2", or "B2:D5" top left to bottom right. */
export function rangeText(a, b) {
  const r0 = Math.min(a.r, b.r);
  const c0 = Math.min(a.c, b.c);
  const r1 = Math.max(a.r, b.r);
  const c1 = Math.max(a.c, b.c);
  return r0 === r1 && c0 === c1 ? refName(r0, c0) : `${refName(r0, c0)}:${refName(r1, c1)}`;
}

/** The references of a formula with their place in the text: [{ start, end, sheet, area }]. None if it doesn't read
 *  as tokens yet (an open quote). */
function refsIn(text) {
  let toks;
  try { toks = tokenize(text.slice(1)); } catch { return []; }
  return toks.filter((t) => t.type === 'ref' && t.area && t.book == null).map((t) => ({ start: t.start + 1, end: t.end + 1, prefix: t.prefix, sheet: t.sheet, area: t.area }));
}

/** F4: the reference at or just before `at` cycled A1 > $A$1 > A$1 > $A1 > A1 (columns and rows alone: A:A and $A:$A).
 *  Returns { text, start, end } (the reference's new place), or null if there is none there. */
export function cycleRef(text, at) {
  const ref = refsIn(text).find((x) => at >= x.start && at <= x.end);
  if (!ref) return null;
  const a = ref.area;
  const states = [[false, false], [true, true], [false, true], [true, false]]; // ([column $, row $])
  const now = states.findIndex(([ac, ar]) => ac === a.ac0 && ar === a.ar0);
  const [ac, ar] = states[(now + 1) % 4];
  const next = a.kind === 'cols' ? { ...a, ac0: !a.ac0, ac1: !a.ac0 } : a.kind === 'rows' ? { ...a, ar0: !a.ar0, ar1: !a.ar0 } : { ...a, ac0: ac, ar0: ar, ac1: ac, ar1: ar };
  const piece = ref.prefix + areaText(next);
  return { text: text.slice(0, ref.start) + piece + text.slice(ref.end), start: ref.start, end: ref.start + piece.length };
}

/** Pointing for one editor. `ctx` = { grid, workbook(), shown() (the index of the sheet shown), origin() ({ si, r, c }:
 *  the cell the formula is for), active() ({ r, c }: the active cell of the sheet shown, where the arrows start),
 *  keys() (whether the arrow keys point now: the cell editor in Enter mode, or the formula bar while another sheet is
 *  shown), say(text) (announced) }. */
export function createPointing(ctx) {
  let control = null; // the text box the formula is typed in
  let point = null; // { start, end, anchor: { r, c }, focus: { r, c } } (anchor and focus null: a sheet's name alone)
  let arg = null; // a Function Arguments box's cell, { si, r, c }, while one is followed

  const wb = () => ctx.workbook();
  const limits = () => {
    const g = ctx.grid.geometry();
    return { rows: g.rows.count - 1, cols: g.cols.count - 1 };
  };
  const origin = () => arg || ctx.origin();
  /** The formula followed: the text box's, or an argument box's with "=" before it (and how far its text moved). */
  const text = () => (arg ? `=${control.value}` : control.value);
  const lead = () => (arg ? 1 : 0);
  const formula = () => !!control && text().startsWith('=');
  const collapsed = () => control.selectionStart === control.selectionEnd;
  /** Whether a reference can be pointed at now: one is being pointed at, or the caret is where one can go (in an
   *  argument box, also with its text selected, which the reference replaces). */
  const canPoint = () => {
    if (!formula()) return false;
    if (point) return collapsed() && control.selectionStart === point.end;
    if (arg) { const v = control.value; return insertable(`=${v.slice(0, control.selectionStart)}${v.slice(control.selectionEnd)}`, control.selectionStart + 1); }
    return collapsed() && insertable(control.value, control.selectionStart);
  };

  function refresh() {
    if (!formula()) { ctx.grid.setMarks([]); return; }
    const shown = ctx.shown();
    const home = origin().si;
    const lim = limits();
    const colors = new Map();
    const marks = [];
    const all = text();
    for (const ref of refsIn(all)) {
      const key = all.slice(ref.start, ref.end).replace(/\$/g, '').toUpperCase();
      if (!colors.has(key)) colors.set(key, MARK_COLORS[colors.size % MARK_COLORS.length]);
      if ((ref.sheet == null ? home : wb().sheetIndexOf(ref.sheet)) !== shown) continue;
      const a = ref.area;
      marks.push({
        r0: Math.min(a.r0, a.r1, lim.rows), c0: Math.min(a.c0, a.c1, lim.cols), r1: Math.min(Math.max(a.r0, a.r1), lim.rows), c1: Math.min(Math.max(a.c0, a.c1), lim.cols),
        color: colors.get(key), dashed: !!point && ref.start - lead() === point.start && ref.end - lead() === point.end,
      });
    }
    ctx.grid.setMarks(marks);
  }

  /** Put the range anchor..focus into the formula: in place of the one pointed at, or at the caret. */
  function put(anchor, focus) {
    const shown = ctx.shown();
    const prefix = shown !== origin().si ? sheetPrefix(wb().sheets[shown].name) : '';
    const piece = prefix + rangeText(anchor, focus);
    const start = point ? point.start : control.selectionStart;
    control.setRangeText(piece, start, point ? point.end : control.selectionEnd, 'end');
    point = { start, end: start + piece.length, anchor, focus };
    control.dispatchEvent(new Event('input')); // (the cell editor grows; an untrusted event doesn't end pointing)
    ctx.grid.scrollIntoView(focus.r, focus.c);
    refresh();
    ctx.say?.(piece); // (the text box's value changed under the caret: a screen reader is told what went in)
  }

  const onInput = (e) => { if (e.isTrusted) point = null; refresh(); };
  const onDown = () => { point = null; };

  return {
    /** Follow a text box a formula may be typed in (null: none any more). */
    attach(el) {
      if (control === el && !arg) return;
      if (control) { control.removeEventListener('input', onInput); control.removeEventListener('mousedown', onDown); }
      control = el;
      arg = null;
      point = null;
      if (el) { el.addEventListener('input', onInput); el.addEventListener('mousedown', onDown); }
      refresh();
    },
    /** Follow a Function Arguments box `el`, for the formula of cell `at` ({ si, r, c }); the arrows point from it. */
    attachArgument(el, at) {
      // (The same box focused again, after a sheet's tab was clicked: what it points at is kept.)
      if (el && control === el && arg) return;
      this.attach(el);
      arg = el ? at : null;
      refresh();
    },
    /** Whether an argument box is followed; and pointed from on a sheet other than its cell's. */
    get argument() { return !!arg; },
    get away() { return !!arg && ctx.shown() !== arg.si; },
    get control() { return control; },
    /** Whether a reference can be pointed at now (in the text box followed). */
    get ready() { return canPoint(); },
    /** A key in the text box: arrows point (when `keys()`), F4 cycles $; true if it was taken. */
    key(e) {
      if (!formula()) return false;
      if (e.key === 'F4' && collapsed()) {
        const next = cycleRef(text(), control.selectionStart + lead());
        if (!next) return false;
        control.value = next.text.slice(lead());
        control.setSelectionRange(next.end - lead(), next.end - lead());
        if (point && point.start === next.start - lead()) point.end = next.end - lead();
        control.dispatchEvent(new Event('input'));
        return true;
      }
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      if (!d || !(arg || ctx.keys()) || !canPoint()) return false;
      const sheet = wb().sheet(ctx.shown());
      const from = point?.focus || ctx.active();
      const to = (e.ctrlKey || e.metaKey ? jump : step)(sheet, from.r, from.c, d[0], d[1], limits());
      put(e.shiftKey ? (point?.anchor || from) : to, to);
      return true;
    },
    /** A press on cell (r, c) of the grid: pointed at (Shift: the range from the one pointed at) if a reference can
     *  go; true if it was taken. */
    press(r, c, shift) {
      // (An argument box: a press that can't point does nothing, the box keeping focus, as the sheet isn't edited.)
      if (!canPoint()) return !!arg;
      const m = wb().sheet(ctx.shown()).merges.find((x) => r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1);
      const at = m ? { r: m.r0, c: m.c0 } : { r, c };
      put(shift && point?.anchor ? point.anchor : at, at);
      return true;
    },
    /** Dragging from a press: the range from where it started. */
    drag(r, c) { if (point?.anchor) put(point.anchor, { r, c }); },
    /** Another sheet is shown: where a reference can go, its name goes in (in place of what was pointed at before;
     *  nothing back on the formula's own sheet). */
    sheetShown() {
      if (canPoint()) {
        const shown = ctx.shown();
        const prefix = shown === origin().si ? '' : sheetPrefix(wb().sheets[shown].name);
        const start = point ? point.start : control.selectionStart;
        control.setRangeText(prefix, start, point ? point.end : control.selectionEnd, 'end');
        point = prefix ? { start, end: start + prefix.length, anchor: null, focus: null } : null;
        control.dispatchEvent(new Event('input'));
      }
      refresh();
    },
    /** The outlines again (another sheet shown). */
    refresh,
  };
}
