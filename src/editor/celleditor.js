// Editing a cell in place: a text box over the active cell in the cell's own font, which grows with what is typed.
// Two modes, as in Excel: typing over a cell (enter mode, where the arrow keys finish and move) and F2 or a double
// click (edit mode, where they move the caret); F2 switches between them, as in Excel. Enter, Tab and their Shift forms
// finish and move; Esc gives up; Alt+Enter is a line break. `onKey` sees each key first (pointing.js takes the arrows
// while a formula is pointing). Part of the text formatted (a rich cell, or Ctrl+B on a selection) is drawn formatted
// as it is typed (richview.js); a formula, or text with nothing formatted, is the text box's own plain text.
import { setVars } from './dom.js';
import { autoCorrect } from './autocorrect.js';
import { anyFormatted, toggleFont, setFont, sizeAt } from './richmarks.js';
import { trackRich, keepsEdit } from './richtrack.js';
import { createRichView } from './richview.js';

const MOVES = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/** `grid` (its overlay and cell boxes), `onFinish(text, [dr, dc], runs)`, `onCancel()`, `onKey(e)` (true: taken),
 *  `onClose()` and `runStyle()` (the workbook's run CSS, cellstyles.js). */
export function createCellEditor({ grid, onFinish, onCancel, onKey = null, onClose = null, runStyle = () => null }) {
  const box = document.createElement('textarea');
  box.className = 'oxe-editor';
  box.spellcheck = false;
  box.setAttribute('aria-label', 'Cell contents');
  // (AutoCorrect and the spelling underlines while text is typed: autocorrect.js.)
  const corrector = autoCorrect(box);
  let open = null; // { r, c, mode }
  let held = null; // (a box over it, Insert Function's: losing focus doesn't finish; the caret kept: [start, end])
  // (Part of the text formatted: one font per character, richtrack.js.)
  const rich = { track: trackRich(box), view: createRichView(box, (f) => runStyle()?.(f) || '') };
  // (Focus gone to the Font group's font or size list while text is selected: the edit kept open for it.)
  let parked = false;

  /** The fonts brought up to the box's text (after an edit, `e` its input event), and the view drawn. */
  const sync = (e = null) => {
    rich.track.sync(e);
    show();
  };
  /** Drawn formatted (the box's text hidden, the view under it) when part of it is and it isn't a formula. */
  const show = () => {
    const on = !!open && !box.value.startsWith('=') && anyFormatted(rich.track.fonts);
    box.classList.toggle('oxe-editor-rich', on);
    rich.view.el.className = `oxe-editor oxe-edit-view ${open?.cls || ''}`;
    if (on) {
      // (The browser's spelling underlines would be under the box's own, differently placed, characters.)
      box.spellcheck = false;
      if (!rich.view.el.isConnected) box.before(rich.view.el);
      rich.view.draw(box.value, rich.track.fonts);
    } else rich.view.el.remove();
    if (open) size();
    if (on) rich.view.place();
  };
  const placeMarks = () => { if (open && rich.view.el.isConnected) requestAnimationFrame(() => rich.view.el.isConnected && rich.view.place()); };

  const size = () => {
    // (As wide as the text needs, never narrower than the cell; as tall as its lines: the drawn runs' when shown.)
    const drawn = rich.view.el.isConnected ? rich.view : null;
    for (const el of [box, drawn?.el]) el?.style.setProperty('--oxe-edit-w', '0px');
    const w = Math.max(open.w, box.scrollWidth + 4, drawn ? drawn.scrollWidth + 4 : 0);
    const h = Math.max(open.h, box.scrollHeight, drawn ? drawn.scrollHeight : 0);
    for (const el of [box, drawn?.el]) if (el) setVars(el, { '--oxe-x': open.x, '--oxe-y': open.y, '--oxe-edit-w': w, '--oxe-edit-h': h });
  };
  const finish = (move) => {
    if (!open) return;
    const text = corrector.finish();
    sync();
    // (Its runs, when part of it is formatted and it isn't a formula.)
    const runs = rich.track.runs(text);
    close();
    onFinish(text, move, runs);
  };
  /** The selection's part of the text given bold ('b'), italic ('i'), underline ('u') or strikethrough ('strike'),
   *  or not; true when there was a selection to format. */
  const mark = (key) => change((fonts, a, b) => toggleFont(fonts, a, b, key));
  /** `fn(fonts, start, end)` on the selection's part of the text; true when there was a selection to format. Focus
   *  comes back to the box from a font list that had it (parked). */
  const change = (fn) => {
    // (Held for a box, Format Cells': the selection it was opened on, which focus coming back loses.)
    if (open && held) box.setSelectionRange(held[0], held[1]);
    if (!open || !rich.track.change(fn)) return false;
    if (parked) { parked = false; box.focus({ preventScroll: true }); }
    show();
    return true;
  };

  box.addEventListener('keydown', (e) => {
    if (!open) return;
    if (onKey?.(e)) { e.preventDefault(); return; }
    if (e.key === 'F2') { e.preventDefault(); open.mode = open.mode === 'enter' ? 'edit' : 'enter'; return; }
    // (Ctrl+B, I, U and 5 on part of the text: that part formatted, as Excel's.)
    const fmtKey = (e.ctrlKey || e.metaKey) && !e.altKey && { b: 'b', i: 'i', u: 'u', 5: 'strike' }[e.key.toLowerCase()];
    if (fmtKey && mark(fmtKey)) { e.preventDefault(); return; }
    if (e.key === 'Escape') { e.preventDefault(); close(); onCancel(); return; }
    if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      rich.track.before();
      box.setRangeText('\n', box.selectionStart, box.selectionEnd, 'end');
      sync();
      return;
    }
    if (e.key === 'Enter') { e.preventDefault(); finish(e.shiftKey ? [-1, 0, 'enter'] : [1, 0, 'enter']); return; }
    if (e.key === 'Tab') { e.preventDefault(); finish(e.shiftKey ? [0, -1, 'tab'] : [0, 1, 'tab']); return; }
    if (open.mode === 'enter' && MOVES[e.key]) { e.preventDefault(); finish(MOVES[e.key]); }
  });
  box.addEventListener('beforeinput', () => { if (open) rich.track.before(); });
  box.addEventListener('input', (e) => { if (open) sync(e); });
  // (The drawn caret and selection follow the box's.)
  for (const ev of ['keydown', 'keyup', 'select', 'selectionchange', 'focus']) box.addEventListener(ev, placeMarks);
  box.addEventListener('blur', (e) => {
    if (held) return;
    // (To the Font group's font or size list with text selected, as Excel's in-cell editor allows: kept open until
    // that list lets focus go somewhere other than back here.)
    if (open && keepsEdit(e.relatedTarget) && box.selectionEnd > box.selectionStart) {
      parked = true;
      e.relatedTarget.addEventListener('focusout', (ev) => {
        if (!parked) return;
        parked = false;
        if (ev.relatedTarget !== box) finish(null);
      }, { once: true });
      return;
    }
    finish(null);
  });
  box.addEventListener('mousedown', (e) => pointAt(e));

  /** A click or a drag in drawn text: the caret put by the drawn characters, not the box's hidden ones (a double
   *  click selects the word). */
  function pointAt(e) {
    if (!open || !rich.view.el.isConnected || e.button !== 0) return;
    e.preventDefault();
    box.focus({ preventScroll: true });
    const anchor = e.shiftKey ? box.selectionStart : rich.view.offsetAt(e.clientX, e.clientY);
    if (e.detail === 2) {
      const t = box.value;
      let [a, b] = [anchor, anchor];
      while (a > 0 && /[\p{L}\p{N}_]/u.test(t[a - 1])) a--;
      while (b < t.length && /[\p{L}\p{N}_]/u.test(t[b])) b++;
      box.setSelectionRange(a, b);
      rich.view.place();
      return;
    }
    const to = (ev) => {
      const at = rich.view.offsetAt(ev.clientX, ev.clientY);
      box.setSelectionRange(Math.min(anchor, at), Math.max(anchor, at), at < anchor ? 'backward' : 'forward');
      rich.view.place();
    };
    to(e);
    const up = () => { window.removeEventListener('mousemove', to); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', to);
    window.addEventListener('mouseup', up);
  }

  function close() {
    if (!open) return;
    open = null;
    box.remove();
    rich.view.el.remove();
    box.classList.remove('oxe-editor-rich');
    onClose?.();
  }

  return {
    /** Open over cell (r, c) with `text`, in `mode` ('enter' replaces the cell, 'edit' keeps it), in class `cls` (the
     *  cell format's, for its font). */
    open(r, c, text, mode, cls, runs = null) {
      close();
      // (A cell of rich text keeps its formatting while edited, one font per character.)
      rich.track.reset(text, runs);
      const rect = grid.cellRect(r, c);
      open = { r, c, mode, x: rect.x, y: rect.y, w: rect.w, h: rect.h, cls: cls || '' };
      box.className = `oxe-editor ${cls || ''}`;
      box.value = text;
      corrector.refresh();
      grid.overlay.appendChild(box);
      box.focus({ preventScroll: true });
      box.setSelectionRange(text.length, text.length);
      show();
    },
    close,
    /** Enter what is typed, without moving (another sheet is about to be shown). */
    commit: () => finish(null),
    get isOpen() { return !!open; },
    mark,
    /** Font properties ({ name, size, color }) set on the selection's part of the text (richmarks.js setFont);
     *  false when nothing is selected. `size` may be a function of the selection's first size (Grow Font). */
    markFont: (patch, base = null) => change((fonts, a, b) => setFont(fonts, a, b, typeof patch === 'function' ? patch(sizeAt(fonts, a, base)) : patch)),
    /** Whether part of the text is selected (the Font group then formats it). */
    get selecting() { return !!open && box.selectionEnd > box.selectionStart && !box.value.startsWith('='); },
    /** The font of the selection's first character (richmarks.js: null, the cell's own), for Format Cells. */
    get selectedFont() { return open ? rich.track.fonts[box.selectionStart] ?? null : null; },
    /** Kept open while a box is over it (`on`), or no longer (`restore`: the selection put back as it was). */
    hold(on, restore = false) {
      if (!on && restore && held && open) box.setSelectionRange(held[0], held[1]);
      held = on && open ? [box.selectionStart, box.selectionEnd] : null;
    },
    /** `text` typed at the caret (in place of what is selected), as edit mode. */
    insert(text) {
      if (!open) return;
      const [start, end] = held || [box.selectionStart, box.selectionEnd];
      box.focus({ preventScroll: true });
      box.setRangeText(text, start, end, 'end');
      open.mode = 'edit';
      size();
      box.dispatchEvent(new Event('input'));
    },
    /** 'enter' or 'edit' (null when closed), and the cell being edited. */
    get mode() { return open ? open.mode : null; },
    get at() { return open ? { r: open.r, c: open.c } : null; },
    el: box,
  };
}
