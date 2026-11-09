// Editing a cell in place: a text box over the active cell in the cell's own font, which grows with what is typed.
// Two modes, as in Excel: typing over a cell (enter mode, where the arrow keys finish and move) and F2 or a double
// click (edit mode, where they move the caret); F2 switches between them, as in Excel. Enter, Tab and their Shift forms
// finish and move; Esc gives up; Alt+Enter is a line break. `onKey` sees each key first (pointing.js takes the arrows
// while a formula is pointing).
import { setVars } from './grid.js';

const MOVES = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/** `grid` (its overlay and cell boxes), `onFinish(text, [dr, dc])`, `onCancel()`, `onKey(e)` (true: taken) and
 *  `onClose()`. */
export function createCellEditor({ grid, onFinish, onCancel, onKey = null, onClose = null }) {
  const box = document.createElement('textarea');
  box.className = 'oxe-editor';
  box.spellcheck = false;
  box.setAttribute('aria-label', 'Cell contents');
  let open = null; // { r, c, mode }

  const size = () => {
    // (As wide as the text needs, never narrower than the cell; as tall as its lines.)
    box.style.setProperty('--oxe-edit-w', '0px');
    setVars(box, { '--oxe-edit-w': Math.max(open.w, box.scrollWidth + 4), '--oxe-edit-h': Math.max(open.h, box.scrollHeight) });
  };
  const finish = (move) => {
    if (!open) return;
    const text = box.value;
    close();
    onFinish(text, move);
  };

  box.addEventListener('keydown', (e) => {
    if (!open) return;
    if (onKey?.(e)) { e.preventDefault(); return; }
    if (e.key === 'F2') { e.preventDefault(); open.mode = open.mode === 'enter' ? 'edit' : 'enter'; return; }
    if (e.key === 'Escape') { e.preventDefault(); close(); onCancel(); return; }
    if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      box.setRangeText('\n', box.selectionStart, box.selectionEnd, 'end');
      size();
      return;
    }
    if (e.key === 'Enter') { e.preventDefault(); finish(e.shiftKey ? [-1, 0] : [1, 0]); return; }
    if (e.key === 'Tab') { e.preventDefault(); finish(e.shiftKey ? [0, -1] : [0, 1]); return; }
    if (open.mode === 'enter' && MOVES[e.key]) { e.preventDefault(); finish(MOVES[e.key]); }
  });
  box.addEventListener('input', () => open && size());
  box.addEventListener('blur', () => finish(null));

  function close() {
    if (!open) return;
    open = null;
    box.remove();
    onClose?.();
  }

  return {
    /** Open over cell (r, c) with `text`, in `mode` ('enter' replaces the cell, 'edit' keeps it), in class `cls` (the
     *  cell format's, for its font). */
    open(r, c, text, mode, cls) {
      close();
      const rect = grid.cellRect(r, c);
      open = { r, c, mode, w: rect.w, h: rect.h };
      box.className = `oxe-editor ${cls || ''}`;
      box.value = text;
      setVars(box, { '--oxe-x': rect.x, '--oxe-y': rect.y });
      grid.overlay.appendChild(box);
      size();
      box.focus({ preventScroll: true });
      box.setSelectionRange(text.length, text.length);
    },
    close,
    /** Enter what is typed, without moving (another sheet is about to be shown). */
    commit: () => finish(null),
    get isOpen() { return !!open; },
    /** 'enter' or 'edit' (null when closed), and the cell being edited. */
    get mode() { return open ? open.mode : null; },
    get at() { return open ? { r: open.r, c: open.c } : null; },
    el: box,
  };
}
