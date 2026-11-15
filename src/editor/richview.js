// The cell editor's text drawn as its runs while part of it is formatted (richmarks.js), as Excel's in-cell editor
// shows bold, italic, underline, strikethrough, colour, font, size and super- and subscript as they are typed. The
// text box stays the one typed in (its caret, selection, IME, undo and keys are the browser's own); its text is made
// invisible and this view under it draws each run in its font (cellstyles.js runCss, as the cell draws it) with the
// caret and selection placed by the drawn text. A click in the box is placed by the drawn text too, since the box's
// own characters are narrower or wider than the drawn ones.
import { setVars } from './dom.js';

/** A view of `box` (the cell editor's textarea): `draw(text, fonts)`, `place()` the caret and selection, `offsetAt()`
 *  a point. `runStyle(font)` gives a run's CSS. */
export function createRichView(box, runStyle) {
  const view = document.createElement('div');
  view.className = 'oxe-editor oxe-edit-view';
  view.setAttribute('aria-hidden', 'true');
  const marks = document.createElement('div');
  marks.className = 'oxe-edit-marks';
  const text = document.createElement('div');
  text.className = 'oxe-edit-runs';
  view.append(marks, text);
  let pieces = []; // [{ node, start }]: each run's text node and where it starts in the text
  let length = 0;

  const piece = (i) => pieces.find((q) => i >= q.start && i < q.start + q.node.length);
  /** The text node and offset in it of position `at` (at a run's edge, the end of the run before). */
  const posAt = (at) => {
    const p = at > 0 ? piece(at - 1) : pieces[0];
    return p ? [p.node, at - p.start] : [text, 0];
  };
  const range = document.createRange();
  /** The client rect of character `i`. */
  const charRect = (i) => {
    const p = piece(i);
    if (!p) return new DOMRect();
    range.setStart(p.node, i - p.start);
    range.setEnd(p.node, i - p.start + 1);
    return range.getBoundingClientRect();
  };
  /** Client coordinates to the view's own (unzoomed CSS pixels inside its border). */
  const local = () => {
    const r = view.getBoundingClientRect();
    const k = r.width ? view.offsetWidth / r.width : 1;
    return (x, y) => [(x - r.left) * k - view.clientLeft, (y - r.top) * k - view.clientTop];
  };
  const box1 = (cls, x, y, w, h) => {
    const el = document.createElement('div');
    el.className = cls;
    setVars(el, { '--oxe-mark-x': x, '--oxe-mark-y': y, '--oxe-mark-w': w, '--oxe-mark-h': h });
    return el;
  };

  return {
    el: view,
    /** The runs of `str` by `fonts` (one font per character, null: the cell's own). */
    draw(str, fonts) {
      text.replaceChildren();
      pieces = [];
      length = str.length;
      let i = 0;
      while (i < str.length) {
        let j = i + 1;
        while (j < str.length && fonts[j] === fonts[i]) j++;
        const span = document.createElement('span');
        const node = document.createTextNode(str.slice(i, j));
        span.appendChild(node);
        if (fonts[i]) span.style.cssText = runStyle(fonts[i]);
        text.appendChild(span);
        pieces.push({ node, start: i });
        i = j;
      }
    },
    /** The caret, or the selection, where the box's are, drawn by the text. */
    place() {
      const [s, e] = [box.selectionStart, box.selectionEnd];
      const at = local();
      const out = [];
      if (e > s) {
        const [n1, o1] = posAt(s);
        const [n2, o2] = posAt(e);
        range.setStart(n1, o1);
        range.setEnd(n2, o2);
        for (const r of range.getClientRects()) {
          const [x, y] = at(r.left, r.top);
          const [x2, y2] = at(r.right, r.bottom);
          if (x2 > x) out.push(box1('oxe-edit-sel', x, y, x2 - x, y2 - y));
        }
      } else if (document.activeElement === box) {
        out.push(this.caretBox(s, at));
      }
      marks.replaceChildren(...out);
    },
    /** The caret's box at position `s`: after the character before it, or before the first of a line. */
    caretBox(s, at = local()) {
      let x = 0;
      let y = 0;
      let h = text.offsetHeight || view.clientHeight;
      const before = s > 0 ? charRect(s - 1) : null;
      if (before && before.height && box.value[s - 1] !== '\n') {
        [x, y] = at(before.right, before.top);
        h = at(before.right, before.bottom)[1] - y;
      } else if (s < length && box.value[s] !== '\n') {
        const after = charRect(s);
        [x, y] = at(after.left, after.top);
        h = at(after.left, after.bottom)[1] - y;
      } else if (before) {
        // (After a line break with nothing after it: the start of the next line.)
        const [, by] = at(before.left, before.bottom);
        [x, y] = [parseFloat(getComputedStyle(view).paddingLeft) || 0, by];
        h = at(before.left, before.bottom)[1] - at(before.left, before.top)[1];
      } else x = parseFloat(getComputedStyle(view).paddingLeft) || 0;
      return box1('oxe-edit-caret', x, y, 1, h);
    },
    /** The position in the text nearest the client point (x, y), by the drawn characters. */
    offsetAt(cx, cy) {
      let best = { d: Infinity, at: length };
      for (let i = 0; i < length; i++) {
        const r = charRect(i);
        if (!r.height) continue;
        const dy = cy < r.top ? r.top - cy : cy > r.bottom ? cy - r.bottom : 0;
        const dx = cx < r.left ? r.left - cx : cx > r.right ? cx - r.right : 0;
        const d = dy * 1e4 + dx;
        // (A line break is the end of its line: a click past the line's text goes before it.)
        if (d < best.d) best = { d, at: i + (box.value[i] !== '\n' && cx > (r.left + r.right) / 2 ? 1 : 0) };
      }
      return best.at;
    },
    /** The drawn text's size (as the box's scrollWidth and scrollHeight are). */
    get scrollWidth() { return view.scrollWidth; },
    get scrollHeight() { return view.scrollHeight; },
  };
}
