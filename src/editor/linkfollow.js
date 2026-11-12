// Following a hyperlink as Excel: a click on its cell follows it (a press held longer, or moved off the cell, selects
// it instead), with the hand over it and Excel's ScreenTip (ooxml/links.js).
import { hyperlinkAt, screenTip } from '../ooxml/links.js';

/** How long a press may last and still follow the link (Excel's "Click and hold to select this cell"). */
const HOLD_MS = 500;

/** The grid's hyperlinks under the mouse: `press` on a mousedown on cell `at`, `release` with the cell the mouse came
 *  up on, `hover` as it moves (null when no cell). ctx: workbook(), sheetIndex(), followLink(h). */
export function linkFollower(view, ctx) {
  let pressed = null;
  let over = null;
  const linkAt = (at) => (at && at.r >= 0 && at.c >= 0 ? hyperlinkAt(ctx.workbook(), ctx.sheetIndex(), at.r, at.c) : null);
  return {
    press(at, e) {
      const plain = e.button === 0 && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey && e.detail === 1;
      pressed = plain && linkAt(at) ? { r: at.r, c: at.c, t: performance.now() } : null;
    },
    release(at) {
      const p = pressed;
      pressed = null;
      if (!p || !at || at.r !== p.r || at.c !== p.c || performance.now() - p.t > HOLD_MS) return;
      const h = linkAt(at);
      if (h) ctx.followLink(h);
    },
    hover(at) {
      const key = at && at.r >= 0 && at.c >= 0 ? `${at.r},${at.c}` : null;
      if (key === over) return;
      over = key;
      const h = linkAt(at);
      view.classList.toggle('oxe-link-cursor', !!h);
      if (h) view.title = screenTip(h);
      else view.removeAttribute('title');
    },
  };
}
