// The grid's mouse while Draw Borders is on (drawborders.js): a drag outlines the cells it will draw on (as the fill
// handle's outline) and, let go, draws there through ctx.drawn; a click within 4px of a gridline draws (or erases)
// that one line, elsewhere the cell's own. Nothing is selected meanwhile. Esc ends the mode. Attached first, so these
// come before the grid's selecting, fill handle and heading edges (interact.js).

const NEAR = 4;
const range = ({ from, to }) => ({ r0: Math.min(from.r, to.r), r1: Math.max(from.r, to.r), c0: Math.min(from.c, to.c), c1: Math.max(from.c, to.c) });

/** `ctx`: drawing() the mode or null, drawn(range, side), stopDrawing(). Returns a function that detaches it. */
export function attachBorderDrag(view, grid, ctx) {
  let drag = null;
  // (The nearest side within NEAR pixels where the cell under the pointer ends.)
  const sideNear = (x, y, at) => {
    for (let d = 1; d <= NEAR; d++) {
      for (const [side, dx, dy] of [['left', -d, 0], ['right', d, 0], ['top', 0, -d], ['bottom', 0, d]]) {
        const h = grid.hitTest(x + dx, y + dy);
        if (h && (h.r !== at.r || h.c !== at.c)) return side;
      }
    }
    return null;
  };
  const down = (e) => {
    if (e.button !== 0 || !ctx.drawing?.() || e.target.closest?.('.oxe-editor')) return;
    const at = grid.hitTest(e.clientX, e.clientY);
    if (!at || at.r < 0 || at.c < 0) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    view.focus({ preventScroll: true });
    drag = { from: at, to: at, side: sideNear(e.clientX, e.clientY, at) };
  };
  const move = (e) => {
    if (!drag) return;
    const at = grid.hitTest(e.clientX, e.clientY);
    if (!at || at.r < 0 || at.c < 0 || (at.r === drag.to.r && at.c === drag.to.c)) return;
    drag.to = at;
    drag.side = null;
    grid.showFill(range(drag));
  };
  const up = () => {
    if (!drag) return;
    const done = drag;
    drag = null;
    grid.showFill(null);
    ctx.drawn?.(range(done), done.side);
  };
  const key = (e) => {
    if (e.key !== 'Escape' || e.target !== view || !ctx.drawing?.()) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    ctx.stopDrawing?.();
  };
  view.addEventListener('mousedown', down, true);
  view.addEventListener('keydown', key, true);
  window.addEventListener('mousemove', move);
  window.addEventListener('mouseup', up);
  return () => {
    view.removeEventListener('mousedown', down, true);
    view.removeEventListener('keydown', key, true);
    window.removeEventListener('mousemove', move);
    window.removeEventListener('mouseup', up);
  };
}
