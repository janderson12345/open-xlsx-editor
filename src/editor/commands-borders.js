// Home > Font's Draw Borders and More Colors (commands.js shares its state with them; docs/FORMATS.md): a range
// dragged over with Draw Border, Draw Border Grid or Erase Border, and a colour of any RGB value given to the
// selection's text or fill, added to the colour menus' Recent Colors as Excel keeps them.
import { drawLines, DRAW_MODES } from '../ooxml/formatops.js';
import { addRecentColor } from '../ooxml/recentcolors.js';

/** A colour as written: 'RRGGBB', '#RRGGBB' or 'FFRRGGBB' to 'FFRRGGBB' (Excel writes a custom colour opaque), or
 *  null when it isn't one. */
export function rgbOf(value) {
  const hex = String(value ?? '').trim().replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex)) return `FF${hex}`;
  if (/^FF[0-9A-F]{6}$/.test(hex)) return hex;
  return null;
}

/** These commands: `env` is createCommands's shared state and helpers (commands.js). */
export function borderCommands(env) {
  const { commands, ctx, restyle } = env;
  return {
    /** Draw Borders over `range` ({ r0, c0, r1, c1 }) on the sheet shown, as one step of the history: `mode` 'draw'
     *  (its outline), 'grid' (every line) or 'erase'; `side` one edge of a one-cell range (a gridline clicked); `line`
     *  { style, color } (formatops.js drawLines). */
    drawBorders(mode, range, { side = null, line } = {}) {
      if (!DRAW_MODES.includes(mode) || !range) return false;
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const r0 = Math.min(range.r0, range.r1 ?? range.r0);
      const c0 = Math.min(range.c0, range.c1 ?? range.c0);
      const area = { r0, c0, r1: Math.max(range.r0, range.r1 ?? range.r0), c1: Math.max(range.c0, range.c1 ?? range.c0) };
      if (side && (area.r0 !== area.r1 || area.c0 !== area.c1)) return false;
      return restyle(drawLines(ctx.workbook(), ctx.sheetIndex(), area, mode, { side, ...(line ? { line } : {}) }));
    },
    /** A colour put first in the workbook's Recent Colors (recentcolors.js); not part of the undo history, as the
     *  format it came with is. */
    recentColor(value) {
      const rgb = rgbOf(value);
      if (!rgb || ctx.readOnly()) return false;
      if (addRecentColor(ctx.workbook(), rgb)) ctx.edited?.({ book: true });
      return true;
    },
    /** More Colors: `value` (an RGB colour, rgbOf) given to the selection's text ('font') or fill ('fill') as one step
     *  of the history, written `rgb="FFRRGGBB"` as Excel writes a custom colour, and put in Recent Colors. */
    moreColor(target, value) {
      const rgb = rgbOf(value);
      if (!rgb || (target !== 'font' && target !== 'fill')) return false;
      const color = { rgb };
      const done = commands.format(target === 'font' ? { font: { color } } : { fill: color });
      if (done) commands.recentColor(rgb);
      return done;
    },
  };
}
