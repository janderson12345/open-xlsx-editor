// Text measured as Excel measures it to break lines (Fill > Justify), not as the browser draws it: each character's
// advance at the font's size in whole pixels (11pt, 14.67px, at 15), rounded to a whole pixel, without kerning, as
// Windows' GDI measures. Fitted on Excel's Range.Justify of 36 paragraphs (Aptos Narrow, Calibri and Arial, 64 to
// 215px wide): 32 broken exactly as Excel broke them, with a line kept while it is at most the width less 4px. The
// browser's own widths times a factor did no better than 28 (docs/FILL.md). Excel's advances are its fonts' hinted
// ones, which only follow this roughly (Aptos Narrow's digits are 7px at 15, not 8).

let ctx = null;
const cache = new Map();

/** The width (px) Excel gives `text` in the font `css` (a canvas font, "bold 14.6667px Calibri, sans-serif") at
 *  `sizePx`. */
export function excelTextWidth(text, css, sizePx) {
  if (!ctx) ctx = document.createElement('canvas').getContext('2d');
  const font = css.replace(/[\d.]+px/, `${Math.round(sizePx)}px`);
  let widths = cache.get(font);
  if (!widths) { widths = new Map(); cache.set(font, widths); }
  ctx.font = font;
  let w = 0;
  for (const ch of text) {
    let a = widths.get(ch);
    if (a == null) { a = Math.round(ctx.measureText(ch).width); widths.set(ch, a); }
    w += a;
  }
  return w;
}

/** How much narrower than the selection Excel keeps a justified line. */
export const JUSTIFY_MARGIN = 4;
