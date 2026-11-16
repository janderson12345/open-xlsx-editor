// The browser's measure of a typeface's widest digit, for ooxml/digitwidth.js: a canvas set to the font at the size
// in CSS pixels, each of 0-9 measured (the font's own advance widths, unhinted, as Chromium lays text out). A font the
// page doesn't have is measured in the one the browser falls back to.
import { setDigitMeasure } from '../ooxml/digitwidth.js';

let ctx = null;
/** The widest of 0-9 in `family` at `px` CSS pixels (fractional), or null without a canvas. */
export function measureDigits(family, px) {
  if (typeof document === 'undefined') return null;
  if (!ctx) ctx = document.createElement('canvas').getContext('2d');
  if (!ctx) return null;
  ctx.font = `${px}px "${String(family).replace(/"/g, '')}", sans-serif`;
  let widest = 0;
  for (let d = 0; d <= 9; d++) widest = Math.max(widest, ctx.measureText(String(d)).width);
  return widest;
}

/** Make the browser's measure the one digitwidth.js asks (once, when an editor is made). */
export function useBrowserDigits() {
  if (typeof document !== 'undefined') setDigitMeasure(measureDigits);
}
