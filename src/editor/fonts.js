// A workbook's typeface as a CSS font stack: the font itself, then a metric-compatible stand-in where a free one
// exists (same advance widths and vertical metrics, so text measures the same where the original isn't installed),
// then a generic family. And each font's line height for celltext.js.
import metrics from '../engine/metrics.js';

/** Free fonts drawn to the same metrics as Office's. Aptos Narrow (the body font of Excel 365's default theme since
 *  2023) has no metric-compatible twin and isn't installed with Windows (Office keeps it in its own cloud-font cache,
 *  out of browsers' reach); Calibri is the closest stand-in measured: as wide within a pixel in Excel's PDFs of
 *  test/fixtures/basic.xlsx ("Amount" 47.7px in Aptos Narrow, 48 in Calibri), with its baseline within half a pixel. */
const STAND_INS = {
  'aptos narrow': ['Calibri', 'Carlito'],
  calibri: ['Carlito'],
  cambria: ['Caladea'],
  arial: ['Liberation Sans', 'Arimo'],
  'times new roman': ['Liberation Serif', 'Tinos'],
  'courier new': ['Liberation Mono', 'Cousine'],
};
const SERIF = /^(cambria|times|georgia|garamond|book antiqua|palatino|constantia|aptos serif|century)/i;
const MONO = /^(courier|consolas|lucida console|aptos mono|cascadia)/i;

const quote = (name) => (/^[\w-]+$/.test(name) ? name : `'${name.replace(/'/g, '')}'`);

/** The CSS font-family for a typeface name. */
export function fontStack(name) {
  const generic = SERIF.test(name) ? 'serif' : MONO.test(name) ? 'monospace' : 'sans-serif';
  return [name, ...(STAND_INS[name.toLowerCase()] || []), generic].map((n) => (n === generic ? n : quote(n))).join(', ');
}

/** Where Excel's baseline sits above the cell's bottom edge, beyond the descent of the font the browser draws:
 *  0.15px for a font drawn as itself (test/fixtures/fonts.xlsx: Calibri 11 to 20, Cambria, Arial, Times New Roman;
 *  small sizes sit higher, Calibri 8 by 1.1px, Courier New lower), and for Aptos Narrow, drawn in Calibri, 0.85px:
 *  Excel's Aptos Narrow baseline is that much higher than Calibri's descent (large.xlsx, 100,000 cells, rows
 *  validated over 43 steps; runs 8 -> 9: baselines within 1px 81.8% -> 100.0%). Run 13 showed the 0.85 is Aptos
 *  Narrow's, not the browser's: applied to Calibri workbooks it put fonts.xlsx's baselines 0.7px high. */
const BASELINE_ABOVE_DESCENT = 0.15;
const BASELINE_STAND_IN = { 'aptos narrow': 0.85 };

/** Top-aligned text's baseline is the font's Windows ascent plus this far below the row's top (Excel's PDF of
 *  test/fixtures/align.xlsx, Calibri 11 in a 45pt row: 14.76px = 13.96 + 0.8). */
const TOP_INSET = 0.8;

/** The font's metrics in px at `sizePx`: Chromium's (hhea ascent and descent, by which it places text) and Excel's
 *  (Windows ascent and descent), from engine/metrics.js (the installed fonts); the stand-in's when the font isn't in
 *  the table (it is the stand-in that the browser draws), Calibri's when neither is. */
function fontMetrics(name, sizePx) {
  const key = name.toLowerCase();
  const own = metrics[key];
  const m = own || (STAND_INS[key] || []).map((s) => metrics[s.toLowerCase()]).find(Boolean) || metrics.calibri;
  return {
    ha: m.hhea[0] * sizePx, hd: m.hhea[1] * sizePx, wa: m.win[0] * sizePx, wd: m.win[1] * sizePx,
    base: own ? BASELINE_ABOVE_DESCENT : BASELINE_STAND_IN[key] ?? BASELINE_ABOVE_DESCENT,
  };
}

/** The line height (px) that puts the browser's baseline where Excel's is for bottom-aligned text. Chromium centres
 *  the hhea ascent and descent in the line (half the leading above, half below), so a line of L px has its baseline
 *  (L - (ascent + descent)) / 2 + descent above its bottom; solved for Excel's baseline. */
export function cellLineHeight(name, sizePx) {
  const f = fontMetrics(name, sizePx);
  return f.ha + f.hd + 2 * (f.wd + f.base - f.hd);
}

/** Everything that places a font's text in a cell, in px: the single line's height (`lh`); how far to move the text
 *  down from where the browser puts it, for top- and middle-aligned single lines (`top`, `mid`); and for wrapped text,
 *  the line pitch and the shifts for bottom, top and middle alignment, on screen and in print (`pitch` gives the
 *  pitch, the font's default row height, in each). Excel places a top-aligned line's baseline TOP_INSET + ascent below
 *  the top, centres a line's ascent + descent in the cell, and steps wrapped lines by the font's default row height
 *  (align.xlsx: Calibri 11 wraps every 19.36px in print, its 15pt row). */
export function cellTextMetrics(name, sizePx, pitch) {
  const f = fontMetrics(name, sizePx);
  const lh = cellLineHeight(name, sizePx);
  const chromeTop = (L) => (L - f.ha - f.hd) / 2 + f.ha; // baseline below the top of a line of L
  const chromeBottom = (L) => (L - f.ha - f.hd) / 2 + f.hd; // baseline above the bottom
  const mid = (f.wa - f.wd) / 2 - (f.ha - f.hd) / 2; // (a centred line box puts the baseline (ha - hd) / 2 below centre)
  const wrap = (P) => ({ lh: P, bottom: chromeBottom(P) - (f.wd + f.base), top: TOP_INSET + f.wa - chromeTop(P), mid });
  return { lh, top: TOP_INSET + f.wa - chromeTop(lh), mid, wrap: wrap(pitch.screen), printWrap: wrap(pitch.print) };
}

/** The font's Windows ascent + descent in px (the height Excel sizes rows by). */
export const fontMetricsWinLine = (name, sizePx) => {
  const f = fontMetrics(name, sizePx);
  return f.wa + f.wd;
};
