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

/** Where Excel's baseline sits above the cell's bottom edge, beyond the font's Windows descent: 0.9px. Read off
 *  Excel's PDF of test/fixtures/rows.xlsx, where fills give every row's bottom exactly: Calibri 11, 14 and 20, Arial 14,
 *  Times New Roman 16 and Cambria 20 all within about 0.3px of it (Calibri 36 sits 0.6px higher, Courier New lower);
 *  large.xlsx, drawn in Aptos Narrow, fitted 0.85 independently. An earlier 0.15 came from fonts.xlsx's rows, whose
 *  positions carried the drift of an early row-height fit. */
const BASELINE_ABOVE_DESCENT = 0.9;

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
    base: BASELINE_ABOVE_DESCENT,
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

/** A font's Windows descent in px (where Excel's baseline sits above a row's bottom, plus BASELINE_ABOVE_DESCENT). */
export const fontDescentPx = (name, sizePx) => fontMetrics(name, sizePx).wd;

/** A font's row height in px, the height Excel gives a row of it (and the pitch of its wrapped lines): its line,
 *  the larger of Windows ascent + descent and hhea ascent + descent + gap, plus 1.6px and 0.035 of its size, rounded.
 *  Fitted on Excel's PDF of test/fixtures/rows.xlsx: Calibri 11, 14, 20, 36 (15, 18.75, 26.25, 46.5pt), Cambria 20
 *  (25.5), Times New Roman 16 (20.25), Arial 14 (18), all exact; a fit, not Excel's own formula. */
export function fontRowPx(name, sizePx) {
  const key = name.toLowerCase();
  const m = metrics[key] || (STAND_INS[key] || []).map((s) => metrics[s.toLowerCase()]).find(Boolean) || metrics.calibri;
  return Math.round(Math.max(m.win[0] + m.win[1], m.line) * sizePx + 1.6 + 0.035 * sizePx);
}
