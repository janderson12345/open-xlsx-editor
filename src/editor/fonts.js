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

/** The line height (px) that puts the browser's baseline where Excel's is, for a font at `sizePx`. Chromium places
 *  text by the font's hhea ascent and descent and centres them in the line (half the leading above, half below), so
 *  a line of L px has its baseline (L - (ascent + descent)) / 2 + descent above its bottom; solved for Excel's
 *  baseline. Metrics from engine/metrics.js (the installed fonts); the stand-in's when the font isn't in the table
 *  (it is the stand-in that the browser draws), Calibri's when neither is. */
export function cellLineHeight(name, sizePx) {
  const key = name.toLowerCase();
  const m = metrics[key] || (STAND_INS[key] || []).map((s) => metrics[s.toLowerCase()]).find(Boolean) || metrics.calibri;
  const [ascent, descent] = m.hhea;
  const target = m.win[1] * sizePx + (metrics[key] ? BASELINE_ABOVE_DESCENT : BASELINE_STAND_IN[key] ?? BASELINE_ABOVE_DESCENT);
  return (ascent + descent) * sizePx + 2 * (target - descent * sizePx);
}
