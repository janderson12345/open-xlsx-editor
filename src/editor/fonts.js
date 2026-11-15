// A workbook's typeface as a CSS font stack: the font itself, then a metric-compatible stand-in where a free one
// exists (same advance widths and vertical metrics, so text measures the same where the original isn't installed),
// then a generic family. And each font's line height for celltext.js.
import metrics from '../engine/metrics.js';
import { omfEntry, omfFamily } from './omffonts.js';

/** Free fonts drawn to the same metrics as Office's. Aptos Narrow (the body font of Excel 365's default theme since
 *  2023) isn't installed with Windows (Office keeps it in its own cloud-font cache, out of browsers' reach); its
 *  metric-compatible twin is Open Metric Fonts' Morrow Sans Narrow, which a page can serve itself (README: Fonts;
 *  omffonts.js puts it in the stack, with the other families' twins) or a computer has installed, and Aptos's is
 *  Morrow Sans. Without them, Calibri is the closest measured: as wide within a pixel in Excel's PDFs
 *  of test/fixtures/basic.xlsx ("Amount" 47.7px in Aptos Narrow, 48 in Calibri), its baseline within half a pixel. */
const STAND_INS = {
  'aptos narrow': ['Morrow Sans Narrow', 'Calibri', 'Carlito'],
  aptos: ['Morrow Sans'],
  calibri: ['Carlito'],
  cambria: ['Caladea'],
  arial: ['Liberation Sans', 'Arimo'],
  'times new roman': ['Liberation Serif', 'Tinos'],
  'courier new': ['Liberation Mono', 'Cousine'],
};
const SERIF = /^(cambria|times|georgia|garamond|book antiqua|palatino|constantia|aptos serif|century)/i;
const MONO = /^(courier|consolas|lucida console|aptos mono|cascadia)/i;

const quote = (name) => (/^[\w-]+$/.test(name) ? name : `'${name.replace(/'/g, '')}'`);

/** The CSS font-family for a typeface name: the font, its open-metric-fonts twin where the page serves them
 *  (omffonts.js: drawn only where the font itself is missing), the stand-ins, a generic family. */
export function fontStack(name) {
  const generic = SERIF.test(name) ? 'serif' : MONO.test(name) ? 'monospace' : 'sans-serif';
  const twin = omfFamily(name);
  return [name, ...(twin ? [twin] : []), ...(STAND_INS[name.toLowerCase()] || []), generic].map((n) => (n === generic ? n : quote(n))).join(', ');
}

/** Where Excel's baseline sits above the cell's bottom edge, beyond the Windows descent of the font the browser draws,
 *  in this grid's coordinates. Excel's own is about 0.9px above its printed row boundary (rows.xlsx, whose fills give
 *  every boundary exactly: Calibri 11, 14, 20, Arial 14, Times New Roman 16, Cambria 20 within about 0.3px), but
 *  Excel's printed grid starts 0.7px lower than this one (decor.js PRINT_GRID_ORIGIN), and text is placed in this
 *  grid: 0.9 - 0.7, fitted at 0.15 by centring numfmt.xlsx's and align.xlsx's 1615 baselines on Excel's (Chromium
 *  snaps baselines to whole pixels, so the best is a band of a pixel centred on Excel's). Aptos Narrow, drawn in
 *  Calibri, sits higher in Excel: 0.3, centring large.xlsx's 100,000. */
const BASELINE_ABOVE_DESCENT = 0.15;
const BASELINE_STAND_IN = { 'aptos narrow': 0.3 };

/** Top-aligned text's baseline is the font's Windows ascent plus this far below the row's top (Excel's PDF of
 *  test/fixtures/align.xlsx, Calibri 11 in a 45pt row: 14.76px = 13.96 + 0.8, from Excel's printed row boundary). */
const TOP_INSET = 0.8;
/** Excel's printed grid starts this much lower than this one (decor.js PRINT_GRID_ORIGIN): targets read from Excel's
 *  boundaries move down by it here. */
const GRID_ORIGIN_Y = 0.7;

/** What each family's check found (hasFont): a web font can arrive after the first look. */
const found = new Map();
const probe = (family) => {
  const ctx = document.createElement('canvas').getContext('2d');
  const w = (f) => { ctx.font = `40px ${f}`; return ctx.measureText('mmmmmmmmmmlli1WQ').width; };
  const has = w(`'${family}', monospace`) !== w('monospace') || w(`'${family}', serif`) !== w('serif');
  const name = family.toLowerCase();
  const web = has && [...document.fonts].some((f) => f.status === 'loaded' && f.family.replace(/["']/g, '').toLowerCase() === name);
  return has ? (web ? 'web' : 'installed') : '';
};
/** How the browser has `family` to draw: 'installed', 'web' (an @font-face already loaded) or '' (not at all): canvas
 *  text in it measures differently from both generic fallbacks. */
function hasFont(family) {
  if (typeof document === 'undefined' || !document.fonts) return '';
  if (!found.has(family)) found.set(family, probe(family));
  return found.get(family);
}

/** Checks again the families looked at (after web fonts load): whether any is now drawn otherwise, so that text must
 *  be placed and measured again. */
export function fontsChanged() {
  let changed = false;
  for (const [family, was] of found) {
    const now = probe(family);
    if (now !== was) { found.set(family, now); changed = true; }
  }
  return changed;
}
/** The metrics of the open-metric-fonts twin the browser draws in place of `name` (the font itself missing), served
 *  (omffonts.js) or installed (a Morrow font of STAND_INS); null when it draws none. Read from the twins' files
 *  (omf.js, numbers only): their Windows ascent and descent are the Office font's own, where Excel puts the text; the
 *  browser places them by their own, the typo metrics where they set USE_TYPO_METRICS (both Morrow fonts: with Morrow
 *  Sans Narrow installed, placed by the Windows metrics, large.xlsx had 61,460 of 100,000 baselines within 1px). Measured
 *  with Morrow Sans Narrow served, Edge, Excel's PDFs of basic.xlsx and large.xlsx (Aptos Narrow): 100,049 of 100,049
 *  cells within 1px across, down and in width (Calibri drawn: the same across and down, 99,028 in width); placed by the
 *  Windows metrics instead, 0.95px high. */
function twinMetrics(name) {
  const entry = omfEntry(name);
  if (!entry || hasFont(name)) return null;
  const served = omfFamily(name);
  const installed = (STAND_INS[name.toLowerCase()] || []).find((s) => /^morrow /i.test(s));
  return (served && hasFont(served)) || (installed && hasFont(installed)) ? entry[1] : null;
}

/** The font's metrics in px at `sizePx`: the twin's when the browser draws one (twinMetrics), else from
 *  engine/metrics.js (the installed fonts), the stand-in's when the font isn't in the table (it is the stand-in that
 *  the browser draws), Calibri's when neither is: Windows ascent and descent (`wa`, `wd`: where Excel places text) and
 *  the same rounded to whole pixels (`ar`, `dr`: where Chromium does). */
function fontMetrics(name, sizePx) {
  const key = name.toLowerCase();
  // (A metric twin drawn: the Office font's own metrics for where Excel puts the text, those the browser places the
  // twin by for where it goes.)
  const twin = twinMetrics(name);
  const m = twin || metrics[key] || (STAND_INS[key] || []).map((s) => metrics[s.toLowerCase()]).find(Boolean) || metrics.calibri;
  const wa = m.win[0] * sizePx;
  const wd = m.win[1] * sizePx;
  // (A font with USE_TYPO_METRICS set placed by its typo ascent and descent, as the Morrow fonts are; installed ones
  // with it, Bahnschrift, Cascadia, and the other twins with it, by the same rule, not measured against Excel.)
  const [ca, cd] = m.typo?.[3] ? [m.typo[0] * sizePx, m.typo[1] * sizePx] : [wa, wd];
  return { wa, wd, ar: Math.round(ca), dr: Math.round(cd), base: twin || metrics[key] ? BASELINE_ABOVE_DESCENT : BASELINE_STAND_IN[key] ?? BASELINE_ABOVE_DESCENT };
}

/** Where Chromium puts the baseline in a line box of L px: floor((L - (ascent + descent)) / 2) + ascent below its top,
 *  with the Windows ascent and descent each rounded to a whole pixel and half the leading rounded down. Measured in
 *  Edge on Calibri 11 to 36, Cambria 20, Times New Roman 16, Arial 14 and Courier New 11 at several line heights (a
 *  zero-size marker on the baseline), exact for all; an earlier model with the hhea metrics drifted with the size
 *  (Calibri 36 placed 3.7px low, Arial 14 2px high). */
const chromeTop = (f, L) => Math.floor((L - f.ar - f.dr) / 2) + f.ar;

/** The line height (px) that puts the browser's baseline where Excel's is for bottom-aligned text, the font's Windows
 *  descent + BASELINE_ABOVE_DESCENT above the line's bottom. A line of ar + dr + 2k + t px (k whole, 0 <= t < 2) has
 *  its baseline dr + k + t above its bottom, so any height can be reached. */
export function cellLineHeight(name, sizePx) {
  const f = fontMetrics(name, sizePx);
  const above = f.wd + f.base - f.dr;
  const k = Math.floor(above);
  return f.ar + f.dr + 2 * k + (above - k);
}

/** Everything that places a font's text in a cell, in px: the single line's height (`lh`); how far to move the text
 *  down from where the browser puts it, for top- and middle-aligned single lines (`top`, `mid`); and for wrapped text,
 *  the line pitch and the shifts for bottom, top and middle alignment, on screen and in print (`pitch` gives the
 *  pitch, the font's row height, in each). Excel places a top-aligned line's baseline TOP_INSET + ascent below the
 *  top, centres a line's ascent + descent in the cell, and steps wrapped lines by the font's row height (align.xlsx:
 *  Calibri 11 wraps every 19.36px in print, its 15pt row). Chromium centres a line box of L in the cell (middle) or
 *  stacks lines of the pitch (wrapped), each with its baseline chromeTop below its top. */
export function cellTextMetrics(name, sizePx, pitch) {
  const f = fontMetrics(name, sizePx);
  const lh = cellLineHeight(name, sizePx);
  // (Excel's baseline below the top of a centred box of L, in this grid's coordinates.)
  const excelMid = (L) => (L - f.wa - f.wd) / 2 + f.wa + GRID_ORIGIN_Y;
  const excelTop = TOP_INSET + GRID_ORIGIN_Y + f.wa;
  const wrap = (P) => ({
    lh: P,
    bottom: P - chromeTop(f, P) - (f.wd + f.base),
    top: excelTop - chromeTop(f, P),
    mid: excelMid(P) - chromeTop(f, P),
  });
  return { lh, top: excelTop - chromeTop(f, lh), mid: excelMid(lh) - chromeTop(f, lh), wrap: wrap(pitch.screen), printWrap: wrap(pitch.print) };
}

/** Extra padding each side of a cell's text for fonts above 11pt: Excel's grows with the font, about 0.13px for each
 *  pixel of size beyond 11pt (Excel's PDF of rows.xlsx: Calibri 20 starts 1.6px further in, Calibri 36 4.2px,
 *  Cambria 20 1.8px; the right side alike, a Calibri 20 number ending 2px earlier). Smaller fonts get none. */
export const extraPadPx = (sizePx) => Math.max(0, 0.13 * (sizePx - 44 / 3));

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
