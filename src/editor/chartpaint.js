// The colours and fonts a chart is drawn in (chartview.js, chartplot.js): its own where its part sets them
// (ooxml/chartlook.js), else Excel's default style's: text in 65% grey (#595959), the title 14pt, axis labels and the
// legend 9pt, axis titles 10pt, data labels 9pt in 75% grey (#404040), all in the theme's body font; series and points
// in the theme's accents by their index, darker past six.
import { resolveDrawingColor, resolveColor } from '../ooxml/theme.js';
import { fontStack } from './fonts.js';

export const SVG = 'http://www.w3.org/2000/svg';
export const GREY = '#595959';
export const LINE = '#D9D9D9';
const LABEL = '#404040';
const PX = 4 / 3;

export const el = (name, attrs = {}, text = null) => {
  const x = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) if (v != null) x.setAttribute(k, String(v));
  if (text != null) x.textContent = text;
  return x;
};

/** A part's colour as CSS: 'none' for no fill or line, the fallback when the part says nothing. */
export function paint(c, theme, fallback) {
  if (!c) return fallback;
  if (c.none) return 'none';
  return resolveDrawingColor(c, theme) || fallback;
}

const shade = (hex, f) => `#${[1, 3, 5].map((k) => Math.round(parseInt(hex.slice(k, k + 2), 16) * f).toString(16).padStart(2, '0')).join('')}`;
/** The theme's accent for series (or point) index i: accent 1 to 6 in turn, darker past six. */
export function accent(theme, i) {
  const a = resolveColor({ theme: 4 + (i % 6) }, theme) || ['#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E'][i % 6];
  return i < 6 ? a : shade(a, 0.6);
}

/** A series' colour: a column's, area's or slice's fill, a line's or scatter series' line, its own when it has one. */
export function seriesColour(s, theme) {
  const solid = (c) => (c && !c.none ? c : null);
  // (A line's colour is its line's; a scatter series of markers only has none, and is drawn in its markers' accent.)
  const own = s.type === 'line' || s.type === 'scatter' ? solid(s.look?.line) || solid(s.look?.fill) : s.look?.fill;
  const auto = accent(theme, s.index ?? 0);
  const c = paint(own, theme, auto);
  return c === 'none' ? 'none' : c;
}

const DEFAULTS = { title: [14, GREY], axis: [9, GREY], legend: [9, GREY], labels: [9, LABEL], axisTitle: [10, GREY] };
/** A text's font for drawing: { family, size (px), bold, color }, from its own (`own`, chartlook.js fontOf), the
 *  chart's base text properties, then the default style's for `role`. */
export function fontFor(chart, own, role, theme) {
  const base = chart.look?.fonts?.base || {};
  const [pt, colour] = DEFAULTS[role];
  const face = own?.name || base.name || '+mn-lt';
  const name = face === '+mn-lt' || face === '+mn-ea' ? theme?.fonts?.minor || 'Aptos Narrow' : face === '+mj-lt' ? theme?.fonts?.major || 'Aptos Display' : face;
  return { family: fontStack(name), size: (own?.size ?? (role === 'title' || role === 'axisTitle' ? null : base.size) ?? pt) * PX, bold: !!(own?.bold ?? false), color: paint(own?.color || base.color, theme, colour) };
}
/** SVG text attributes for a font. */
export const textAttrs = (f) => ({ 'font-family': f.family, 'font-size': Number(f.size.toFixed(2)), 'font-weight': f.bold ? 'bold' : null, fill: f.color });
/** Text width, roughly, for laying out labels (digits are about half the font size wide). */
export const widthOf = (s, size, bold = false) => String(s).length * size * (bold ? 0.56 : 0.52);
