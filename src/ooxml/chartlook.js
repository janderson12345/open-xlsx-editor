// A chart's own formatting as its part has it (charts.js readChart), for drawing it in its own look: fills and lines
// (a shape's spPr: a colour, none, or nothing said, which is the default style's), fonts (a text's defRPr and run
// properties: name, size, bold, colour) and the axes' titles. What chartformat.js writes reads back here, and so does
// what Excel writes (test/chartformat-excel.json, chartlook.test.js).
import { kids, kid, attr } from './xml.js';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const FILLS = ['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'];

/** A DrawingML colour element's colour: { rgb } or { scheme }, with its lumMod and lumOff (fractions) when it has them;
 *  a system colour as the RGB it was last. */
function colourIn(fill) {
  const c = fill && [...fill.childNodes].find((n) => n.nodeType === 1 && n.namespaceURI === A && /Clr$/.test(n.localName));
  if (!c) return null;
  const mod = (name) => { const x = kid(c, name, A); return x ? Number(attr(x, 'val', null)) / 100000 : null; };
  const out = c.localName === 'schemeClr' ? { scheme: attr(c, 'val', null) } : { rgb: attr(c, c.localName === 'sysClr' ? 'lastClr' : 'val', null) };
  const lumMod = mod('lumMod');
  const lumOff = mod('lumOff');
  if (lumMod != null) out.lumMod = lumMod;
  if (lumOff != null) out.lumOff = lumOff;
  return out;
}

/** The fill a parent's first fill child gives: a colour, { none: true }, or null (none said, or not solid). */
function fillIn(el) {
  const f = el && [...el.childNodes].find((n) => n.nodeType === 1 && n.namespaceURI === A && FILLS.includes(n.localName));
  if (!f) return null;
  if (f.localName === 'noFill') return { none: true };
  return f.localName === 'solidFill' ? colourIn(f) : null;
}

/** A shape's look from its spPr: { fill, line, width (points) }, each null when the part says nothing of it. */
export function shapeLook(spPr) {
  const ln = kid(spPr, 'ln', A);
  const w = attr(ln, 'w', null);
  return { fill: fillIn(spPr), line: fillIn(ln), width: w == null ? null : Number(w) / 12700 };
}

/** The font an element's text is set in: { name, size (points), bold, color }, each null when not said. A title's
 *  own text (its first run, then its paragraph's defaults) comes before its txPr; '+mn-lt' is the theme's body font. */
export function fontOf(el) {
  const out = { name: null, size: null, bold: null, color: null };
  if (!el) return out;
  const rich = kid(kid(el, 'tx', C), 'rich', C);
  const p = kid(rich, 'p', A);
  const props = [kid(kid(p, 'r', A), 'rPr', A), kid(kid(p, 'pPr', A), 'defRPr', A), kid(kid(kid(kid(el, 'txPr', C), 'p', A), 'pPr', A), 'defRPr', A)];
  for (const rp of props.filter(Boolean)) {
    const sz = attr(rp, 'sz', null);
    const b = attr(rp, 'b', null);
    const face = attr(kid(rp, 'latin', A), 'typeface', null);
    out.size ??= sz == null ? null : Number(sz) / 100;
    out.bold ??= b == null ? null : b === '1' || b === 'true';
    out.name ??= face;
    out.color ??= fillIn(rp);
  }
  return out;
}

/** An axis's title: null for none, else { text (its own, or Excel's "Axis Title"), own, font }. */
export function axisTitle(ax) {
  const t = kid(ax, 'title', C);
  if (!t) return null;
  const rich = kid(kid(t, 'tx', C), 'rich', C);
  const text = rich ? [...rich.getElementsByTagNameNS(A, 't')].map((x) => x.textContent).join('') : 'Axis Title';
  return { text, own: !!rich, font: fontOf(t) };
}

/** A series' look: { fill, line, width } of its spPr, and its points' own fills by index. */
export function seriesLook(ser) {
  const points = new Map();
  for (const d of kids(ser, 'dPt', C)) points.set(Number(attr(kid(d, 'idx', C), 'val', 0)), shapeLook(kid(d, 'spPr', C)));
  return { ...shapeLook(kid(ser, 'spPr', C)), points };
}

/** The chart's own fonts and areas: { area, plot (shapeLooks), fonts: { base, title, legend, labels } }; the part's own
 *  text properties (chartSpace's txPr) are the base the others fall back to. */
export function chartLookOf(space, chart, plot, groups) {
  const labels = groups.flatMap((g) => [...kids(g, 'ser', C).map((s) => kid(s, 'dLbls', C)), kid(g, 'dLbls', C)]).find((d) => d && kid(d, 'txPr', C));
  return {
    area: shapeLook(kid(space, 'spPr', C)),
    plot: shapeLook(kid(plot, 'spPr', C)),
    fonts: { base: fontOf(space), title: fontOf(kid(chart, 'title', C)), legend: fontOf(kid(chart, 'legend', C)), labels: fontOf(labels) },
  };
}
