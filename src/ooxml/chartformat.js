// A chart's formatting, as Excel writes it (measured over COM on a clustered column chart:
// tools/excel-compare/excel-chart-format.ps1, test/chartformat-excel.json): its axes' titles, the fill and line of a
// series, a point, the chart area and the plot area, and the font of the title, the axes' labels and titles, the legend
// and the data labels. Each changes the chart part and nothing else, through chartedit.js's change (c: prefix only).
import { change } from './chartedit.js';
import { chartLook } from './chartxml.js';
import { rootOf, kidOf, kidsOf, pathOf, elementAt, textOf, valOf, splice, placeIn, opened, setAttr, esc } from './charttext.js';

const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
const solid = (hex) => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
const scheme = (clr, mod) => (mod ? `<a:solidFill><a:schemeClr val="${clr}"><a:lumMod val="${mod[0]}"/><a:lumOff val="${mod[1]}"/></a:schemeClr></a:solidFill>` : `<a:solidFill><a:schemeClr val="${clr}"/></a:solidFill>`);
const HEX = /^[0-9A-F]{6}$/;

/** The plot area's chart groups and axes: { plot, groups: [el], axes: [{ el, id, pos, kind }], primary: [ids],
 *  secondary: [ids] or null }. The first group's axes are the primary ones. */
export function plotOf(x) {
  const plot = pathOf(x, rootOf(x), 'c:chart', 'c:plotArea');
  if (!plot) throw new Error("This chart can't be edited here yet.");
  const kids = kidsOf(x, plot);
  const groups = kids.filter((k) => /^c:\w+Chart$/.test(k.name));
  const axes = kids.filter((k) => /^c:(catAx|valAx|dateAx|serAx)$/.test(k.name)).map((el) => ({ el, kind: el.name.slice(2), id: valOf(x, kidOf(x, el, 'c:axId')), pos: valOf(x, kidOf(x, el, 'c:axPos')) }));
  const ids = (g) => kidsOf(x, g).filter((k) => k.name === 'c:axId').map((k) => valOf(x, k));
  const primary = groups.length ? ids(groups[0]) : [];
  const other = groups.map(ids).find((p) => p.length && p.join() !== primary.join());
  return { plot, groups, axes, primary, secondary: other || null };
}

/** An axis by where it is drawn: 'h' (the primary horizontal one), 'v' (the primary upright one), 'v2' (the secondary
 *  value axis). */
export function axisOf(x, which) {
  const p = plotOf(x);
  if (which === 'v2') return p.secondary && p.axes.find((a) => p.secondary.includes(a.id) && a.kind === 'valAx') || null;
  const mine = p.axes.filter((a) => p.primary.includes(a.id));
  return mine.find((a) => (which === 'h' ? ['b', 't'] : ['l', 'r']).includes(a.pos)) || null;
}

/** The axis's title, as Excel's Axis.HasTitle and AxisTitle.Text write it: 10pt in the axis's grey, upright text
 *  turned a quarter (-5400000) on an upright axis; `text` its own, 'auto' Excel's "Axis Title" (no text of its own),
 *  null none (the title taken out: Excel's part is then as it was before). */
export function setAxisTitle(wb, si, k, which, text) {
  return change(wb, si, k, (x) => {
    const axis = axisOf(x, which);
    if (!axis) throw new Error('The chart has no such axis.');
    const old = kidOf(x, axis.el, 'c:title');
    if (text == null) return old ? splice(x, old.start, old.end, '') : x;
    const rot = ['l', 'r'].includes(axis.pos) ? -5400000 : 0;
    // (A title already there keeps its font: its own text takes the font its txPr has, as Excel's does.)
    const def = old ? defsOf(x, old)[0] : null;
    const pPr = def ? textOf(x, def) : chartLook.DEF(1000);
    const rich = text === 'auto' ? '' : `<c:tx><c:rich>${chartLook.BODY(rot)}<a:lstStyle/><a:p><a:pPr>${pPr}</a:pPr><a:r><a:rPr lang="en-US"/><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
    if (old) {
      const tx = kidOf(x, old, 'c:tx');
      return tx ? splice(x, tx.start, tx.end, rich) : splice(x, old.inner, old.inner, rich);
    }
    const title = `<c:title>${rich}<c:overlay val="0"/>${chartLook.SP_NONE}<c:txPr>${chartLook.BODY(rot)}<a:lstStyle/><a:p><a:pPr>${chartLook.DEF(1000)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:title>`;
    const at = placeIn(x, axis.el, ['c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx']);
    return splice(x, at, at, title);
  });
}

// ---- Fonts ----

/** The a:defRPr elements that set an element's text: in its txPr, and in its own rich text (a title's). */
function defsOf(x, el) {
  const out = [];
  for (const host of [kidOf(x, el, 'c:txPr'), pathOf(x, el, 'c:tx', 'c:rich')]) {
    if (!host) continue;
    for (const p of kidsOf(x, host).filter((c) => c.name === 'a:p')) {
      const d = pathOf(x, p, 'a:pPr', 'a:defRPr');
      if (d) out.push(d);
    }
  }
  return out.sort((a, b) => a.start - b.start);
}

/** A defRPr's text with `font` given: { name, size (points), bold, color ('RRGGBB') }, each only when set. Excel's
 *  Font object (ChartTitle.Font, TickLabels.Font, Legend.Font, DataLabels.Font, AxisTitle.Font) changes these and
 *  nothing else: sz, b, the fill, and latin, ea and cs all the one name. */
export function fontDef(t, font) {
  let out = opened(t);
  const gt = out.indexOf('>');
  let open = out.slice(0, gt + 1);
  let body = out.slice(gt + 1, out.lastIndexOf('</'));
  if (font.size != null) open = setAttr(open, 'sz', Math.round(font.size * 100));
  if (font.bold != null) open = setAttr(open, 'b', font.bold ? 1 : 0, ['sz']);
  const kids = (s) => { const el = { start: 0, inner: 0, close: s.length, end: s.length }; return kidsOf(s, el); };
  if (font.color) {
    const f = kids(body).find((k) => FILLS.includes(k.name));
    const ln = kids(body).find((k) => k.name === 'a:ln');
    body = f ? splice(body, f.start, f.end, solid(font.color)) : splice(body, ln ? ln.end : 0, ln ? ln.end : 0, solid(font.color));
  }
  if (font.name) {
    const faces = kids(body).filter((k) => ['a:latin', 'a:ea', 'a:cs'].includes(k.name));
    const n = esc(font.name);
    const text = `<a:latin typeface="${n}"/><a:ea typeface="${n}"/><a:cs typeface="${n}"/>`;
    if (faces.length) body = body.slice(0, faces[0].start) + text + body.slice(faces[faces.length - 1].end);
    else { const f = kids(body).filter((k) => FILLS.includes(k.name) || ['a:ln', 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill'].includes(k.name)).pop(); body = splice(body, f ? f.end : 0, f ? f.end : 0, text); }
  }
  return `${open}${body}${out.slice(out.lastIndexOf('</'))}`;
}

/** The elements whose text a font target names: 'title', 'legend', 'h', 'v', 'v2' (an axis's labels), 'hTitle',
 *  'vTitle' (an axis's title), 'labels' (every series' data labels, or series `series`'s). */
function textTargets(x, target, series) {
  const chart = kidOf(x, rootOf(x), 'c:chart');
  if (target === 'title' || target === 'legend') { const el = kidOf(x, chart, `c:${target}`); return el ? [el] : []; }
  if (target === 'hTitle' || target === 'vTitle') { const a = axisOf(x, target[0]); const t = a && kidOf(x, a.el, 'c:title'); return t ? [t] : []; }
  if (['h', 'v', 'v2'].includes(target)) { const a = axisOf(x, target); return a ? [a.el] : []; }
  if (target === 'labels') return seriesOf(x).filter((s) => series == null || s.idx === series).map((s) => kidOf(x, s.el, 'c:dLbls')).filter(Boolean);
  throw new Error(`No chart element ${target}`);
}

/** The font of a chart's title, legend, axes' labels and titles or data labels (fontDef). An element without text
 *  properties of its own gets them. */
export function setChartFont(wb, si, k, target, font, series = null) {
  return change(wb, si, k, (x0) => {
    let x = x0;
    const els = textTargets(x, target, series);
    if (!els.length) throw new Error('The chart has nothing there to format.');
    for (const el of els.sort((a, b) => b.start - a.start)) {
      let defs = defsOf(x, el);
      if (!defs.length) {
        const at = placeIn(x, el, ['c:crossAx', 'c:dLblPos', 'c:showLegendKey', 'c:extLst', 'c:legendEntry']);
        x = splice(x, at, at, '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>');
        defs = defsOf(x, elementAt(x, el.start));
      }
      for (const d of defs.sort((a, b) => b.start - a.start)) x = splice(x, d.start, d.end, fontDef(textOf(x, d), font));
    }
    return x;
  });
}

// ---- Fills and lines ----

/** Every series in the part, in the order written: [{ el, idx, group }]. */
export function seriesOf(x) {
  const { groups } = plotOf(x);
  return groups.flatMap((g) => kidsOf(x, g).filter((k) => k.name === 'c:ser').map((el) => ({ el, idx: Number(valOf(x, kidOf(x, el, 'c:idx'))), group: g })));
}

/** An spPr's text with its fill and line set: each a colour ('RRGGBB'), 'none', or left alone (undefined). A line
 *  given a colour keeps its width and ends. */
export function spPrWith(t, { fill, line }) {
  let out = opened(t);
  const el = () => ({ start: 0, inner: out.indexOf('>') + 1, close: out.lastIndexOf('</'), end: out.length });
  const pick = (v) => (v === 'none' ? '<a:noFill/>' : solid(v));
  if (fill !== undefined) {
    const f = kidsOf(out, el()).find((k) => FILLS.includes(k.name));
    out = f ? splice(out, f.start, f.end, pick(fill)) : splice(out, el().inner, el().inner, pick(fill));
  }
  if (line !== undefined) {
    let ln = kidsOf(out, el()).find((k) => k.name === 'a:ln');
    if (!ln) {
      const at = placeIn(out, el(), ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']);
      out = splice(out, at, at, '<a:ln></a:ln>');
      ln = kidsOf(out, el()).find((k) => k.name === 'a:ln');
    }
    const lt = opened(textOf(out, ln));
    const inner = { start: 0, inner: lt.indexOf('>') + 1, close: lt.lastIndexOf('</'), end: lt.length };
    const f = kidsOf(lt, inner).find((k) => FILLS.includes(k.name));
    const next = f ? splice(lt, f.start, f.end, pick(line)) : splice(lt, inner.inner, inner.inner, pick(line));
    out = splice(out, ln.start, ln.end, next);
  }
  return out;
}

/** The look Excel's default style gives back ('auto'): a series' or point's accent, the chart area white with a 15%
 *  grey line, the plot area without either. */
function autoOf(target, family, idx) {
  if (target === 'chart') return { fill: scheme('bg1'), line: scheme('tx1', [15000, 85000]) };
  if (target === 'plot') return { fill: '<a:noFill/>', line: '<a:noFill/>' };
  return { fill: scheme(`accent${(idx % 6) + 1}`), line: family === 'line' || family === 'scatter' ? scheme(`accent${(idx % 6) + 1}`) : '<a:noFill/>' };
}
/** spPrWith, where 'auto' puts back the default style's own fill or line (autoOf's text). */
function spPrLook(t, look, auto) {
  let out = spPrWith(t, { fill: look.fill === 'auto' ? undefined : look.fill, line: look.line === 'auto' ? undefined : look.line });
  const back = (part) => {
    const el = { start: 0, inner: out.indexOf('>') + 1, close: out.lastIndexOf('</'), end: out.length };
    if (part === 'fill') { const f = kidsOf(out, el).find((k) => FILLS.includes(k.name)); out = f ? splice(out, f.start, f.end, auto.fill) : splice(out, el.inner, el.inner, auto.fill); return; }
    const ln = kidsOf(out, el).find((k) => k.name === 'a:ln');
    if (!ln) return;
    const lt = opened(textOf(out, ln));
    const f = kidsOf(lt, { start: 0, inner: lt.indexOf('>') + 1, close: lt.lastIndexOf('</'), end: lt.length }).find((k) => FILLS.includes(k.name));
    out = splice(out, ln.start, ln.end, f ? splice(lt, f.start, f.end, auto.line) : lt);
  };
  if (look.fill === 'auto') back('fill');
  if (look.line === 'auto') back('line');
  return out;
}

const FAMILY = { 'c:barChart': 'bar', 'c:bar3DChart': 'bar', 'c:lineChart': 'line', 'c:line3DChart': 'line', 'c:areaChart': 'area', 'c:area3DChart': 'area', 'c:pieChart': 'pie', 'c:pie3DChart': 'pie', 'c:doughnutChart': 'doughnut', 'c:scatterChart': 'scatter' };

/**
 * The fill and line of the chart area ('chart'), the plot area ('plot'), a series ({ series: idx }) or one of its
 * points ({ series, point }), as Excel's Format.Fill.ForeColor.RGB and Format.Line.ForeColor.RGB write them: a colour
 * ('RRGGBB') as an srgbClr in place of the old one; 'none' a noFill; 'auto' the default style's own. A point given its
 * own look is a c:dPt (idx, invertIfNegative on a bar, bubble3D, the series' line, a unique id), in its place among
 * the series' points; a pie's slices have theirs already.
 */
export function setChartFill(wb, si, k, target, look) {
  return change(wb, si, k, (x) => {
    for (const v of [look.fill, look.line]) if (v !== undefined && v !== 'none' && v !== 'auto' && !HEX.test(v)) throw new Error(`Not a colour: ${v}`);
    const root = rootOf(x);
    if (target === 'chart' || target === 'plot') {
      const host = target === 'chart' ? root : plotOf(x).plot;
      const sp = kidOf(x, host, 'c:spPr');
      if (!sp) throw new Error("This chart can't be formatted here yet.");
      return splice(x, sp.start, sp.end, spPrLook(textOf(x, sp), look, autoOf(target)));
    }
    const ser = seriesOf(x).find((s) => s.idx === target.series);
    if (!ser) throw new Error('No such series');
    const family = FAMILY[ser.group.name];
    if (target.point == null) {
      let sp = kidOf(x, ser.el, 'c:spPr');
      if (!sp) { const at = placeIn(x, ser.el, kidsOf(x, ser.el).map((c) => c.name).filter((n) => !['c:idx', 'c:order', 'c:tx'].includes(n))); x = splice(x, at, at, '<c:spPr></c:spPr>'); sp = kidOf(x, elementAt(x, ser.el.start), 'c:spPr'); }
      return splice(x, sp.start, sp.end, spPrLook(textOf(x, sp), look, autoOf('series', family, ser.idx)));
    }
    return pointFill(x, ser, family, target.point, look);
  });
}

/** A point's own look (setChartFill): its c:dPt changed, or one made. 'auto' on a bar's point takes its dPt out. */
function pointFill(x, ser, family, p, look) {
  const pts = kidsOf(x, ser.el).filter((c) => c.name === 'c:dPt');
  const old = pts.find((d) => Number(valOf(x, kidOf(x, d, 'c:idx'))) === p);
  const pie = family === 'pie' || family === 'doughnut';
  if (old) {
    if (!pie && look.fill === 'auto' && look.line === undefined) return splice(x, old.start, old.end, '');
    // (A slice given its own colour gets a unique id too, as Excel's did.)
    const sp = kidOf(x, old, 'c:spPr');
    const id = kidOf(x, old, 'c:extLst') ? '' : uidExt(x);
    return splice(x, sp.start, sp.end, spPrLook(textOf(x, sp), look, autoOf('series', family, pie ? p : ser.idx)) + id);
  }
  if (look.fill === 'auto' && look.line === undefined) return x;
  // (Excel's new point: the series' own spPr with the fill changed, and a unique id numbered after the part's others.)
  const sp = kidOf(x, ser.el, 'c:spPr');
  const spText = spPrLook(sp ? textOf(x, sp) : '<c:spPr></c:spPr>', look, autoOf('series', family, ser.idx));
  const dPt = `<c:dPt><c:idx val="${p}"/>${family === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}${family === 'line' || family === 'scatter' ? '' : '<c:bubble3D val="0"/>'}${spText}${uidExt(x)}</c:dPt>`;
  const after = pts.find((d) => Number(valOf(x, kidOf(x, d, 'c:idx'))) > p);
  const at = after ? after.start : placeIn(x, ser.el, ['c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:xVal', 'c:yVal', 'c:smooth', 'c:shape', 'c:extLst']);
  return splice(x, at, at, dPt);
}

/** A unique id for a new point, as Excel numbers them: after the part's others, with their GUID (none without). */
function uidExt(x) {
  const guid = /c16:uniqueId val="\{\d{8}-([0-9A-F-]{27})\}"/i.exec(x)?.[1];
  const n = [...x.matchAll(/c16:uniqueId val="\{(\d{8})-/g)].reduce((m, a) => Math.max(m, Number(a[1]) + 1), 0);
  return guid ? `<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{${String(n).padStart(8, '0')}-${guid}}"/></c:ext></c:extLst>` : '';
}
