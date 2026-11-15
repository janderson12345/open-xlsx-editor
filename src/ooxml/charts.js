// A sheet's charts as the file has them (docs/CHARTS.md): where each sits (its drawing's two-cell anchor), its kind
// (column, bar, line, area, pie, doughnut; clustered, stacked or 100% stacked), its title, legend and series (names,
// categories and values, read from the cells they point to when they are in this workbook, else from the file's
// cache), and the value axis Excel would draw: its own minimum, maximum and major unit when the chart sets them, else
// Excel's automatic scale (measured over COM: test/axes-excel.json). Read only: the parts are saved as they are.
import { kids, kid, attr, NS } from './xml.js';
import { parseRange } from './refs.js';
import { emuToPx } from './units.js';
import { fontOf, axisTitle, seriesLook, chartLookOf } from './chartlook.js';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const val = (el, name, ns = C) => attr(kid(el, name, ns), 'val', null);

/** Excel's step for an axis: the smallest of 1, 2 and 5 times a power of ten that takes [lo, hi] out to whole steps in
 *  no more than ten. */
function stepFor(lo, hi) {
  const span = hi - lo || Math.abs(hi) || 1;
  let k = 10 ** Math.floor(Math.log10(span / 10));
  for (;;) {
    for (const m of [1, 2, 5]) {
      const u = m * k;
      if (Math.ceil(hi / u - 1e-9) - Math.floor(lo / u + 1e-9) <= 10) return u;
    }
    k *= 10;
  }
}
const tidy = (x) => Number(x.toPrecision(15));

/**
 * Excel's automatic value axis for the numbers drawn ({ min, max } of them): { min, max, major }. Zero is on the axis
 * when the numbers' spread is more than a sixth of the largest (of the smallest, all negative); else the axis starts
 * half the spread beyond them. The far end gets 5% of the spread more; both ends go out to whole major units.
 */
export function autoScale(lo, hi) {
  if (lo === 0 && hi === 0) return { min: 0, max: 1, major: 0.1 };
  let a;
  let b;
  if (lo >= 0) {
    a = lo === hi || (hi - lo) / hi > 1 / 6 ? 0 : lo - (hi - lo) / 2;
    b = hi + 0.05 * (hi - a);
  } else if (hi <= 0) {
    b = lo === hi || (hi - lo) / -lo > 1 / 6 ? 0 : hi + (hi - lo) / 2;
    a = lo - 0.05 * (b - lo);
  } else {
    a = lo - 0.05 * (hi - lo);
    b = hi + 0.05 * (hi - lo);
  }
  const major = stepFor(a, b);
  return { min: tidy(Math.floor(a / major + 1e-9) * major), max: tidy(Math.ceil(b / major - 1e-9) * major), major: tidy(major) };
}

/** The points of a strCache or numCache: [values] by index. */
function cache(el, numeric) {
  if (!el) return [];
  const n = Number(val(el, 'ptCount') ?? 0);
  const out = new Array(n).fill(null);
  for (const pt of kids(el, 'pt', C)) {
    const v = kid(pt, 'v', C)?.textContent ?? null;
    out[Number(attr(pt, 'idx', null))] = numeric ? (v == null || v === '' ? null : Number(v)) : v;
  }
  return out;
}

/** A series part (tx, cat or val): its reference's cells read from the workbook when it can be, else its cache. */
function data(wb, el, numeric) {
  if (!el) return { values: [], ref: null };
  const r = kid(el, 'numRef', C) || kid(el, 'strRef', C) || kid(el, 'multiLvlStrRef', C);
  const lit = kid(el, 'numLit', C) || kid(el, 'strLit', C);
  const f = r ? kid(r, 'f', C)?.textContent : null;
  const kept = cache(kid(r || lit, 'numCache', C) || kid(r || lit, 'strCache', C) || (lit && lit) || kid(r, 'multiLvlStrCache', C), numeric);
  const m = f && /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(f);
  const si = m ? wb.sheetIndexOf((m[1] ?? m[2]).replace(/''/g, "'")) : -1;
  const area = m ? parseRange(m[3].replace(/\$/g, '')) : null;
  if (si < 0 || !area) return { values: kept, ref: f };
  const sheet = wb.sheet(si);
  const out = [];
  for (let r = area.r0; r <= area.r1; r++) {
    for (let c = area.c0; c <= area.c1; c++) {
      const v = sheet.cells.get(r)?.get(c)?.v ?? null;
      out.push(numeric ? (typeof v === 'number' ? v : null) : v == null ? null : String(v));
    }
  }
  return { values: out, ref: f, code: kid(kid(r, 'numCache', C), 'formatCode', C)?.textContent || null };
}

/** The text of a title, or null when the chart has none (its own rich text, the cell it points to). */
function titleText(wb, chart, series, kind) {
  const t = kid(chart, 'title', C);
  if (!t) return null;
  const rich = kid(kid(t, 'tx', C), 'rich', C);
  if (rich) return [...rich.getElementsByTagNameNS(A, 't')].map((x) => x.textContent).join('');
  const ref = kid(kid(t, 'tx', C), 'strRef', C);
  if (ref) return data(wb, kid(t, 'tx', C), false).values.join('');
  // (A title with no text of its own: the series' name when there is one series, or for a pie, which draws only its
  // first; else Excel's 'Chart Title'. Excel's reading of charts written here: test/chartnew-readback-excel.json.)
  return series.length === 1 || (kind === 'pie' && series.length) ? series[0].name : 'Chart Title';
}

/** A shape's solid fill or line colour, as the theme or RGB the file names: { theme, rgb, lumMod, lumOff } or null. */
function colorOf(spPr) {
  const fill = kid(spPr, 'solidFill', A) || kid(kid(spPr, 'ln', A), 'solidFill', A);
  const c = fill && (kid(fill, 'srgbClr', A) || kid(fill, 'schemeClr', A));
  if (!c) return null;
  const mod = (name) => { const x = kid(c, name, A); return x ? Number(attr(x, 'val', null)) / 100000 : null; };
  return c.localName === 'srgbClr' ? { rgb: attr(c, 'val', null) } : { scheme: attr(c, 'val', null), lumMod: mod('lumMod'), lumOff: mod('lumOff') };
}

const KINDS = { barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter' };

/** One chart group's settings: { kind, dir, grouping, gapWidth, overlap, secondary }. */
function groupOf(g, secondary) {
  const grouping = val(g, 'grouping') || 'clustered';
  return { kind: KINDS[g.localName], dir: val(g, 'barDir') === 'bar' ? 'bar' : 'col', grouping, gapWidth: Number(val(g, 'gapWidth') ?? 150),
    overlap: Number(val(g, 'overlap') ?? (/stacked/i.test(grouping) ? 100 : 0)), secondary };
}

/** A value axis's own settings: { min, max, major, gridlines, code, linked, hidden, title, font, pos }. */
function axisSettings(ax) {
  const scaling = kid(ax, 'scaling', C);
  const num = (name) => { const v = val(scaling, name) ?? val(ax, name); return v == null ? null : Number(v); };
  return { min: num('min'), max: num('max'), major: ax && val(ax, 'majorUnit') != null ? Number(val(ax, 'majorUnit')) : null, gridlines: !!kid(ax, 'majorGridlines', C),
    code: attr(kid(ax, 'numFmt', C), 'formatCode', null), linked: attr(kid(ax, 'numFmt', C), 'sourceLinked', null) === '1', hidden: val(ax, 'delete') === '1',
    title: axisTitle(ax), font: fontOf(ax), pos: val(ax, 'axPos') };
}

/** One chart part: { kind, dir, grouping, title, legend, series, groups, axis, axis2, cats, look } (kind null for a
 *  kind not drawn yet). A combo chart's series each say their group's kind and whether they are on the secondary axes
 *  (the first group's axes are the primary ones). */
function readChart(wb, doc) {
  const space = doc.documentElement;
  const chart = kid(space, 'chart', C);
  const plot = kid(chart, 'plotArea', C);
  const groupEls = plot ? [...plot.childNodes].filter((n) => n.nodeType === 1 && n.namespaceURI === C && KINDS[n.localName]) : [];
  if (!groupEls.length) return { kind: null };
  const group = groupEls[0];
  const kind = KINDS[group.localName];
  const idsOf = (g) => kids(g, 'axId', C).map((a) => attr(a, 'val', null));
  const primary = idsOf(group);
  const groups = groupEls.map((g) => groupOf(g, idsOf(g).join() !== primary.join()));
  let n = 0;
  const series = groupEls.flatMap((g, gi) => kids(g, 'ser', C).map((s) => {
    const i = n++;
    const name = data(wb, kid(s, 'tx', C), false).values[0] ?? kid(kid(s, 'tx', C), 'v', C)?.textContent ?? `Series${i + 1}`;
    // (A scatter chart's series have x and y values: xs numbers, or 1, 2, 3... when they are text, as Excel plots them.)
    const values = data(wb, kid(s, 'val', C) || kid(s, 'yVal', C), true);
    const xVal = kid(s, 'xVal', C);
    const xs = xVal ? data(wb, xVal, true).values.map((x, k) => (typeof x === 'number' ? x : k + 1)) : null;
    const line = kid(kid(s, 'spPr', C), 'ln', A);
    return { name, values: values.values, code: values.code, cats: data(wb, kid(s, 'cat', C) || xVal, false).values, xs, color: colorOf(kid(s, 'spPr', C)), index: Number(val(s, 'idx') ?? i),
      order: Number(val(s, 'order') ?? i), marker: val(kid(s, 'marker', C), 'symbol'), line: !kid(line, 'noFill', A), smooth: val(s, 'smooth') === '1',
      type: groups[gi].kind, group: gi, secondary: groups[gi].secondary, look: seriesLook(s) };
  })).sort((a, b) => a.order - b.order);
  // (The value axes: a scatter chart's is its upright one, the other its x axis; the secondary one, the value axis of
  // the other groups' axes.)
  const axes = [...plot.childNodes].filter((x) => x.nodeType === 1 && /^(catAx|valAx|dateAx|serAx)$/.test(x.localName));
  const byId = (ids) => axes.filter((a) => ids.includes(val(a, 'axId')));
  const mine = byId(primary);
  const valAxes = mine.filter((a) => a.localName === 'valAx');
  const valAx = valAxes.find((ax) => ['l', 'r'].includes(val(ax, 'axPos'))) || valAxes[0] || kids(plot, 'valAx', C)[0];
  const xAx = kind === 'scatter' ? valAxes.find((ax) => ax !== valAx) : null;
  const catAx = mine.find((a) => a.localName !== 'valAx') || xAx;
  const second = groupEls.map(idsOf).find((ids) => ids.join() !== primary.join());
  const valAx2 = second ? byId(second).find((a) => a.localName === 'valAx') : null;
  const legend = kid(chart, 'legend', C);
  return {
    kind,
    dir: groups[0].dir,
    grouping: groups[0].grouping,
    varyColors: val(group, 'varyColors') === '1',
    gapWidth: groups[0].gapWidth,
    overlap: groups[0].overlap,
    holeSize: Number(val(group, 'holeSize') ?? 50),
    title: val(chart, 'autoTitleDeleted') === '1' ? null : titleText(wb, chart, series, kind),
    ownTitle: !!kid(kid(chart, 'title', C), 'tx', C),
    // (Data labels showing values: the chart's own setting, or any series'.)
    labels: groupEls.flatMap((g) => [kid(g, 'dLbls', C), ...kids(g, 'ser', C).map((s) => kid(s, 'dLbls', C))]).some((d) => val(d, 'showVal') === '1'),
    legend: legend ? val(legend, 'legendPos') || 'r' : null,
    series,
    groups,
    cats: series[0]?.cats || [],
    axis: axisSettings(valAx),
    axis2: valAx2 ? axisSettings(valAx2) : null,
    catHidden: val(catAx, 'delete') === '1',
    catTitle: axisTitle(catAx),
    catFont: fontOf(catAx),
    xGridlines: !!kid(xAx, 'majorGridlines', C),
    look: chartLookOf(space, chart, plot, groupEls),
  };
}

/** A two-cell anchor's corner: { r, c, dx, dy } (dx and dy in pixels into the cell). */
const corner = (el) => ({ c: Number(kid(el, 'col', XDR)?.textContent || 0), r: Number(kid(el, 'row', XDR)?.textContent || 0), dx: emuToPx(kid(el, 'colOff', XDR)?.textContent || 0), dy: emuToPx(kid(el, 'rowOff', XDR)?.textContent || 0) });

/** Sheet i's charts: [{ from, to (corners), name, ...readChart }], in the drawing's order. */
export function readCharts(wb, i) {
  const path = wb.sheets[i]?.path;
  if (!path) return [];
  const out = [];
  for (const rel of wb.pkg.readRels(path).filter((r) => r.type === `${NS.r}/drawing`)) {
    const doc = wb.pkg.xml(rel.target);
    if (!doc) continue;
    const drawingRels = wb.pkg.readRels(rel.target);
    for (const anchor of kids(doc.documentElement, 'twoCellAnchor', XDR)) {
      const frame = kid(anchor, 'graphicFrame', XDR);
      const ref = frame && [...frame.getElementsByTagNameNS(C, 'chart')][0];
      const id = ref && attr(ref, 'id', NS.r);
      const target = id && drawingRels.find((r) => r.id === id)?.target;
      const chartDoc = target && wb.pkg.xml(target);
      if (!chartDoc) continue;
      const name = attr(kid(kid(frame, 'nvGraphicFramePr', XDR), 'cNvPr', XDR), 'name', null) || 'Chart';
      out.push({ from: corner(kid(anchor, 'from', XDR)), to: corner(kid(anchor, 'to', XDR)), name, ...readChart(wb, chartDoc) });
    }
  }
  return out;
}

/** The value axis a chart draws (the secondary one with `secondary`): its own settings where it has them, Excel's
 *  automatic scale for the rest, over the series on that axis. 100% stacked charts run 0 to 1 (0% to 100%); stacked
 *  ones scale the stacks' sums; groups on one axis scale together. (Excel's secondary axis over Cost, 45 to 90, beside
 *  Sales' columns of 60 to 140: 0 to 100 by 10, the primary 0 to 160 by 20: test/chartformat-excel.json.) */
export function valueAxis(chart, secondary = false) {
  let lo = Infinity;
  let hi = -Infinity;
  const own = secondary ? chart.axis2 : chart.axis;
  const groups = chart.groups || [{ grouping: chart.grouping, secondary: false }];
  let percent = false;
  groups.forEach((g, gi) => {
    if (!!g.secondary !== secondary) return;
    const list = chart.series.filter((s) => (s.group ?? 0) === gi);
    const n = Math.max(...list.map((s) => s.values.length), 0);
    if (g.grouping === 'percentStacked') { percent = true; lo = Math.min(lo, 0); hi = Math.max(hi, 1); }
    else if (g.grouping === 'stacked') {
      for (let i = 0; i < n; i++) {
        let up = 0;
        let down = 0;
        for (const s of list) { const v = s.values[i]; if (typeof v === 'number') { if (v >= 0) up += v; else down += v; } }
        lo = Math.min(lo, down); hi = Math.max(hi, up);
      }
    } else {
      for (const s of list) for (const v of s.values) if (typeof v === 'number') { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    }
  });
  if (lo === Infinity) { lo = 0; hi = 0; }
  const auto = percent && groups.filter((g) => !!g.secondary === secondary).length === 1 ? { min: 0, max: 1, major: 0.1 } : autoScale(lo, hi);
  const min = own?.min ?? auto.min;
  const max = own?.max ?? auto.max;
  const major = own?.major ?? (own?.min != null || own?.max != null ? stepFor(min, max) : auto.major);
  return { min, max, major };
}
