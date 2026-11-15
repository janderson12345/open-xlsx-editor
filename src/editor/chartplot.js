// A chart's axes and plot (chartview.js draws the area, title, legend and pies): the value axis Excel would draw
// (charts.js valueAxis) with its labels and gridlines, the category axis, the axes' titles, a secondary value axis on
// the right for the series on it, and each chart group's series in its kind: columns and bars clustered or stacked
// within their group (a secondary group's columns over the primary's, as Excel draws them), lines and areas. The plot
// area's own fill and border under it all.
//
// Where things go, from Excel's PDFs of a 5 by 3 inch clustered column chart (excel-chart-format.ps1, chartdraw.test.js):
// a 10pt axis title takes 28.3px (2.12 times its size) off the plot; a horizontal one's baseline is its size above the
// space's bottom, an upright one (turned to read upwards) centred on the plot, its baseline 17.2px (the band less 0.83
// of the size) in from the chart's left; a secondary axis's labels start 11.2px right of the plot, which gives up their
// width and 5px; its title's baseline 0.72 of the band in from the chart's right.
import { valueAxis, autoScale } from '../ooxml/charts.js';
import { cellText } from '../ooxml/numfmt.js';
import { el, paint, seriesColour, fontFor, textAttrs, widthOf, LINE } from './chartpaint.js';

const PX = 4 / 3;
const ticksOf = (axis) => { const out = []; for (let v = axis.min, k = 0; v <= axis.max + axis.major * 1e-9 && k < 200; k++, v = axis.min + k * axis.major) out.push(Number(v.toPrecision(15))); return out; };
const bandOf = (f) => (f ? 2.12 * f.size : 0);

/** An axis's labels: its numbers in its format (100% stacked charts as percentages). */
function labeller(chart, own, secondary, date1904) {
  const first = chart.series.find((s) => !!s.secondary === secondary);
  const code = own?.linked ? first?.code || 'General' : own?.code || 'General';
  const percent = (chart.groups || []).some((g) => !!g.secondary === secondary && g.grouping === 'percentStacked');
  return (v) => (percent ? `${Math.round(v * 100)}%` : cellText({ t: 'n', v }, Infinity, code, date1904).trim());
}

export function drawAxes(svg, chart, theme, { top, right, bottom, width, date1904 }) {
  const horizontal = chart.kind === 'bar' && chart.dir === 'bar';
  const scatter = chart.kind === 'scatter';
  const vFont = fontFor(chart, chart.axis.font, 'axis', theme);
  const cFont = fontFor(chart, chart.catFont, 'axis', theme);
  const axis = valueAxis(chart);
  const ticks = ticksOf(axis);
  const label = labeller(chart, chart.axis, false, date1904);
  const two = chart.axis2 && chart.series.some((s) => s.secondary) ? chart.axis2 : null;
  const axis2 = two ? valueAxis(chart, true) : null;
  const ticks2 = axis2 ? ticksOf(axis2) : [];
  const label2 = labeller(chart, two, true, date1904);
  const v2Font = two ? fontFor(chart, two.font, 'axis', theme) : null;
  // (The titles by where they are drawn: a bar chart's value axis is the horizontal one.)
  const hTitle = horizontal ? chart.axis.title : chart.catTitle;
  const vTitle = horizontal ? chart.catTitle : chart.axis.title;
  const hF = hTitle && fontFor(chart, hTitle.font, 'axisTitle', theme);
  const vF = vTitle && fontFor(chart, vTitle.font, 'axisTitle', theme);
  const v2F = two?.title && fontFor(chart, two.title.font, 'axisTitle', theme);
  // (A scatter chart's x axis is a value axis too: Excel's automatic scale over the x values.)
  const xs = scatter ? chart.series.flatMap((s) => s.xs || []).filter((x) => typeof x === 'number') : [];
  const xAxis = scatter ? autoScale(xs.length ? Math.min(...xs) : 0, xs.length ? Math.max(...xs) : 0) : null;
  const xTicks = xAxis ? ticksOf(xAxis) : [];
  const cats = scatter ? xTicks.map((t) => cellText({ t: 'n', v: t }, Infinity, 'General', date1904).trim()) : chart.cats.length ? chart.cats : (chart.series[0]?.values || []).map((_, i) => String(i + 1));
  const n = cats.length || 1;
  const valueW = chart.axis.hidden ? 0 : Math.max(...ticks.map((t) => widthOf(label(t), vFont.size, vFont.bold)), 0);
  const catW = chart.catHidden ? 0 : Math.max(...cats.map((c) => widthOf(c ?? '', cFont.size, cFont.bold)), 0);
  const v2W = two && !two.hidden ? Math.max(...ticks2.map((t) => widthOf(label2(t), v2Font.size, v2Font.bold)), 0) : 0;
  const left = 12 + bandOf(vF) + (horizontal ? catW + 8 : valueW + 8);
  const rightEdge = right - (two && !two.hidden ? v2W + 5 : 0) - bandOf(v2F);
  const below = horizontal ? (chart.axis.hidden ? 0 : 8 + vFont.size) : (chart.catHidden ? 0 : 10 + cFont.size);
  const plotBottom = bottom - bandOf(hF) - below;
  const plot = { x: left, y: top + 6, w: Math.max(10, rightEdge - left), h: Math.max(10, plotBottom - top - 6) };
  const scale = (a) => (v) => (horizontal ? plot.x + ((v - a.min) / (a.max - a.min)) * plot.w : plot.y + plot.h - ((v - a.min) / (a.max - a.min)) * plot.h);
  const toV = scale(axis);
  const toV2 = axis2 ? scale(axis2) : toV;
  // The plot area's own fill and border, then gridlines and value labels.
  const pf = paint(chart.look?.plot?.fill, theme, 'none');
  const pl = paint(chart.look?.plot?.line, theme, 'none');
  if (pf !== 'none' || pl !== 'none') svg.append(el('rect', { x: plot.x, y: plot.y, width: plot.w, height: plot.h, fill: pf, stroke: pl }));
  const vText = textAttrs(vFont);
  for (const t of ticks) {
    const p = toV(t);
    if (chart.axis.gridlines) svg.append(horizontal ? el('line', { x1: p, x2: p, y1: plot.y, y2: plot.y + plot.h, stroke: LINE }) : el('line', { x1: plot.x, x2: plot.x + plot.w, y1: p, y2: p, stroke: LINE }));
    if (!chart.axis.hidden) svg.append(horizontal ? el('text', { x: p, y: plot.y + plot.h + 4 + vFont.size, 'text-anchor': 'middle', ...vText }, label(t)) : el('text', { x: plot.x - 6, y: p + vFont.size / 3, 'text-anchor': 'end', ...vText }, label(t)));
  }
  if (two && !two.hidden) for (const t of ticks2) svg.append(el('text', { x: plot.x + plot.w + 11.2, y: toV2(t) + v2Font.size / 3, ...textAttrs(v2Font) }, label2(t)));
  // The axes' titles.
  const cy = plot.y + plot.h / 2;
  if (hF) svg.append(el('text', { x: plot.x + plot.w / 2, y: bottom - hF.size, 'text-anchor': 'middle', ...textAttrs(hF) }, hTitle.text));
  if (vF) { const x = 12 + bandOf(vF) - 0.83 * vF.size; svg.append(el('text', { x, y: cy, 'text-anchor': 'middle', transform: `rotate(-90 ${x} ${cy})`, ...textAttrs(vF) }, vTitle.text)); }
  if (v2F) { const x = width - 0.72 * bandOf(v2F); svg.append(el('text', { x, y: cy, 'text-anchor': 'middle', transform: `rotate(-90 ${x} ${cy})`, ...textAttrs(v2F) }, two.title.text)); }
  const cText = textAttrs(cFont);
  if (scatter) {
    const toX = (x) => plot.x + ((x - xAxis.min) / (xAxis.max - xAxis.min)) * plot.w;
    xTicks.forEach((t, i) => {
      const p = toX(t);
      if (chart.xGridlines) svg.append(el('line', { x1: p, x2: p, y1: plot.y, y2: plot.y + plot.h, stroke: LINE }));
      if (!chart.catHidden) svg.append(el('text', { x: p, y: plot.y + plot.h + 5 + cFont.size, 'text-anchor': 'middle', ...cText }, cats[i]));
    });
    drawScatter(svg, chart, { toX, toV, theme });
    return;
  }
  // Category labels, and the category axis line at zero.
  const band = (horizontal ? plot.h : plot.w) / n;
  const zero = toV(Math.min(Math.max(0, axis.min), axis.max));
  svg.append(horizontal ? el('line', { x1: zero, x2: zero, y1: plot.y, y2: plot.y + plot.h, stroke: LINE }) : el('line', { x1: plot.x, x2: plot.x + plot.w, y1: zero, y2: zero, stroke: LINE }));
  if (!chart.catHidden) {
    cats.forEach((c, i) => {
      // (A bar chart's first category is at the bottom, as Excel draws it.)
      const mid = horizontal ? plot.y + plot.h - (i + 0.5) * band : plot.x + (i + 0.5) * band;
      svg.append(horizontal ? el('text', { x: plot.x - 6, y: mid + cFont.size / 3, 'text-anchor': 'end', ...cText }, c ?? '') : el('text', { x: mid, y: plot.y + plot.h + 5 + cFont.size, 'text-anchor': 'middle', ...cText }, c ?? ''));
    });
  }
  // Each group in the order the part has them (Excel's: areas, columns, lines; the secondary axis's after).
  const groups = chart.groups || [{ kind: chart.kind, grouping: chart.grouping, gapWidth: chart.gapWidth, overlap: chart.overlap }];
  const labels = fontFor(chart, chart.look?.fonts?.labels, 'labels', theme);
  groups.forEach((g, gi) => {
    const list = chart.series.filter((s) => (s.group ?? 0) === gi);
    const ctx = { plot, band, toV: g.secondary ? toV2 : toV, theme, horizontal, n, group: g, list, labels };
    if (g.kind === 'bar') drawBars(svg, chart, ctx);
    else drawLines(svg, chart, ctx);
  });
}

function drawBars(svg, chart, { plot, band, toV, theme, horizontal, n, group, list, labels }) {
  const stacked = group.grouping === 'stacked' || group.grouping === 'percentStacked';
  const k = stacked ? 1 : list.length;
  const overlap = stacked ? 100 : group.overlap;
  // (Excel's gap width and overlap: k bars of width w, each overlapping the next by overlap% of w, with gap% of w
  // between groups.)
  const w = band / (k - ((k - 1) * overlap) / 100 + group.gapWidth / 100);
  const totals = new Array(n).fill(0).map((_, i) => list.reduce((s, x) => s + Math.abs(typeof x.values[i] === 'number' ? x.values[i] : 0), 0) || 1);
  const up = new Array(n).fill(0);
  const down = new Array(n).fill(0);
  list.forEach((s, si) => {
    const colour = seriesColour(s, theme);
    const stroke = paint(s.look?.line, theme, 'none');
    for (let i = 0; i < n; i++) {
      let v = s.values[i];
      if (typeof v !== 'number') continue;
      if (group.grouping === 'percentStacked') v /= totals[i];
      let a;
      let b;
      if (stacked) { const base = v >= 0 ? up[i] : down[i]; a = base; b = base + v; if (v >= 0) up[i] = b; else down[i] = b; }
      else { a = 0; b = v; }
      const groupStart = (horizontal ? n - 1 - i : i) * band + (band - w * (k - ((k - 1) * overlap) / 100)) / 2;
      const offset = stacked ? 0 : si * w * (1 - overlap / 100);
      const p1 = toV(a);
      const p2 = toV(b);
      const along = (horizontal ? plot.y : plot.x) + groupStart + offset;
      const point = s.look?.points?.get(i);
      const fill = point ? paint(point.fill, theme, colour) : colour;
      const line = point?.line ? paint(point.line, theme, stroke) : stroke;
      svg.append(horizontal
        ? el('rect', { x: Math.min(p1, p2), y: along, width: Math.abs(p2 - p1), height: w, fill, stroke: line })
        : el('rect', { x: along, y: Math.min(p1, p2), width: w, height: Math.abs(p2 - p1), fill, stroke: line }));
      if (chart.labels) {
        const t = String(s.values[i]);
        const at = textAttrs(labels);
        svg.append(horizontal ? el('text', { x: Math.max(p1, p2) + 4, y: along + w / 2 + labels.size / 3, ...at }, t)
          : el('text', { x: along + w / 2, y: Math.min(p1, p2) - 4, 'text-anchor': 'middle', ...at }, t));
      }
    }
  });
}

function drawLines(svg, chart, { plot, band, toV, theme, n, group, list, labels }) {
  const stacked = group.grouping === 'stacked' || group.grouping === 'percentStacked';
  const sums = new Array(n).fill(0);
  const totals = new Array(n).fill(0).map((_, i) => list.reduce((s, x) => s + (typeof x.values[i] === 'number' ? x.values[i] : 0), 0) || 1);
  for (const s of list) {
    const colour = seriesColour(s, theme);
    const pts = [];
    for (let i = 0; i < n; i++) {
      let v = s.values[i];
      if (typeof v !== 'number') { pts.push(null); continue; }
      if (group.grouping === 'percentStacked') v /= totals[i];
      if (stacked) { sums[i] += v; v = sums[i]; }
      pts.push([plot.x + (i + 0.5) * band, toV(v), i]);
    }
    const runs = [];
    let run = [];
    for (const p of pts) { if (p) run.push(p); else if (run.length) { runs.push(run); run = []; } }
    if (run.length) runs.push(run);
    // (A line of Excel's default style is 2.25pt, 3px: the PDF's.)
    const width = (s.look?.width ?? 2.25) * PX;
    for (const r of runs) {
      const d = r.map(([x, y], j) => `${j ? 'L' : 'M'} ${x} ${y}`).join(' ');
      if (group.kind === 'area') svg.append(el('path', { d: `${d} L ${r[r.length - 1][0]} ${toV(0)} L ${r[0][0]} ${toV(0)} Z`, fill: colour, 'fill-opacity': 0.85 }));
      else svg.append(el('path', { d, fill: 'none', stroke: colour, 'stroke-width': width, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      if (group.kind === 'line' && s.marker && s.marker !== 'none') for (const [x, y] of r) svg.append(el('circle', { cx: x, cy: y, r: 3.5, fill: colour }));
      if (chart.labels) for (const [x, y, i] of r) svg.append(el('text', { x, y: y - 7, 'text-anchor': 'middle', ...textAttrs(labels) }, String(s.values[i])));
    }
  }
}

/** Scatter series: a marker at each (x, y), joined by straight lines (or, smoothed, by curves through them) when the
 *  series has a line. */
function drawScatter(svg, chart, { toX, toV, theme }) {
  for (const s of chart.series) {
    const colour = seriesColour(s, theme);
    const pts = [];
    s.values.forEach((v, i) => { if (typeof v === 'number' && typeof s.xs?.[i] === 'number') pts.push([toX(s.xs[i]), toV(v)]); });
    if (s.line && pts.length > 1) {
      let d = `M ${pts[0][0]} ${pts[0][1]}`;
      for (let i = 1; i < pts.length; i++) {
        if (!s.smooth) { d += ` L ${pts[i][0]} ${pts[i][1]}`; continue; }
        // (A Catmull-Rom curve through the points, drawn as cubic Béziers.)
        const [p0, p1, p2, p3] = [pts[i - 2] || pts[i - 1], pts[i - 1], pts[i], pts[i + 1] || pts[i]];
        d += ` C ${p1[0] + (p2[0] - p0[0]) / 6} ${p1[1] + (p2[1] - p0[1]) / 6} ${p2[0] - (p3[0] - p1[0]) / 6} ${p2[1] - (p3[1] - p1[1]) / 6} ${p2[0]} ${p2[1]}`;
      }
      svg.append(el('path', { d, fill: 'none', stroke: colour, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
    }
    if (s.marker !== 'none') for (const [x, y] of pts) svg.append(el('circle', { cx: x, cy: y, r: 3.5, fill: colour }));
  }
}
