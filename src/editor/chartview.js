// A chart (ooxml/charts.js) drawn as SVG, in the look of Excel's default chart style: a white area with a light grey
// border, the title in grey 14pt, grey 9pt axis labels, light grey gridlines, series in the theme's accent colours,
// the legend under or beside the plot. Column, bar (clustered, stacked, 100%), line, area, pie and doughnut. The value
// axis is Excel's (charts.js valueAxis). Pictures only: what a screen reader hears is chartFigure's table.
import { valueAxis } from '../ooxml/charts.js';
import { resolveColor } from '../ooxml/theme.js';
import { cellText } from '../ooxml/numfmt.js';

const SVG = 'http://www.w3.org/2000/svg';
const GREY = '#595959';
const LINE = '#D9D9D9';
const FONT = 'Aptos Narrow, Calibri, Carlito, sans-serif';

/** The colour of series (or, varying by point, point) i: its own, else the theme's accents in turn, darker past six. */
function colorFor(chart, theme, i, own) {
  if (own?.rgb) return `#${own.rgb}`;
  if (own?.scheme && SCHEME[own.scheme] != null) return resolveColor({ theme: SCHEME[own.scheme], tint: own.lumMod ? own.lumMod - 1 : 0 }, theme) || '#4472C4';
  const accent = resolveColor({ theme: 4 + (i % 6) }, theme) || ['#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E'][i % 6];
  return i < 6 ? accent : shade(accent, 0.6);
}
// (The theme's colours by their scheme names, in the theme's own order.)
const SCHEME = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9 };
const shade = (hex, f) => `#${[1, 3, 5].map((k) => Math.round(parseInt(hex.slice(k, k + 2), 16) * f).toString(16).padStart(2, '0')).join('')}`;

const el = (name, attrs = {}, text = null) => {
  const x = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) x.setAttribute(k, String(v));
  if (text != null) x.textContent = text;
  return x;
};
/** Text width, roughly, for laying out labels (9pt digits are about 6px). */
const widthOf = (s, size) => String(s).length * size * 0.52;

/** A chart as an <svg> of `width` by `height` pixels. */
export function chartSvg(chart, { width, height, theme, date1904 = false }) {
  const svg = el('svg', { xmlns: SVG, width, height, viewBox: `0 0 ${width} ${height}`, 'font-family': FONT });
  svg.append(el('rect', { x: 0.5, y: 0.5, width: width - 1, height: height - 1, fill: '#FFFFFF', stroke: LINE }));
  let top = 10;
  if (chart.title) {
    svg.append(el('text', { x: width / 2, y: top + 18, 'text-anchor': 'middle', 'font-size': 18.6, fill: GREY }, chart.title));
    top += 30;
  }
  const pie = chart.kind === 'pie' || chart.kind === 'doughnut';
  const entries = pie ? chart.cats.map((c, i) => ({ name: c ?? String(i + 1), color: colorFor(chart, theme, i, null) }))
    : chart.series.map((s, i) => ({ name: s.name, color: colorFor(chart, theme, i, s.color) }));
  let right = width - 12;
  let bottom = height - 10;
  if (chart.legend && entries.length) {
    if (chart.legend === 'b' || chart.legend === 't') {
      const widths = entries.map((e) => 14 + widthOf(e.name, 12) + 12);
      let x = (width - widths.reduce((a, b) => a + b, 0)) / 2;
      const y = height - 14;
      entries.forEach((e, i) => {
        svg.append(el('rect', { x, y: y - 8, width: 7, height: 7, fill: e.color }));
        svg.append(el('text', { x: x + 11, y, 'font-size': 12, fill: GREY }, e.name));
        x += widths[i];
      });
      bottom = height - 30;
    } else {
      const w = Math.max(...entries.map((e) => widthOf(e.name, 12))) + 22;
      const x0 = width - w - 6;
      let y = (top + height) / 2 - (entries.length * 18) / 2 + 12;
      for (const e of entries) {
        svg.append(el('rect', { x: x0, y: y - 8, width: 7, height: 7, fill: e.color }));
        svg.append(el('text', { x: x0 + 11, y, 'font-size': 12, fill: GREY }, e.name));
        y += 18;
      }
      right = x0 - 10;
    }
  }
  if (pie) drawPie(svg, chart, entries, { left: 12, top, right, bottom });
  else drawAxes(svg, chart, theme, { top, right, bottom, date1904 });
  return svg;
}

function drawPie(svg, chart, entries, box) {
  const values = (chart.series[0]?.values || []).map((v) => (typeof v === 'number' && v > 0 ? v : 0));
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const cx = (box.left + box.right) / 2;
  const cy = (box.top + box.bottom) / 2;
  const r = Math.max(10, Math.min(box.right - box.left, box.bottom - box.top) / 2 - 6);
  const hole = chart.kind === 'doughnut' ? r * chart.holeSize / 100 : 0;
  let a = 0;
  values.forEach((v, i) => {
    if (!v) return;
    const b = a + (v / total) * Math.PI * 2;
    const pt = (rad, ang) => [cx + rad * Math.sin(ang), cy - rad * Math.cos(ang)];
    const big = b - a > Math.PI ? 1 : 0;
    const [x1, y1] = pt(r, a);
    const [x2, y2] = pt(r, b);
    let d;
    if (v === total) d = `M ${cx} ${cy - r} A ${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r} Z`;
    else if (hole) { const [x3, y3] = pt(hole, b); const [x4, y4] = pt(hole, a); d = `M ${x1} ${y1} A ${r} ${r} 0 ${big} 1 ${x2} ${y2} L ${x3} ${y3} A ${hole} ${hole} 0 ${big} 0 ${x4} ${y4} Z`; }
    else d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${big} 1 ${x2} ${y2} Z`;
    svg.append(el('path', { d, fill: entries[i]?.color || '#999', stroke: '#FFFFFF', 'stroke-width': 1 }));
    a = b;
  });
}

function drawAxes(svg, chart, theme, { top, right, bottom, date1904 }) {
  const axis = valueAxis(chart);
  const code = chart.axis.linked ? chart.series[0]?.code || 'General' : chart.axis.code || 'General';
  const label = (v) => (chart.grouping === 'percentStacked' ? `${Math.round(v * 100)}%` : cellText({ t: 'n', v }, Infinity, code, date1904).trim());
  const ticks = [];
  for (let v = axis.min, k = 0; v <= axis.max + axis.major * 1e-9 && k < 200; k++, v = axis.min + k * axis.major) ticks.push(Number(v.toPrecision(15)));
  const horizontal = chart.kind === 'bar' && chart.dir === 'bar';
  const cats = chart.cats.length ? chart.cats : (chart.series[0]?.values || []).map((_, i) => String(i + 1));
  const n = cats.length || 1;
  const valueW = chart.axis.hidden ? 0 : Math.max(...ticks.map((t) => widthOf(label(t), 12)), 0);
  const catW = chart.catHidden ? 0 : Math.max(...cats.map((c) => widthOf(c ?? '', 12)), 0);
  const left = 12 + (horizontal ? catW + 8 : valueW + 8);
  const plotBottom = bottom - (horizontal ? (chart.axis.hidden ? 0 : 20) : (chart.catHidden ? 0 : 22));
  const plot = { x: left, y: top + 6, w: Math.max(10, right - left), h: Math.max(10, plotBottom - top - 6) };
  const toV = (v) => (horizontal ? plot.x + ((v - axis.min) / (axis.max - axis.min)) * plot.w : plot.y + plot.h - ((v - axis.min) / (axis.max - axis.min)) * plot.h);
  // Gridlines and value labels.
  for (const t of ticks) {
    const p = toV(t);
    if (chart.axis.gridlines) svg.append(horizontal ? el('line', { x1: p, x2: p, y1: plot.y, y2: plot.y + plot.h, stroke: LINE }) : el('line', { x1: plot.x, x2: plot.x + plot.w, y1: p, y2: p, stroke: LINE }));
    if (!chart.axis.hidden) svg.append(horizontal ? el('text', { x: p, y: plot.y + plot.h + 16, 'text-anchor': 'middle', 'font-size': 12, fill: GREY }, label(t)) : el('text', { x: plot.x - 6, y: p + 4, 'text-anchor': 'end', 'font-size': 12, fill: GREY }, label(t)));
  }
  // Category labels, and the category axis line at zero.
  const band = (horizontal ? plot.h : plot.w) / n;
  const zero = toV(Math.min(Math.max(0, axis.min), axis.max));
  svg.append(horizontal ? el('line', { x1: zero, x2: zero, y1: plot.y, y2: plot.y + plot.h, stroke: LINE }) : el('line', { x1: plot.x, x2: plot.x + plot.w, y1: zero, y2: zero, stroke: LINE }));
  if (!chart.catHidden) {
    cats.forEach((c, i) => {
      // (A bar chart's first category is at the bottom, as Excel draws it.)
      const mid = horizontal ? plot.y + plot.h - (i + 0.5) * band : plot.x + (i + 0.5) * band;
      svg.append(horizontal ? el('text', { x: plot.x - 6, y: mid + 4, 'text-anchor': 'end', 'font-size': 12, fill: GREY }, c ?? '') : el('text', { x: mid, y: plot.y + plot.h + 17, 'text-anchor': 'middle', 'font-size': 12, fill: GREY }, c ?? ''));
    });
  }
  const colors = chart.series.map((s, i) => colorFor(chart, theme, i, s.color));
  if (chart.kind === 'bar') drawBars(svg, chart, { plot, band, toV, colors, horizontal, n });
  else drawLines(svg, chart, { plot, band, toV, colors, n });
}

function drawBars(svg, chart, { plot, band, toV, colors, horizontal, n }) {
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
  const k = stacked ? 1 : chart.series.length;
  const overlap = stacked ? 100 : chart.overlap;
  // (Excel's gap width and overlap: k bars of width w, each overlapping the next by overlap% of w, with gap% of w
  // between groups.)
  const w = band / (k - ((k - 1) * overlap) / 100 + chart.gapWidth / 100);
  const totals = new Array(n).fill(0).map((_, i) => chart.series.reduce((s, x) => s + Math.abs(typeof x.values[i] === 'number' ? x.values[i] : 0), 0) || 1);
  const up = new Array(n).fill(0);
  const down = new Array(n).fill(0);
  chart.series.forEach((s, si) => {
    for (let i = 0; i < n; i++) {
      let v = s.values[i];
      if (typeof v !== 'number') continue;
      if (chart.grouping === 'percentStacked') v /= totals[i];
      let a;
      let b;
      if (stacked) { const base = v >= 0 ? up[i] : down[i]; a = base; b = base + v; if (v >= 0) up[i] = b; else down[i] = b; }
      else { a = 0; b = v; }
      const groupStart = (horizontal ? n - 1 - i : i) * band + (band - w * (k - ((k - 1) * overlap) / 100)) / 2;
      const offset = stacked ? 0 : si * w * (1 - overlap / 100);
      const p1 = toV(a);
      const p2 = toV(b);
      const along = (horizontal ? plot.y : plot.x) + groupStart + offset;
      svg.append(horizontal
        ? el('rect', { x: Math.min(p1, p2), y: along, width: Math.abs(p2 - p1), height: w, fill: colors[si] })
        : el('rect', { x: along, y: Math.min(p1, p2), width: w, height: Math.abs(p2 - p1), fill: colors[si] }));
    }
  });
}

function drawLines(svg, chart, { plot, band, toV, colors, n }) {
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
  const sums = new Array(n).fill(0);
  const totals = new Array(n).fill(0).map((_, i) => chart.series.reduce((s, x) => s + (typeof x.values[i] === 'number' ? x.values[i] : 0), 0) || 1);
  chart.series.forEach((s, si) => {
    const pts = [];
    for (let i = 0; i < n; i++) {
      let v = s.values[i];
      if (typeof v !== 'number') { pts.push(null); continue; }
      if (chart.grouping === 'percentStacked') v /= totals[i];
      if (stacked) { sums[i] += v; v = sums[i]; }
      pts.push([plot.x + (i + 0.5) * band, toV(v)]);
    }
    const runs = [];
    let run = [];
    for (const p of pts) { if (p) run.push(p); else if (run.length) { runs.push(run); run = []; } }
    if (run.length) runs.push(run);
    for (const r of runs) {
      const d = r.map(([x, y], j) => `${j ? 'L' : 'M'} ${x} ${y}`).join(' ');
      if (chart.kind === 'area') svg.append(el('path', { d: `${d} L ${r[r.length - 1][0]} ${toV(0)} L ${r[0][0]} ${toV(0)} Z`, fill: colors[si], 'fill-opacity': 0.85 }));
      else svg.append(el('path', { d, fill: 'none', stroke: colors[si], 'stroke-width': 3, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      if (chart.kind === 'line' && s.marker && s.marker !== 'none') for (const [x, y] of r) svg.append(el('circle', { cx: x, cy: y, r: 3.5, fill: colors[si] }));
    }
  });
}

const KIND_NAMES = { bar: 'Column chart', line: 'Line chart', area: 'Area chart', pie: 'Pie chart', doughnut: 'Doughnut chart' };
/** What a screen reader is given for a chart: a figure with its title and kind, and its data as a table. */
export function chartFigure(chart) {
  const fig = document.createElement('figure');
  const cap = document.createElement('figcaption');
  const kind = chart.kind === 'bar' && chart.dir === 'bar' ? 'Bar chart' : KIND_NAMES[chart.kind] || 'Chart';
  cap.textContent = `${kind}${chart.title ? `: ${chart.title}` : `: ${chart.name}`}`;
  const table = document.createElement('table');
  const head = document.createElement('tr');
  for (const t of ['', ...chart.series.map((s) => s.name)]) head.append(Object.assign(document.createElement('th'), { textContent: t, scope: 'col' }));
  table.append(head);
  chart.cats.forEach((c, i) => {
    const row = document.createElement('tr');
    row.append(Object.assign(document.createElement('th'), { textContent: c ?? String(i + 1), scope: 'row' }));
    for (const s of chart.series) row.append(Object.assign(document.createElement('td'), { textContent: s.values[i] ?? '' }));
    table.append(row);
  });
  fig.append(cap, table);
  return fig;
}

/** Draw `st.charts` (charts.js readCharts) that reach into window `win` into `st.chartLayer`, each where its anchor
 *  puts it over the cells; a chart's picture is made again only when its size changes. */
export function drawCharts(st, win) {
  const layer = st.chartLayer;
  if (!layer) return;
  st.chartCache ??= new Map();
  const out = [];
  for (const ch of st.charts || []) {
    if (!ch.kind || ch.to.r < win.r0 || ch.from.r > win.r1 || ch.to.c < win.c0 || ch.from.c > win.c1) continue;
    const x = st.cols.start(ch.from.c) + ch.from.dx;
    const y = st.rows.start(ch.from.r) + ch.from.dy;
    const w = Math.round(st.cols.start(ch.to.c) + ch.to.dx - x);
    const h = Math.round(st.rows.start(ch.to.r) + ch.to.dy - y);
    if (w < 20 || h < 20) continue;
    const key = `${w}x${h}`;
    let box = st.chartCache.get(ch);
    if (!box || box.key !== key) {
      const div = document.createElement('div');
      div.className = 'oxe-chart';
      div.append(chartSvg(ch, { width: w, height: h, theme: st.theme, date1904: st.date1904 }));
      box = { key, div };
      st.chartCache.set(ch, box);
    }
    box.div.style.setProperty('--oxe-x', `${x}px`);
    box.div.style.setProperty('--oxe-y', `${y}px`);
    out.push(box.div);
  }
  layer.replaceChildren(...out);
}
