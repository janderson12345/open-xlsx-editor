// A chart (ooxml/charts.js) drawn as SVG, in its own look where its part sets one (fills, lines, fonts: chartpaint.js)
// and Excel's default chart style for the rest: a white area with a light grey border, the title in grey 14pt, the
// legend under or beside the plot, series in the theme's accent colours. Column, bar (clustered, stacked, 100%), line,
// area, pie, doughnut and scatter, and combo charts of columns, lines and areas with a secondary axis; the axes and
// plot are chartplot.js's. Pictures only: what a screen reader hears is chartFigure's table.
import { el, paint, seriesColour, accent, fontFor, textAttrs, widthOf, SVG, LINE } from './chartpaint.js';
import { drawAxes } from './chartplot.js';
import { maxOf } from '../ooxml/listmath.js';

/** A chart as an <svg> of `width` by `height` pixels. */
export function chartSvg(chart, { width, height, theme, date1904 = false }) {
  const tf = fontFor(chart, chart.look?.fonts?.title, 'title', theme);
  const svg = el('svg', { xmlns: SVG, width, height, viewBox: `0 0 ${width} ${height}`, 'font-family': tf.family });
  const area = chart.look?.area || {};
  svg.append(el('rect', { x: 0.5, y: 0.5, width: width - 1, height: height - 1, fill: paint(area.fill, theme, '#FFFFFF'), stroke: paint(area.line, theme, LINE) }));
  let top = 10;
  if (chart.title) {
    // (The default 14pt title, 18.7px: its baseline 28px down, the plot 40px; a bigger or smaller one in proportion.)
    svg.append(el('text', { x: width / 2, y: top + tf.size * 0.97, 'text-anchor': 'middle', ...textAttrs(tf) }, chart.title));
    top += tf.size * 1.61;
  }
  const pie = chart.kind === 'pie' || chart.kind === 'doughnut';
  const own = pie ? chart.series[0]?.look?.points : null;
  const entries = pie ? chart.cats.map((c, i) => ({ name: c ?? String(i + 1), color: paint(own?.get(i)?.fill, theme, accent(theme, i)) }))
    : chart.series.map((s) => ({ name: s.name, color: seriesColour(s, theme), line: s.type === 'line' }));
  const lf = fontFor(chart, chart.look?.fonts?.legend, 'legend', theme);
  const lt = textAttrs(lf);
  // (A line series' key is a short line, 25.6px in Excel's PDF; the others a square.)
  const key = (e, x, y) => (e.line ? el('line', { x1: x - 9, x2: x + 16, y1: y - 4, y2: y - 4, stroke: e.color, 'stroke-width': 3 }) : el('rect', { x, y: y - 8, width: 7, height: 7, fill: e.color }));
  const keyW = (e) => (e.line ? 27 : 11);
  let right = width - 12;
  let bottom = height - 10;
  if (chart.legend && entries.length) {
    if (chart.legend === 'b' || chart.legend === 't') {
      const widths = entries.map((e) => 3 + keyW(e) + widthOf(e.name, lf.size, lf.bold) + 12);
      let x = (width - widths.reduce((p, q) => p + q, 0)) / 2;
      const y = height - 14;
      entries.forEach((e, i) => {
        const kx = x + keyW(e) - 11;
        svg.append(key(e, kx, y), el('text', { x: kx + 11 + (e.line ? 16 : 0), y, ...lt }, e.name));
        x += widths[i];
      });
      bottom = height - 30;
    } else {
      const w = maxOf(entries.map((e) => widthOf(e.name, lf.size, lf.bold) + keyW(e))) + 11;
      const x0 = width - w - 6;
      let y = (top + height) / 2 - (entries.length * 18) / 2 + 12;
      for (const e of entries) {
        const kx = x0 + keyW(e) - 11;
        svg.append(key(e, kx, y), el('text', { x: kx + 11 + (e.line ? 16 : 0), y, ...lt }, e.name));
        y += 18;
      }
      right = x0 - 10;
    }
  }
  if (pie) drawPie(svg, chart, entries, { left: 12, top, right, bottom, theme });
  else drawAxes(svg, chart, theme, { top, right, bottom, width, date1904 });
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
    const ln = chart.series[0]?.look?.points?.get(i)?.line;
    svg.append(el('path', { d, fill: entries[i]?.color || '#999', stroke: paint(ln, box.theme, '#FFFFFF'), 'stroke-width': 1 }));
    if (chart.labels) {
      const [lx, ly] = pt(hole ? (r + hole) / 2 : r * 0.65, (a + b) / 2);
      svg.append(el('text', { x: lx, y: ly + 4, 'text-anchor': 'middle', ...textAttrs(fontFor(chart, chart.look?.fonts?.labels, 'labels', box.theme)) }, String(chart.series[0].values[i])));
    }
    a = b;
  });
}

const KIND_NAMES = { bar: 'Column chart', line: 'Line chart', area: 'Area chart', pie: 'Pie chart', doughnut: 'Doughnut chart', scatter: 'Scatter chart' };
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

/** The boxes of `st.charts` (charts.js readCharts) that reach into window `win`, each where its anchor puts it over
 *  the cells: [{ ord (its place in the drawing's order), el }]. A chart's picture is made again only when its size
 *  changes (`fresh`: a box of its own, for a frozen pane). drawview.js drawDrawings draws them with the objects. */
export function chartBoxes(st, win, fresh = false) {
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
    let box = fresh ? null : st.chartCache.get(ch);
    if (!box || box.key !== key) {
      const div = document.createElement('div');
      div.className = 'oxe-chart';
      div.dataset.oxeChart = String(ch.index ?? st.charts.indexOf(ch));
      div.dataset.oxeOrd = String(ch.ord ?? -1 - (ch.index ?? 0));
      div.append(chartSvg(ch, { width: w, height: h, theme: st.theme, date1904: st.date1904 }));
      box = { key, div };
      if (!fresh) st.chartCache.set(ch, box);
    }
    box.div.style.setProperty('--oxe-x', `${x}px`);
    box.div.style.setProperty('--oxe-y', `${y}px`);
    out.push({ ord: ch.ord ?? -1 - (ch.index ?? 0), el: box.div });
  }
  return out;
}
