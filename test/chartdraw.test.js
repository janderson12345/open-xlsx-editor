import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { chartParts } from '../src/ooxml/chartedit.js';
import { readCharts } from '../src/ooxml/charts.js';
import { excel, withChart } from './chartformat-cases.js';

// Formatted charts drawn here against Excel's PDF of the same charts (excel-chart-format.ps1 exported the workbook of
// test/chartformat-excel.json; tools/excel-compare/shapes.cjs and cells.cjs took its fills, lines and text, kept in
// test/chartdraw-excel.json as offsets from each chart's top left in pixels). Each chart part Excel saved is read here
// and drawn at the size Excel printed it, and what the change does to the drawing compared with what it did in
// Excel's: the plot's edges, where the titles and the secondary axis's labels go, and the colours.

/** Just enough of a DOM for chartSvg: elements with attributes, children and text. */
class Node {
  constructor(name) { this.name = name; this.attrs = {}; this.kids = []; this.textContent = ''; }
  setAttribute(k, v) { this.attrs[k] = v; }
  append(...n) { this.kids.push(...n); }
  all() { return [this, ...this.kids.flatMap((k) => k.all())]; }
}
globalThis.document ??= { createElementNS: (ns, name) => new Node(name) };
const { chartSvg } = await import('../src/editor/chartview.js');
const pdf = JSON.parse(fs.readFileSync('test/chartdraw-excel.json', 'utf8'));

function drawn(name) {
  const wb = withChart(name);
  wb.pkg.setText(chartParts(wb, 0)[0].path, excel.charts[name]);
  const chart = readCharts(wb, 0)[0];
  const svg = chartSvg(chart, { width: pdf.frame.w, height: pdf.frame.h, theme: wb.theme });
  const nodes = svg.all();
  const n = (k, a) => Number(k.attrs[a]);
  const grid = nodes.filter((k) => k.name === 'line' && k.attrs.stroke === '#D9D9D9' && n(k, 'y1') === n(k, 'y2') && n(k, 'x2') - n(k, 'x1') > 100);
  const text = (t) => nodes.find((k) => k.name === 'text' && k.textContent === t);
  return {
    plot: { left: Math.min(...grid.map((g) => n(g, 'x1'))), right: Math.max(...grid.map((g) => n(g, 'x2'))), top: Math.min(...grid.map((g) => n(g, 'y1'))), bottom: Math.max(...grid.map((g) => n(g, 'y1'))) },
    text, n, nodes,
    fills: nodes.filter((k) => k.name === 'rect').map((k) => String(k.attrs.fill).toLowerCase()),
    strokes: nodes.filter((k) => k.name === 'rect' || k.name === 'path').map((k) => String(k.attrs.stroke).toLowerCase()),
  };
}

test('axis titles take the room from the plot Excel\'s take, and sit where Excel puts them', () => {
  const none = drawn('none');
  const e = pdf.charts;
  const cat = drawn('catTitle');
  const val = drawn('valTitle');
  // (Excel: the plot's bottom 28.2px higher for a horizontal title, its left 28.3px further in for an upright one.)
  const near = (got, want, what, tol = 1.5) => assert.ok(Math.abs(got - want) <= tol, `${what}: Excel ${want.toFixed(1)}, here ${got.toFixed(1)}`);
  near(none.plot.bottom - cat.plot.bottom, e.none.plot.bottom - e.catTitle.plot.bottom, 'the plot\'s bottom with a horizontal title');
  near(val.plot.left - none.plot.left, e.valTitle.plot.left - e.none.plot.left, 'the plot\'s left with an upright title');
  // (The horizontal title centred under the plot, its baseline below the labels as Excel's is.)
  const t = cat.text('Region');
  near(cat.n(t, 'x'), (cat.plot.left + cat.plot.right) / 2, 'the horizontal title centred on the plot', 0.5);
  near(cat.n(t, 'y') - cat.plot.bottom, e.catTitle.title.y - e.catTitle.plot.bottom, 'the horizontal title below the plot', 3);
  const v = val.text('Amount');
  assert.match(v.attrs.transform, /^rotate\(-90 /);
  near(val.n(v, 'x'), e.valTitle.title.x, 'the upright title\'s baseline from the chart\'s left', 3);
  near(val.n(v, 'y'), (val.plot.top + val.plot.bottom) / 2, 'the upright title centred on the plot', 0.5);
});

test('the secondary axis on the right, its own scale, its columns over the primary ones, a line on it', () => {
  const none = drawn('none');
  const e = pdf.charts;
  const near = (got, want, what, tol = 3) => assert.ok(Math.abs(got - want) <= tol, `${what}: Excel ${want.toFixed(1)}, here ${got.toFixed(1)}`);
  const two = drawn('secondary');
  near(none.plot.right - two.plot.right, e.none.plot.right - e.secondary.plot.right, 'the plot\'s right edge given up to the labels');
  // (Excel's secondary labels: 0 to 100 by 10, left-aligned 11.2px right of the plot.)
  const labels = two.nodes.filter((k) => k.name === 'text' && Number(k.attrs.x) > two.plot.right).map((k) => k.textContent);
  assert.deepEqual(labels, ['0', '10', '20', '30', '40', '50', '60', '70', '80', '90', '100']);
  near(two.n(two.nodes.find((k) => k.name === 'text' && k.textContent === '100' && Number(k.attrs.x) > two.plot.right), 'x') - two.plot.right, e.secondary.labels.x - e.secondary.plot.right, 'the secondary labels from the plot', 0.5);
  // (Each group clustered on its own: one column a category in each, at the same place, over each other.)
  const bars = two.nodes.filter((k) => k.name === 'rect' && ['#156082', '#e97132'].includes(String(k.attrs.fill).toLowerCase()) && Number(k.attrs.height) > 20);
  const [a, b] = [bars.filter((k) => String(k.attrs.fill).toLowerCase() === '#156082'), bars.filter((k) => String(k.attrs.fill).toLowerCase() === '#e97132')];
  assert.deepEqual(a.map((k) => two.n(k, 'x').toFixed(1)), b.map((k) => two.n(k, 'x').toFixed(1)));
  near(two.n(a[0], 'width') / (two.plot.right - two.plot.left), e.secondary.barWidth / (e.secondary.plot.right - e.secondary.plot.left), 'a column\'s share of the plot', 0.005);
  // (A column of 140 on the secondary axis of 0 to 100 is as high as the plot; Cost's 80 at 0.8 of it.)
  near((two.plot.bottom - two.n(b[0], 'y')) / (two.plot.bottom - two.plot.top), e.secondary.costShare, 'the secondary column\'s height', 0.01);
  const combo = drawn('comboSecondary');
  const line = combo.nodes.find((k) => k.name === 'path' && String(k.attrs.stroke).toLowerCase() === '#e97132');
  assert.equal(Number(line.attrs['stroke-width']), 3);
  const ys = [...line.attrs.d.matchAll(/[ML] ([\d.]+) ([\d.]+)/g)].map((m) => (combo.plot.bottom - Number(m[2])) / (combo.plot.bottom - combo.plot.top));
  assert.deepEqual(ys.map((y) => y.toFixed(2)), e.comboSecondary.lineShares.map((y) => y.toFixed(2)));
  near(none.plot.right - drawn('secondaryTitle').plot.right, e.none.plot.right - e.secondaryTitle.plot.right, 'the plot\'s right edge with a secondary title');
});

test('the chart\'s own fills, lines and fonts drawn in Excel\'s colours', () => {
  const e = pdf.charts;
  assert.ok(drawn('chartFill').fills.includes(e.chartFill.fill), 'the chart area\'s fill');
  assert.ok(drawn('seriesFill').fills.includes(e.seriesFill.fill), 'a series\' fill');
  assert.ok(drawn('pointFill').fills.includes(e.pointFill.fill), 'a point\'s fill');
  assert.ok(drawn('plotLine').strokes.includes(e.plotLine.line), 'the plot area\'s border');
  assert.ok(drawn('chartLine').strokes.includes(e.chartLine.line), 'the chart area\'s border');
  assert.ok(drawn('plotFill').fills.includes(e.plotFill.fill), 'the plot area\'s fill');
  const title = drawn('titleFont1').text('Chart Title');
  assert.deepEqual([title.attrs['font-size'], title.attrs['font-weight'], title.attrs.fill, title.attrs['font-family'].split(',')[0]], ['24', 'bold', '#C00000', 'Arial']);
  const cat = drawn('catFont').text('North');
  assert.deepEqual([cat.attrs['font-size'], cat.attrs.fill], ['14.67', '#1F4E79']);
});
