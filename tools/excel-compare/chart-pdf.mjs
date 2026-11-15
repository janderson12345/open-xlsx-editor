// Excel's PDF of the formatted charts (excel-chart-format.ps1: one sheet a page, in the order of its cases) taken down
// to what chartdraw.test.js compares: each chart's plot edges (its gridlines), its titles' and secondary labels'
// places, its columns, a line's points and the colours, as offsets in pixels from the chart's top left.
// usage: node tools/excel-compare/cells.cjs <pdfDir> <textDir>; node tools/excel-compare/shapes.cjs <pdfDir> <shapeDir>;
//        node tools/excel-compare/chart-pdf.mjs <excel.json> <text.json> <shapes.json> <out.json>
import fs from 'node:fs';

const [casesFile, textFile, shapesFile, out] = process.argv.slice(2);
const names = Object.keys(JSON.parse(fs.readFileSync(casesFile, 'utf8')).charts);
const text = JSON.parse(fs.readFileSync(textFile, 'utf8'));
const shapes = JSON.parse(fs.readFileSync(shapesFile, 'utf8'));
const r1 = (v) => Math.round(v * 10) / 10;

function chartOf(page) {
  const s = shapes.filter((x) => x[1] === page);
  // (The chart's area: the biggest fill on the page; its border, the lines on its edges.)
  const area = s.filter((x) => x[0] === 'fill').sort((a, b) => b[5] * b[6] - a[5] * a[6])[0];
  const [fx, fy, fw, fh] = area.slice(3);
  const onEdge = (x) => [x[3], x[5]].some((v) => Math.abs(v - fx) < 0.5 || Math.abs(v - fx - fw) < 0.5);
  const grid = s.filter((x) => x[0] === 'line' && x[2] === '#d9d9d9' && Math.abs(x[4] - x[6]) < 0.1 && Math.abs(x[5] - x[3]) > 100 && !onEdge(x));
  const plot = { left: r1(Math.min(...grid.map((g) => Math.min(g[3], g[5]))) - fx), right: r1(Math.max(...grid.map((g) => Math.max(g[3], g[5]))) - fx),
    top: r1(Math.min(...grid.map((g) => g[4])) - fy), bottom: r1(Math.max(...grid.map((g) => g[4])) - fy) };
  const words = text.filter((x) => x[1] === page && x[2] > fx && x[3] > fy).map(([k, , x, y, w]) => ({ k, x: r1(x - fx), y: r1(y - fy), w: r1(w) }));
  const bars = s.filter((x) => x[0] === 'fill' && x[6] > 20 && x !== area).map((x) => ({ c: x[2], x: r1(x[3] - fx), y: r1(x[4] - fy), w: r1(x[5]), h: r1(x[6]) }));
  const border = s.filter((x) => x[0] === 'line' && onEdge(x) && x[2] !== '#000000').map((x) => x[2]);
  return { fx, fy, fw, fh, area, plot, words, bars, border, s };
}

const charts = {};
const at = (name) => chartOf(names.indexOf(name) + 1);
const none = at('none');
charts.none = { plot: none.plot };
charts.catTitle = (() => { const c = at('catTitle'); const t = c.words.find((w) => w.k === 'Region' && w.y > c.plot.bottom); return { plot: c.plot, title: { x: r1(t.x + t.w / 2), y: t.y } }; })();
charts.valTitle = (() => { const c = at('valTitle'); const t = c.words.find((w) => w.k === 'Amount'); return { plot: c.plot, title: { x: t.x, y: t.y } }; })();
charts.secondary = (() => {
  const c = at('secondary');
  const label = c.words.find((w) => w.k === '100' && w.x > c.plot.right);
  const blue = c.bars.find((b) => b.c === '#156082');
  const orange = c.bars.find((b) => b.c === '#e97132');
  return { plot: c.plot, labels: { x: label.x }, barWidth: blue.w, costShare: Math.round((c.plot.bottom - orange.y) / (c.plot.bottom - c.plot.top) * 1000) / 1000 };
})();
charts.comboSecondary = (() => {
  const c = at('comboSecondary');
  const segs = c.s.filter((x) => x[0] === 'line' && x[2] === '#e97132' && x[7] === 3 && Math.abs(x[4] - x[6]) > 0.1);
  const ys = [segs[0][4] - c.fy, ...segs.map((x) => x[6] - c.fy)];
  return { plot: c.plot, lineShares: ys.map((y) => Math.round((c.plot.bottom - y) / (c.plot.bottom - c.plot.top) * 1000) / 1000) };
})();
charts.secondaryTitle = { plot: at('secondaryTitle').plot };
charts.chartFill = { fill: at('chartFill').area[2] };
charts.chartLine = { line: at('chartLine').border[0] };
charts.seriesFill = { fill: at('seriesFill').bars[0].c };
charts.pointFill = { fill: at('pointFill').bars.find((b) => !['#156082', '#e97132'].includes(b.c)).c };
charts.plotFill = { fill: (() => { const c = at('plotFill'); return c.s.find((x) => x[0] === 'fill' && x !== c.area && x[5] > 300 && x[6] > 100)[2]; })() };
charts.plotLine = { line: (() => { const c = at('plotLine'); return c.s.find((x) => x[0] === 'line' && !['#d9d9d9', '#000000'].includes(x[2]))[2]; })() };
fs.writeFileSync(out, `${JSON.stringify({ source: 'Excel 365 PDF of excel-chart-format.ps1\'s workbook, through cells.cjs and shapes.cjs: offsets in px from each chart\'s top left', frame: { w: r1(none.fw), h: r1(none.fh) }, charts }, null, 1)}\n`);
console.log('done');
