import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { insertChart } from '../src/ooxml/chartnew.js';
import { setChartTitle, setChartLegend, setChartLabels, setChartKind, setChartAnchor, deleteChart, chartParts } from '../src/ooxml/chartedit.js';
import { readCharts } from '../src/ooxml/charts.js';
import { parseRange } from '../src/ooxml/refs.js';

// A chart's edits as Excel writes them: test/chartedit-excel.json holds Excel's chart part after each edit of a
// clustered column chart (tools/excel-compare/excel-chart-edit.ps1). Each is made here on the same chart, inserted
// here (chartnew.test.js shows it is Excel's own, text for text), and compared text for text.
const excel = JSON.parse(fs.readFileSync('test/chartedit-excel.json', 'utf8')).charts;
const DATA = [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90], ['West', 60, 45]];
function withChart(name) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.renameSheet(0, name);
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  DATA.forEach((row, r) => row.forEach((v, c) => wb.setCell(0, { r, c }, v)));
  insertChart(wb, 0, parseRange('A1:C5'), 'columnClustered', { at: { r: 1, c: 5 } });
  return wb;
}
const bare = (xml, name) => xml.replace(/axId val="\d+"/g, 'axId').replace(/crossAx val="\d+"/g, 'crossAx').replace(/uniqueId val="\{[^}]*\}"/g, 'uid').replace(new RegExp(`${name}!`, 'g'), 'S!');
const EDITS = {
  titleText: (wb) => setChartTitle(wb, 0, 0, 'Sales & cost'),
  noTitle: (wb) => setChartTitle(wb, 0, 0, null),
  noLegend: (wb) => setChartLegend(wb, 0, 0, null),
  legendRight: (wb) => setChartLegend(wb, 0, 0, 'r'),
  legendTop: (wb) => setChartLegend(wb, 0, 0, 't'),
  legendLeft: (wb) => setChartLegend(wb, 0, 0, 'l'),
  dataLabels: (wb) => setChartLabels(wb, 0, 0, true),
  toLine: (wb) => setChartKind(wb, 0, 0, 'line'),
  toLineMarkers: (wb) => setChartKind(wb, 0, 0, 'lineMarkers'),
  toBar: (wb) => setChartKind(wb, 0, 0, 'barClustered'),
  toColumnStacked: (wb) => setChartKind(wb, 0, 0, 'columnStacked'),
  toPie: (wb) => setChartKind(wb, 0, 0, 'pie'),
  toDoughnut: (wb) => setChartKind(wb, 0, 0, 'doughnut'),
  toArea: (wb) => setChartKind(wb, 0, 0, 'area'),
  toScatter: (wb) => setChartKind(wb, 0, 0, 'scatter'),
};

test('each edit of a chart written as Excel writes it (15 edits)', () => {
  const wrong = [];
  for (const [name, edit] of Object.entries(EDITS)) {
    const wb = withChart(name);
    edit(wb);
    const got = bare(wb.pkg.text(chartParts(wb, 0)[0].path), name);
    const want = bare(excel[name], name);
    if (got !== want) {
      let k = 0;
      while (got[k] === want[k]) k++;
      wrong.push(`${name} at ${k}: Excel …${want.slice(k - 50, k + 150)}… here …${got.slice(k - 50, k + 150)}…`);
    }
  }
  // (The unedited chart, as a check of the comparison itself.)
  assert.equal(bare(wb0().pkg.text(chartParts(wb0(), 0)[0].path), 'none'), bare(excel.none, 'none'));
  assert.deepEqual(wrong, []);
});
const wb0 = (() => { let wb; return () => (wb ??= withChart('none')); })();

test('read back after each edit: the title, the legend, the kind; data labels off again; the title back', () => {
  const wb = withChart('S');
  setChartTitle(wb, 0, 0, 'Sales & cost');
  setChartLegend(wb, 0, 0, 'r');
  setChartKind(wb, 0, 0, 'line');
  let c = readCharts(wb, 0)[0];
  assert.deepEqual([c.title, c.legend, c.kind, c.series.length], ['Sales & cost', 'r', 'line', 2]);
  setChartLabels(wb, 0, 0, true);
  setChartLabels(wb, 0, 0, false);
  assert.doesNotMatch(wb.pkg.text(chartParts(wb, 0)[0].path), /<c:showVal val="1"\/>/);
  setChartTitle(wb, 0, 0, null);
  assert.equal(readCharts(wb, 0)[0].title, null);
  setChartTitle(wb, 0, 0, 'Back');
  c = readCharts(wb, 0)[0];
  assert.equal(c.title, 'Back');
  const again = Workbook.fromBytes(wb.toBytes());
  assert.equal(readCharts(again, 0)[0].title, 'Back');
});

test('a chart moved, sized and deleted: its anchor, then its part and relationship gone', () => {
  const wb = withChart('S');
  insertChart(wb, 0, parseRange('A1:C5'), 'pie', { at: { r: 20, c: 1 } });
  setChartAnchor(wb, 0, 0, { col: 2, colOff: 0, row: 3, rowOff: 0 }, { col: 9, colOff: 9525, row: 18, rowOff: 0 });
  const [a, b] = readCharts(wb, 0);
  assert.deepEqual([a.from.c, a.from.r, a.to.c, a.to.r, a.to.dx], [2, 3, 9, 18, 1]);
  assert.equal(b.kind, 'pie');
  const gone = chartParts(wb, 0)[0].path;
  deleteChart(wb, 0, 0);
  assert.deepEqual(readCharts(wb, 0).map((c) => c.kind), ['pie']);
  assert.equal(wb.pkg.text(gone), null);
  assert.ok(Workbook.fromBytes(wb.toBytes()));
});
