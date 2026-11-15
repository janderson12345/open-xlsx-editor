import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { insertChart } from '../src/ooxml/chartnew.js';
import { setChartTitle, setChartLabels, setChartKind, chartParts } from '../src/ooxml/chartedit.js';
import { chartData, setChartData, addChartSeries, editChartSeries, removeChartSeries, moveChartSeries, setChartBlanks } from '../src/ooxml/chartdata.js';
import { readCharts } from '../src/ooxml/charts.js';
import { parseRange } from '../src/ooxml/refs.js';

// A chart's data changed as Excel changes it: test/chartdata-excel.json holds Excel's chart part after each change of
// a clustered column chart of A1:C5 (tools/excel-compare/excel-chart-data.ps1). Each is made here on the same chart,
// inserted here (chartnew.test.js shows it is Excel's own), and compared text for text.
const excel = JSON.parse(fs.readFileSync('test/chartdata-excel.json', 'utf8'));
const DATA = [['Region', 'Sales', 'Cost', 'Profit'], ['North', 120, 80, 40], ['South', 95, 70, 25], ['East', 140, 90, 50], ['West', 60, 45, 15]];
function withChart(name) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.renameSheet(0, name);
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  DATA.forEach((row, r) => row.forEach((v, c) => wb.setCell(0, { r, c }, v)));
  insertChart(wb, 0, parseRange('A1:C5'), 'columnClustered', { at: { r: 1, c: 6 } });
  return wb;
}
const part = (wb) => chartParts(wb, 0)[0].path;
const bare = (xml, name) => xml.replace(/axId val="\d+"/g, 'axId').replace(/crossAx val="\d+"/g, 'crossAx')
  .replace(/uniqueId val="\{(\d{8})-[^}]*\}"/g, 'uid $1').split(`${name}!`).join('S!');
/** Series 2 given a red fill of its own, as Excel writes Format.Fill.ForeColor.RGB = 255. */
const red = (wb) => {
  const t = wb.pkg.text(part(wb));
  const sers = [...t.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)];
  const two = sers[1][0];
  wb.pkg.setText(part(wb), t.replace(two, two.replace('<a:schemeClr val="accent2"/>', '<a:srgbClr val="FF0000"/>')));
};
const flip = (wb) => { const d = chartData(wb, 0, 0); return setChartData(wb, 0, 0, d.range, !d.byRows); };
const EDITS = {
  sourceWider: (wb, s) => setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`),
  sourceSmaller: (wb, s) => setChartData(wb, 0, 0, `=${s}!$A$1:$B$4`),
  sourceRows: (wb, s) => setChartData(wb, 0, 0, `=${s}!$A$1:$C$5`, true),
  sourceCols: (wb, s) => setChartData(wb, 0, 0, `=${s}!$A$1:$D$3`, false),
  sourceShort: (wb, s) => setChartData(wb, 0, 0, `=${s}!$A$1:$D$3`),
  plotByRows: (wb) => flip(wb),
  rowsThenSource: (wb, s) => { flip(wb); setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); },
  newSeries: (wb, s) => addChartSeries(wb, 0, 0, { name: `='${s}'!$D$1`, val: `='${s}'!$D$2:$D$5`, cat: `='${s}'!$A$2:$A$5` }),
  newSeriesBare: (wb, s) => addChartSeries(wb, 0, 0, { val: `='${s}'!$D$2:$D$5` }),
  nameText: (wb) => editChartSeries(wb, 0, 0, 0, { name: 'Takings' }),
  nameRef: (wb, s) => editChartSeries(wb, 0, 0, 0, { name: `='${s}'!$D$1` }),
  nameEmpty: (wb) => editChartSeries(wb, 0, 0, 0, { name: '' }),
  valuesRef: (wb, s) => editChartSeries(wb, 0, 0, 1, { val: `='${s}'!$D$2:$D$5` }),
  valuesRelative: (wb, s) => editChartSeries(wb, 0, 0, 1, { val: `=${s}!D2:D5` }),
  xvaluesRef: (wb, s) => editChartSeries(wb, 0, 0, 0, { cat: `='${s}'!$D$2:$D$5` }),
  xvaluesEmpty: (wb) => editChartSeries(wb, 0, 0, 0, { cat: '' }),
  formula: (wb, s) => editChartSeries(wb, 0, 0, 1, { name: `=${s}!$D$1`, cat: `=${s}!$A$2:$A$5`, val: `=${s}!$D$2:$D$5` }),
  plotOrder: (wb) => moveChartSeries(wb, 0, 0, 1, -1),
  deleteFirst: (wb) => removeChartSeries(wb, 0, 0, 0),
  deleteThenAdd: (wb, s) => { removeChartSeries(wb, 0, 0, 0); addChartSeries(wb, 0, 0, { name: `=${s}!$D$1`, val: `=${s}!$D$2:$D$5` }); },
  threeMoveDown: (wb, s) => { setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); moveChartSeries(wb, 0, 0, 0, 1); },
  blanksZero: (wb) => setChartBlanks(wb, 0, 0, { blanks: 'zero' }),
  blanksSpan: (wb) => setChartBlanks(wb, 0, 0, { blanks: 'span' }),
  hiddenShown: (wb) => setChartBlanks(wb, 0, 0, { hidden: true }),
  labelsThenSource: (wb, s) => { setChartLabels(wb, 0, 0, true); setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); },
  labelsThenNew: (wb, s) => { setChartLabels(wb, 0, 0, true); addChartSeries(wb, 0, 0, { val: `=${s}!$D$2:$D$5` }); },
  labelsThenRows: (wb) => { setChartLabels(wb, 0, 0, true); flip(wb); },
  titleThenSource: (wb, s) => { setChartTitle(wb, 0, 0, 'Own'); setChartData(wb, 0, 0, `=${s}!$A$1:$B$5`); },
  colorThenSource: (wb, s) => { red(wb); setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); },
  colorThenRows: (wb) => { red(wb); flip(wb); },
  lineNewSeries: (wb, s) => { setChartKind(wb, 0, 0, 'lineMarkers'); addChartSeries(wb, 0, 0, { val: `=${s}!$D$2:$D$5` }); },
  lineSource: (wb, s) => { setChartKind(wb, 0, 0, 'lineMarkers'); setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); },
  pieSource: (wb, s) => { setChartKind(wb, 0, 0, 'pie'); setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); },
  scatterSource: (wb, s) => { setChartKind(wb, 0, 0, 'scatter'); setChartData(wb, 0, 0, `=${s}!$A$1:$D$5`); },
};

test(`each change of a chart's data written as Excel writes it (${Object.keys(EDITS).length} changes)`, () => {
  const wrong = [];
  for (const [name, edit] of Object.entries(EDITS)) {
    const wb = withChart(name);
    edit(wb, name);
    const got = bare(wb.pkg.text(part(wb)), name);
    const want = bare(excel.charts[name], name);
    if (got !== want) {
      let k = 0;
      while (got[k] === want[k]) k++;
      wrong.push(`${name} at ${k}: Excel …${want.slice(k - 60, k + 140)}… here …${got.slice(k - 60, k + 140)}…`);
    }
  }
  assert.equal(bare(wb0(), 'none'), bare(excel.charts.none, 'none'));
  assert.deepEqual(wrong, []);
});
const wb0 = () => { const wb = withChart('none'); return wb.pkg.text(part(wb)); };

test('Select Data reads the range, the way round and each series as Excel gives them', () => {
  for (const name of ['none', 'sourceRows', 'nameText', 'deleteThenAdd']) {
    const wb = withChart(name);
    if (EDITS[name]) EDITS[name](wb, name);
    const d = chartData(wb, 0, 0);
    // (Excel's own formulas: =SERIES(name, categories, values, order).)
    const formulas = d.series.map((s, k) => `=SERIES(${s.name == null ? '' : 'text' in s.name ? `"${s.name.text}"` : s.name.f},${s.cat ?? ''},${s.val},${k + 1})`);
    assert.deepEqual(formulas, excel.formulas[name], name);
  }
  const wb = withChart('S');
  assert.deepEqual({ range: chartData(wb, 0, 0).range, byRows: chartData(wb, 0, 0).byRows }, { range: '=S!$A$1:$C$5', byRows: false });
  flip(wb);
  assert.deepEqual({ range: chartData(wb, 0, 0).range, byRows: chartData(wb, 0, 0).byRows }, { range: '=S!$A$1:$C$5', byRows: true });
  // (A series of its own: no one range gives these series.)
  addChartSeries(wb, 0, 0, { val: '=S!$D$2:$D$3' });
  assert.equal(chartData(wb, 0, 0).range, null);
});

test('a chart whose data changed is drawn from its new series; Excel refuses what is not a reference', () => {
  const wb = withChart('S');
  setChartData(wb, 0, 0, 'A1:D5');
  assert.deepEqual(readCharts(wb, 0)[0].series.map((s) => s.name), ['Sales', 'Cost', 'Profit']);
  moveChartSeries(wb, 0, 0, 2, -1);
  assert.deepEqual(readCharts(wb, 0)[0].series.map((s) => s.name), ['Sales', 'Profit', 'Cost']);
  assert.throws(() => editChartSeries(wb, 0, 0, 0, { val: '=Nowhere!$A$1' }), /not valid/);
  assert.throws(() => setChartData(wb, 0, 0, 'A1'), /numbers/);
  removeChartSeries(wb, 0, 0, 0);
  removeChartSeries(wb, 0, 0, 0);
  assert.throws(() => removeChartSeries(wb, 0, 0, 0), /at least one/);
});
