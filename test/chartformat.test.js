import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { insertChart } from '../src/ooxml/chartnew.js';
import { setChartKind, setChartLabels, chartParts } from '../src/ooxml/chartedit.js';
import { setAxisTitle, setChartFont, setChartFill } from '../src/ooxml/chartformat.js';
import { setSeriesType, setSeriesAxis } from '../src/ooxml/chartcombo.js';
import { parseRange } from '../src/ooxml/refs.js';

// A chart's formatting as Excel writes it: test/chartformat-excel.json holds Excel's chart part after each change of
// a clustered column chart of A1:C5 (tools/excel-compare/excel-chart-format.ps1: Axis.HasTitle and AxisTitle.Text,
// Format.Fill and Format.Line .ForeColor.RGB of a series, a point, the chart and plot areas, the Font of the title,
// axes, legend, data labels and an axis title, Series.ChartType and Series.AxisGroup). Each is made here on the same
// chart, inserted here (chartnew.test.js: Excel's own, text for text), and compared text for text.
export const excel = JSON.parse(fs.readFileSync('test/chartformat-excel.json', 'utf8'));
const DATA = [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90], ['West', 60, 45]];
export function withChart(name) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.renameSheet(0, name);
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  DATA.forEach((row, r) => row.forEach((v, c) => wb.setCell(0, { r, c }, v)));
  insertChart(wb, 0, parseRange('A1:C5'), 'columnClustered', { at: { r: 1, c: 5 } });
  return wb;
}
// (Axis ids and unique ids are Excel's own afresh each time; a secondary axis's ids are compared by the part they play.)
const bare = (xml, name) => {
  const ids = new Map();
  const id = (v) => { if (!ids.has(v)) ids.set(v, `#${ids.size}`); return ids.get(v); };
  return xml.replace(/(axId|crossAx) val="(\d+)"/g, (m, a, v) => `${a} ${id(v)}`).replace(/uniqueId val="\{(\d{8})-[^}]*\}"/g, 'uid $1').replace(new RegExp(`${name}!`, 'g'), 'S!');
};
const font = { name: 'Arial', size: 11, bold: true, color: '1F4E79' };
export const EDITS = {
  catTitle: (wb) => setAxisTitle(wb, 0, 0, 'h', 'Region'),
  valTitle: (wb) => setAxisTitle(wb, 0, 0, 'v', 'Amount'),
  catTitleAuto: (wb) => setAxisTitle(wb, 0, 0, 'h', 'auto'),
  titlesOff: (wb) => { setAxisTitle(wb, 0, 0, 'h', 'X'); setAxisTitle(wb, 0, 0, 'v', 'X'); setAxisTitle(wb, 0, 0, 'h', null); setAxisTitle(wb, 0, 0, 'v', null); },
  seriesFill: (wb) => setChartFill(wb, 0, 0, { series: 0 }, { fill: 'C00000' }),
  seriesLine: (wb) => setChartFill(wb, 0, 0, { series: 0 }, { line: '7030A0' }),
  pointFill: (wb) => setChartFill(wb, 0, 0, { series: 0, point: 1 }, { fill: '00B050' }),
  chartFill: (wb) => setChartFill(wb, 0, 0, 'chart', { fill: 'FFF2CC' }),
  chartLine: (wb) => setChartFill(wb, 0, 0, 'chart', { line: '1F4E79' }),
  plotFill: (wb) => setChartFill(wb, 0, 0, 'plot', { fill: 'DDEBF7' }),
  plotLine: (wb) => setChartFill(wb, 0, 0, 'plot', { line: '808080' }),
  titleFont1: (wb) => setChartFont(wb, 0, 0, 'title', { name: 'Arial', size: 18, bold: true, color: 'C00000' }),
  titleFontName1: (wb) => setChartFont(wb, 0, 0, 'title', { name: 'Arial' }),
  catFontSize1: (wb) => setChartFont(wb, 0, 0, 'h', { size: 12 }),
  catFont: (wb) => setChartFont(wb, 0, 0, 'h', font),
  valFont: (wb) => setChartFont(wb, 0, 0, 'v', font),
  legendFont1: (wb) => setChartFont(wb, 0, 0, 'legend', font),
  labelFont1: (wb) => { setChartLabels(wb, 0, 0, true); setChartFont(wb, 0, 0, 'labels', font, 0); },
  catTitleFont1: (wb) => { setAxisTitle(wb, 0, 0, 'h', 'Region'); setChartFont(wb, 0, 0, 'hTitle', { name: 'Arial', size: 12, bold: true, color: 'C00000' }); },
  lineColor: (wb) => { setChartKind(wb, 0, 0, 'line'); setChartFill(wb, 0, 0, { series: 0 }, { line: 'C00000' }); },
  pieSlice: (wb) => { setChartKind(wb, 0, 0, 'pie'); setChartFill(wb, 0, 0, { series: 0, point: 1 }, { fill: '00B050' }); },
  comboLine: (wb) => setSeriesType(wb, 0, 0, 1, 'line'),
  comboLineMarkers: (wb) => setSeriesType(wb, 0, 0, 1, 'lineMarkers'),
  comboOwnColor: (wb) => { setChartFill(wb, 0, 0, { series: 1 }, { fill: 'C00000' }); setSeriesType(wb, 0, 0, 1, 'line'); },
  secondaryTitle: (wb) => { setSeriesAxis(wb, 0, 0, 1, true); setAxisTitle(wb, 0, 0, 'v2', 'Cost'); },
  comboArea: (wb) => setSeriesType(wb, 0, 0, 1, 'area'),
  secondary: (wb) => setSeriesAxis(wb, 0, 0, 1, true),
  comboSecondary: (wb) => { setSeriesType(wb, 0, 0, 1, 'line'); setSeriesAxis(wb, 0, 0, 1, true); },
  secondaryBack: (wb) => { setSeriesAxis(wb, 0, 0, 1, true); setSeriesAxis(wb, 0, 0, 1, false); },
  comboBack: (wb) => { setSeriesType(wb, 0, 0, 1, 'line'); setSeriesType(wb, 0, 0, 1, 'column'); },
  comboSecondaryBack: (wb) => { setSeriesType(wb, 0, 0, 1, 'line'); setSeriesAxis(wb, 0, 0, 1, true); setSeriesAxis(wb, 0, 0, 1, false); setSeriesType(wb, 0, 0, 1, 'column'); },
};

test(`each format change written as Excel writes it (${Object.keys(EDITS).length} changes)`, () => {
  const wrong = [];
  for (const [name, edit] of Object.entries(EDITS)) {
    const wb = withChart(name);
    edit(wb);
    const got = bare(wb.pkg.text(chartParts(wb, 0)[0].path), name);
    const want = bare(excel.charts[name], name);
    if (got !== want) {
      let k = 0;
      while (got[k] === want[k]) k++;
      wrong.push(`${name} at ${k}: Excel …${want.slice(k - 80, k + 150)}… here …${got.slice(k - 80, k + 150)}…`);
    }
  }
  // (The unedited chart, as a check of the comparison itself.)
  assert.equal(bare(withChart('none').pkg.text(chartParts(withChart('none'), 0)[0].path), 'none'), bare(excel.charts.none, 'none'));
  assert.deepEqual(wrong, []);
});
