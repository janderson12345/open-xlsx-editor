// The chart format cases (chartformat.test.js, chartlook.test.js, chartdraw.test.js): Excel's measured parts, the
// same chart inserted here, and each change made here as Excel made it.
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { insertChart } from '../src/ooxml/chartnew.js';
import { setChartKind, setChartLabels, setChartTitle } from '../src/ooxml/chartedit.js';
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
export const bare = (xml, name) => {
  const ids = new Map();
  const id = (v) => { if (!ids.has(v)) ids.set(v, `#${ids.size}`); return ids.get(v); };
  return xml.replace(/(axId|crossAx) val="(\d+)"/g, (m, a, v) => `${a} ${id(v)}`).replace(/uniqueId val="\{(\d{8})-[^}]*\}"/g, 'uid $1').replace(new RegExp(`${name}!`, 'g'), 'S!');
};
const font = { name: 'Arial', size: 11, bold: true, color: '1F4E79' };
export const EDITS = {
  catTitle: (wb, si = 0) => setAxisTitle(wb, si, 0, 'h', 'Region'),
  valTitle: (wb, si = 0) => setAxisTitle(wb, si, 0, 'v', 'Amount'),
  catTitleAuto: (wb, si = 0) => setAxisTitle(wb, si, 0, 'h', 'auto'),
  titlesOff: (wb, si = 0) => { setAxisTitle(wb, si, 0, 'h', 'X'); setAxisTitle(wb, si, 0, 'v', 'X'); setAxisTitle(wb, si, 0, 'h', null); setAxisTitle(wb, si, 0, 'v', null); },
  seriesFill: (wb, si = 0) => setChartFill(wb, si, 0, { series: 0 }, { fill: 'C00000' }),
  seriesLine: (wb, si = 0) => setChartFill(wb, si, 0, { series: 0 }, { line: '7030A0' }),
  pointFill: (wb, si = 0) => setChartFill(wb, si, 0, { series: 0, point: 1 }, { fill: '00B050' }),
  chartFill: (wb, si = 0) => setChartFill(wb, si, 0, 'chart', { fill: 'FFF2CC' }),
  chartLine: (wb, si = 0) => setChartFill(wb, si, 0, 'chart', { line: '1F4E79' }),
  plotFill: (wb, si = 0) => setChartFill(wb, si, 0, 'plot', { fill: 'DDEBF7' }),
  plotLine: (wb, si = 0) => setChartFill(wb, si, 0, 'plot', { line: '808080' }),
  titleFont1: (wb, si = 0) => setChartFont(wb, si, 0, 'title', { name: 'Arial', size: 18, bold: true, color: 'C00000' }),
  titleFontName1: (wb, si = 0) => setChartFont(wb, si, 0, 'title', { name: 'Arial' }),
  catFontSize1: (wb, si = 0) => setChartFont(wb, si, 0, 'h', { size: 12 }),
  catFont: (wb, si = 0) => setChartFont(wb, si, 0, 'h', font),
  valFont: (wb, si = 0) => setChartFont(wb, si, 0, 'v', font),
  legendFont1: (wb, si = 0) => setChartFont(wb, si, 0, 'legend', font),
  labelFont1: (wb, si = 0) => { setChartLabels(wb, si, 0, true); setChartFont(wb, si, 0, 'labels', font, 0); },
  catTitleFont1: (wb, si = 0) => { setAxisTitle(wb, si, 0, 'h', 'Region'); setChartFont(wb, si, 0, 'hTitle', { name: 'Arial', size: 12, bold: true, color: 'C00000' }); },
  lineColor: (wb, si = 0) => { setChartKind(wb, si, 0, 'line'); setChartFill(wb, si, 0, { series: 0 }, { line: 'C00000' }); },
  pieSlice: (wb, si = 0) => { setChartKind(wb, si, 0, 'pie'); setChartFill(wb, si, 0, { series: 0, point: 1 }, { fill: '00B050' }); },
  comboLine: (wb, si = 0) => setSeriesType(wb, si, 0, 1, 'line'),
  comboLineMarkers: (wb, si = 0) => setSeriesType(wb, si, 0, 1, 'lineMarkers'),
  titleFontThenText: (wb, si = 0) => { setChartFont(wb, si, 0, 'title', { name: 'Arial', size: 18, bold: true, color: 'C00000' }); setChartTitle(wb, si, 0, 'Sales'); },
  titleTextFontText: (wb, si = 0) => { setChartTitle(wb, si, 0, 'Sales'); setChartFont(wb, si, 0, 'title', { name: 'Arial', size: 18, bold: true, color: 'C00000' }); setChartTitle(wb, si, 0, 'Sales & cost'); },
  comboOwnColor: (wb, si = 0) => { setChartFill(wb, si, 0, { series: 1 }, { fill: 'C00000' }); setSeriesType(wb, si, 0, 1, 'line'); },
  secondaryTitle: (wb, si = 0) => { setSeriesAxis(wb, si, 0, 1, true); setAxisTitle(wb, si, 0, 'v2', 'Cost'); },
  comboArea: (wb, si = 0) => setSeriesType(wb, si, 0, 1, 'area'),
  secondary: (wb, si = 0) => setSeriesAxis(wb, si, 0, 1, true),
  comboSecondary: (wb, si = 0) => { setSeriesType(wb, si, 0, 1, 'line'); setSeriesAxis(wb, si, 0, 1, true); },
  secondaryBack: (wb, si = 0) => { setSeriesAxis(wb, si, 0, 1, true); setSeriesAxis(wb, si, 0, 1, false); },
  comboBack: (wb, si = 0) => { setSeriesType(wb, si, 0, 1, 'line'); setSeriesType(wb, si, 0, 1, 'column'); },
  comboSecondaryBack: (wb, si = 0) => { setSeriesType(wb, si, 0, 1, 'line'); setSeriesAxis(wb, si, 0, 1, true); setSeriesAxis(wb, si, 0, 1, false); setSeriesType(wb, si, 0, 1, 'column'); },
};

