import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { spawnSync } from 'node:child_process';
import { Workbook } from '../src/ooxml/workbook.js';
import { chartSeries, insertChart } from '../src/ooxml/chartnew.js';
import { readCharts } from '../src/ooxml/charts.js';
import { parseRange } from '../src/ooxml/refs.js';

// Insert > Charts as Excel does it: test/chartnew-excel.json holds what Excel 365 wrote for each kind of chart
// (Shapes.AddChart2 in its default style on A1:C5 of a small table) and the series it made from ten shapes of
// selection. tools/excel-compare/excel-chart-insert.ps1 made the first.
const excel = JSON.parse(fs.readFileSync('test/chartnew-excel.json', 'utf8'));
const open = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
/** A workbook whose first sheet, renamed `name`, has only `cells` (rows of values from A1). */
function bookWith(cells, name = 'Sheet1') {
  const wb = open();
  wb.renameSheet(0, name);
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  cells.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setCell(0, { r, c }, v); }));
  // (Default column widths, as Excel's new workbook has: basic.xlsx widens some.)
  sheet.cols.length = 0;
  return wb;
}
const formulaOf = (s, i) => `=SERIES(${s.name?.f ?? ''},${s.cat?.f ?? ''},${s.val.f},${i + 1})`;

test('the series of a selection, as Excel takes them (ten shapes of selection)', () => {
  const wrong = [];
  for (const c of excel.series) {
    const wb = bookWith(c.cells);
    const got = chartSeries(wb, 0, parseRange(c.range)).series.map(formulaOf);
    if (JSON.stringify(got) !== JSON.stringify(c.formulas)) wrong.push(`${c.name}: Excel ${c.formulas.join(' ; ')}, here ${got.join(' ; ')}`);
  }
  assert.deepEqual(wrong, []);
});

/** Excel's own made-up ids, and the drawing's creation id, out of the way. */
const bare = (xml) => xml.replace(/axId val="\d+"/g, 'axId').replace(/crossAx val="\d+"/g, 'crossAx').replace(/\{0000000(\d)-[0-9A-F-]{27}\}/g, '{$1}')
  .replace(/<a:extLst><a:ext uri="\{FF2B5EF4[\s\S]*?<\/a:extLst>/, '').replace(/(<xdr:cNvPr id="\d+" name="[^"]+")>(<\/xdr:cNvPr>)/, '$1/>');

test('each kind of chart written as Excel writes it: its chart part and its anchor beside the data (14 kinds)', () => {
  const data = [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90], ['West', 60, 45]];
  const wrong = [];
  for (const want of excel.charts) {
    const wb = bookWith(data, want.kind);
    // (Excel's AddChart2 put it 240pt across and 15pt down: F2.)
    const made = insertChart(wb, 0, parseRange('A1:C5'), want.kind, { at: { r: 1, c: 5 } });
    const files = unzipSync(wb.toBytes());
    const chart = strFromU8(files[made.chart]);
    const drawing = strFromU8(files[made.drawing]);
    if (bare(chart) !== bare(want.chart)) {
      const a = bare(chart);
      const b = bare(want.chart);
      let k = 0;
      while (a[k] === b[k]) k++;
      wrong.push(`${want.kind} chart differs at ${k}: Excel …${b.slice(k - 60, k + 120)}… here …${a.slice(k - 60, k + 120)}…`);
    }
    if (bare(drawing) !== bare(want.drawing)) wrong.push(`${want.kind} drawing: Excel ${bare(want.drawing)} here ${bare(drawing)}`);
  }
  assert.deepEqual(wrong, []);
});

test('an inserted chart is read and drawn here, and the saved workbook passes the checker', () => {
  const wb = bookWith([['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90]]);
  insertChart(wb, 0, parseRange('A1:C4'), 'columnClustered');
  insertChart(wb, 0, parseRange('A1:C4'), 'scatter');
  const charts = readCharts(wb, 0);
  assert.deepEqual(charts.map((c) => [c.kind, c.series.length, c.series[0].name, c.series[0].values.join(',')]), [['bar', 2, 'Sales', '120,95,140'], ['scatter', 2, 'Sales', '120,95,140']]);
  const again = Workbook.fromBytes(wb.toBytes());
  assert.equal(readCharts(again, 0).length, 2);
  fs.mkdirSync('test/out/charts', { recursive: true });
  fs.writeFileSync('test/out/charts/chartnew.xlsx', wb.toBytes());
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/charts/chartnew.xlsx'], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
});

test('Excel reads the charts written here as the kinds and series it made (test/chartnew-readback-excel.json)', () => {
  // (Every kind on a sheet of its own, saved to test/out/charts/chartnew-all.xlsx: excel-chart-readback.ps1 opened it
  // in Excel and read each chart's type and first series' formula back.)
  const data = [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90], ['West', 60, 45]];
  const wb = bookWith(data, excel.charts[0].kind);
  for (let k = 1; k < excel.charts.length; k++) {
    wb.addSheet(k, excel.charts[k].kind);
    data.forEach((row, r) => row.forEach((v, c) => wb.setCell(k, { r, c }, v)));
  }
  excel.charts.forEach((c, k) => insertChart(wb, k, parseRange('A1:C5'), c.kind, { at: { r: 1, c: 5 } }));
  fs.mkdirSync('test/out/charts', { recursive: true });
  fs.writeFileSync('test/out/charts/chartnew-all.xlsx', wb.toBytes());
  const back = JSON.parse(fs.readFileSync('test/chartnew-readback-excel.json', 'utf8'));
  const titles = excel.charts.map((c, k) => readCharts(wb, k)[0].title);
  const wrong = [];
  for (const [k, c] of excel.charts.entries()) {
    const got = back.charts.find((x) => x.sheet === c.kind);
    if (got?.title !== titles[k]) wrong.push(`${c.kind}: Excel's title ${got?.title}, here ${titles[k]}`);
    if (got?.type !== c.type || got?.series !== 2 || got?.formula1 !== `=SERIES(${c.kind}!$B$1,${c.kind}!$A$2:$A$5,${c.kind}!$B$2:$B$5,1)`) wrong.push(`${c.kind}: Excel made type ${c.type}, read back ${JSON.stringify(got)}`);
  }
  assert.deepEqual(wrong, []);
});
