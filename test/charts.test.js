import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { autoScale } from '../src/ooxml/charts.js';

// Excel's automatic value axis, each case as Excel scaled it (a chart over the numbers, Axes(xlValue) MinimumScale,
// MaximumScale and MajorUnit over COM: tools/excel-compare/excel-axes.ps1 on axes-cases.json, from axes-cases.mjs,
// giving test/axes-excel.json).
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/axes-cases.json', 'utf8'));
const excel = JSON.parse(fs.readFileSync('test/axes-excel.json', 'utf8').replace(/^\uFEFF/, '')).cases;

test('a chart\'s automatic value axis, as Excel scales it', () => {
  const wrong = [];
  for (const c of cases) {
    const want = excel.find((x) => x.name === c.name);
    const got = autoScale(Math.min(...c.values), Math.max(...c.values));
    const same = (a, b) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));
    if (!same(got.min, want.min) || !same(got.max, want.max) || !same(got.major, want.major)) wrong.push(`${c.name}: Excel ${want.min} to ${want.max} by ${want.major}, here ${got.min} to ${got.max} by ${got.major}`);
  }
  assert.deepEqual(wrong, []);
  assert.equal(cases.length, 34);
});

test('charts.xlsx\'s four charts read as Excel reads them', async () => {
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const { readCharts, valueAxis } = await import('../src/ooxml/charts.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/charts.xlsx')));
  const charts = readCharts(wb, 0);
  // (Excel over COM: Chart 1 type 51 (clustered columns), 2 series, "Sales by month", axis 0 to 200 by 20, at 240,15pt,
  // 384 by 225pt; Chart 2 type 4 (lines), 3 series, no title, 0 to 200 by 20; Chart 3 type 5 (pie), 1 series, "North";
  // Chart 4 type 58 (stacked bars), 3 series, "Sales, stacked", 0 to 450 by 50.)
  assert.deepEqual(charts.map((c) => [c.kind, c.dir, c.grouping, c.series.length, c.title]), [
    ['bar', 'col', 'clustered', 2, 'Sales by month'], ['line', 'col', 'standard', 3, null], ['pie', 'col', 'clustered', 1, 'North'], ['bar', 'bar', 'stacked', 3, 'Sales, stacked']]);
  assert.deepEqual([0, 1, 3].map((i) => valueAxis(charts[i])), [{ min: 0, max: 200, major: 20 }, { min: 0, max: 200, major: 20 }, { min: 0, max: 450, major: 50 }]);
  const col = 64; // (a default column: 64 pixels for Calibri 11, docs/LAYOUT.md)
  assert.deepEqual([charts[0].from.c * col, charts[0].from.r * 20, (charts[0].to.c - charts[0].from.c) * col, (charts[0].to.r - charts[0].from.r) * 20], [320, 20, 512, 300], '240,15pt and 384 by 225pt in pixels');
  // (The series follow their cells: a value changed, the chart's data and axis with it.)
  wb.apply(0, [{ r: 6, c: 1, value: 500 }]);
  const again = readCharts(wb, 0)[0];
  assert.equal(again.series[0].values[5], 500);
  assert.deepEqual(valueAxis(again), { min: 0, max: 600, major: 100 });
});
