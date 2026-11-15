import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { chartParts } from '../src/ooxml/chartedit.js';
import { excel, withChart, EDITS, bare } from './chartformat-cases.js';

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

test('Excel reads the formatted charts written here as it reads its own (test/chartformat-readback-excel.json)', async () => {
  // (Every change on a sheet of its own, saved to test/out/charts/chartformat-all.xlsx: excel-chart-format-check.ps1
  // read it in Excel, and Excel's own workbook of the same changes, each chart's titles, series' types, axes and
  // colours, areas and fonts; the two readings compared here.)
  const { insertChart } = await import('../src/ooxml/chartnew.js');
  const { parseRange } = await import('../src/ooxml/refs.js');
  const { spawnSync } = await import('node:child_process');
  const names = Object.keys(EDITS);
  const wb = withChart(names[0]);
  const data = [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90], ['West', 60, 45]];
  for (let k = 1; k < names.length; k++) {
    wb.addSheet(k, names[k]);
    data.forEach((row, r) => row.forEach((v, c) => wb.setCell(k, { r, c }, v)));
    insertChart(wb, k, parseRange('A1:C5'), 'columnClustered', { at: { r: 1, c: 5 } });
  }
  names.forEach((n, k) => EDITS[n](wb, k));
  fs.mkdirSync('test/out/charts', { recursive: true });
  fs.writeFileSync('test/out/charts/chartformat-all.xlsx', wb.toBytes());
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/charts/chartformat-all.xlsx'], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
  const read = JSON.parse(fs.readFileSync('test/chartformat-readback-excel.json', 'utf8'));
  // (Excel's own workbook, with the chart style parts it writes beside each chart, reports a line or fill that isn't
  // there as visible black, 000000, its unedited chart's too, where it reads ours as not visible: taken as none.)
  const norm = (x) => JSON.stringify(x).replace(/"000000"/g, '"none"');
  const wrong = names.filter((n) => norm(read.ours[n]) !== norm(read.excel[n])).map((n) => `${n}: Excel's own ${JSON.stringify(read.excel[n])}, ours ${JSON.stringify(read.ours[n])}`);
  assert.deepEqual(wrong, []);
  assert.equal(Object.keys(read.ours).length, names.length);
});
