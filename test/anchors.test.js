import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { shiftCells } from '../src/ooxml/rowcol.js';
import { readNotes } from '../src/ooxml/notes.js';
import { parseRange, refName, colIndex } from '../src/ooxml/refs.js';

// test/anchors-excel.json (tools/excel-compare/excel-anchors.ps1 on anchor-cases.mjs --no-table): Excel's chart,
// rectangle and notes after each insert or delete, on the sheet test/fixtures/anchors.xlsx holds.
const EXCEL = JSON.parse(fs.readFileSync('test/anchors-excel.json', 'utf8').replace(/^\uFEFF/, ''));
const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/anchor-cases-plain.json', 'utf8')).cases;
const EMU_PT = 12700; // (English Metric Units in a point; default rows 15pt, columns 48pt here)

/** The sheet's drawings as Excel's COM gives them: top-left and bottom-right cells, height and width in points. */
function boxes(wb) {
  const text = wb.pkg.text('xl/drawings/drawing1.xml');
  const out = { chart: [], shape: [] };
  for (const m of text.matchAll(/<xdr:twoCellAnchor>([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
    const n = (tag, side) => Number(new RegExp(`<xdr:${side}>[\\s\\S]*?<xdr:${tag}>(-?\\d+)</xdr:${tag}>`).exec(m[1])[1]);
    const [c0, r0, c1, r1] = [n('col', 'from'), n('row', 'from'), n('col', 'to'), n('row', 'to')];
    const box = {
      from: refName(r0, c0), to: refName(r1, c1),
      height: (r1 - r0) * 15 + (n('rowOff', 'to') - n('rowOff', 'from')) / EMU_PT,
      width: (c1 - c0) * 48 + (n('colOff', 'to') - n('colOff', 'from')) / EMU_PT,
    };
    out[m[1].includes('graphicFrame') ? 'chart' : 'shape'].push(box);
  }
  return out;
}
const excelBox = (b) => ({ from: b.from, to: b.to, height: b.height, width: b.width });
const order = (refs) => [...refs].sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));

/** One step of a case, done as the editor does it. */
function step(wb, s) {
  const lines = /^(\d+):(\d+)$/.exec(s.range) || /^([A-Z]+):([A-Z]+)$/.exec(s.range);
  if (!s.shift && lines) {
    const rows = /^\d/.test(s.range);
    const a = rows ? Number(lines[1]) - 1 : colIndex(lines[1]);
    const b = rows ? Number(lines[2]) - 1 : colIndex(lines[2]);
    wb.changeLines(0, rows ? 'row' : 'col', a, (s.do === 'insert' ? 1 : -1) * (b - a + 1));
    return;
  }
  const sel = parseRange(s.range);
  shiftCells(wb, 0, sel, s.do === 'insert' ? 1 : -1, s.shift === 'down' || s.shift === 'up' ? 'row' : 'col');
}

CASES.forEach((c, k) => {
  test(`notes and drawings, ${c.name}: where Excel puts them`, () => {
    const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/anchors.xlsx')));
    step(wb, c.step);
    const e = EXCEL[k];
    assert.equal(e.name, c.name);
    const got = boxes(wb);
    assert.deepEqual(got.chart, (e.charts || []).map(excelBox), 'the chart');
    assert.deepEqual(got.shape, (e.shapes || []).map(excelBox), 'the rectangle');
    assert.deepEqual(order([...readNotes(wb, 0).keys()].map((key) => { const [r, col] = key.split(',').map(Number); return refName(r, col); })), order([].concat(e.notes || [])), 'the notes');
  });
});

test('a chart wholly deleted goes with its part; the saved workbook is whole', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/anchors.xlsx')));
  wb.changeLines(0, 'col', 5, -6);
  assert.ok(!wb.pkg.has('xl/charts/chart1.xml'), 'the chart part removed');
  assert.ok(!wb.pkg.text('xl/drawings/_rels/drawing1.xml.rels').includes('chart1.xml'), 'and its relationship');
  const again = Workbook.fromBytes(wb.toBytes());
  assert.equal(again.sheets.length, 1);
});
