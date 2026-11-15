import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseXml } from '../src/ooxml/xml.js';
import { readSparklines, sparklineLooks, parseSheetRange } from '../src/ooxml/sparklines.js';
import { sparkShapes, sparkPaths, markerSize } from '../src/editor/sparkgeom.js';
import { SPARK_ROWS } from './fixtures/sparklines.mjs';

// Sparklines (docs/SPARKLINES.md). test/sparklines-xml-excel.json is what tools/excel-compare/excel-sparklines.ps1
// recorded: 15 groups built over COM with each option, COM's view of each and the extension Excel saved.
// test/sparklines-excel.json is what Excel drew for test/fixtures/sparklines.xlsx: its PDF's 28 sparkline pictures,
// measured by tools/excel-compare/spark-features.py (bars, blobs of colour, the line's trace).
const probe = JSON.parse(fs.readFileSync('test/sparklines-xml-excel.json', 'utf8'));
const drawn = JSON.parse(fs.readFileSync('test/sparklines-excel.json', 'utf8'));
const open = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/sparklines.xlsx')));
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

test('Excel\'s saved sparkline groups read as COM reported them', () => {
  const doc = parseXml(`<worksheet xmlns="${MAIN}" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><extLst>${probe.ext}</extLst></worksheet>`);
  const groups = readSparklines(doc.documentElement, MAIN);
  assert.equal(groups.length, 15);
  const TYPES = { 1: 'line', 2: 'column', 3: 'stacked' };
  const EMPTY = { 1: 'gap', 2: 'zero', 3: 'span' };
  const AXIS = { 1: 'group', 2: 'individual', 3: 'custom' };
  const bgr = (n) => `${[n & 255, (n >> 8) & 255, (n >> 16) & 255].map((v) => v.toString(16).padStart(2, '0')).join('')}`.toUpperCase();
  for (const [name, com] of Object.entries(probe.groups)) {
    const g = groups.find((x) => x.sparklines.map((s) => `${String.fromCharCode(65 + s.c)}${s.r + 1}`).join(':') === com.location);
    assert.ok(g, name);
    assert.equal(g.type, TYPES[com.type], name);
    assert.equal(g.lineWeight, com.lineWeight, name);
    assert.equal(g.empty, EMPTY[com.blanks], name);
    assert.equal(g.hidden, com.hidden, name);
    assert.equal(g.rightToLeft, com.rtl, name);
    assert.equal(g.minType, AXIS[com.minType], name);
    assert.equal(g.maxType, AXIS[com.maxType], name);
    if (com.maxType === 3) assert.deepEqual([g.min, g.max], [com.min, com.max], name);
    assert.equal(g.dates, com.dateRange ? `Sheet1!${com.dateRange}` : null, name);
    assert.equal(g.axis, com.axis[0], name);
    for (const k of ['negative', 'markers', 'high', 'low', 'first', 'last']) {
      assert.equal(g[k], com[k][0], `${name} ${k}`);
      assert.equal(g.colors[k].rgb.slice(2), bgr(com[k][1]), `${name} ${k} colour`);
    }
    assert.equal(g.colors.series.rgb.slice(2), bgr(com.series), name);
    assert.equal(g.sparklines[0].source, `Sheet1!${com.source.split(':').length === 2 && com.location.includes(':') ? com.source.replace(/(\d+):([A-Z]+)(\d+)/, (m, a, c) => `${a}:${c}${a}`) : com.source}`, name);
  }
});

test('references with a quoted sheet name and dollars', () => {
  assert.deepEqual(parseSheetRange("'Q1 ''data'''!$B$2:$M$2"), { sheet: "Q1 'data'", range: { r0: 1, c0: 1, r1: 1, c1: 12 } });
  assert.equal(parseSheetRange('B2:M2'), null);
});

test('the fixture\'s sparklines: values, hidden data, text, empty cells, dates, axis and colours', () => {
  const wb = open();
  const groups = wb.sheet(0).sparklines;
  assert.equal(groups.length, 27);
  assert.deepEqual(groups[16].sparklines.map((s) => [s.r, s.c]), [[17, 14], [18, 14]]);
  const looks = sparklineLooks(wb, 0);
  const at = (r) => looks.get(`${r - 1},14`);
  assert.equal(looks.size, 28);
  // (Excel's Style 1 in the fixture's theme: accent 1 darkened by half, the colour of Excel's own pictures.)
  assert.equal(at(2).colors.series, '#203764');
  assert.equal(drawn[0].main, '#203764');
  assert.equal(at(2).colors.negative, '#ED7D31');
  assert.deepEqual([at(2).lo, at(2).hi], [2, 10]);
  // Hidden column N: left out, or shown.
  assert.equal(at(24).values.length, 12);
  assert.deepEqual([at(25).values.length, at(25).hi], [13, 30]);
  // Empty cells: kept as gaps, made 0.
  assert.equal(at(14).values[2], null);
  assert.equal(at(15).values[2], 0);
  assert.equal(at(15).lo, 0);
  // Text: plotted as 0.
  assert.deepEqual([at(27).values[2], at(27).lo], [0, 0]);
  // Group and custom axes.
  assert.deepEqual([at(18).lo, at(18).hi, at(19).lo, at(19).hi], [2, 24, 2, 24]);
  assert.deepEqual([at(20).lo, at(20).hi], [0, 20]);
  assert.deepEqual([at(21).lo, at(21).hi], [1, 8]);
  assert.deepEqual(at(22).dates, [45658, 45659, 45661, 45665, 45666, 45670, 45671, 45680, 45681, 45682, 45690, 45700]);
});

test('a sparkline follows its data when a cell changes', () => {
  const wb = open();
  wb.setCell(0, 'B2', 50);
  assert.equal(sparklineLooks(wb, 0).get('1,14').hi, 50);
});

test('marker sizes by line weight, as measured', () => {
  assert.equal(markerSize(0.75), 3.36);
  assert.equal(markerSize(3), 9.12);
  assert.equal(+markerSize(2.25).toFixed(2), 7.68);
  assert.ok(markerSize(9) > 12.96);
});

/** Each picture's row of the fixture (the group O18:O19 draws two). */
const rowOf = (k) => k + 2;

test('drawn as Excel\'s PDF draws them', (t) => {
  const wb = open();
  const looks = sparklineLooks(wb, 0);
  assert.equal(drawn.length, 28);
  const worst = { barEdge: 0, trace: 0, marker: 0, markerSize: 0, axis: 0 };
  const sum = { barEdge: 0, trace: 0 };
  const counted = { bars: 0, segments: 0, markers: 0 };
  drawn.forEach((ex, k) => {
    const r = rowOf(k);
    const label = SPARK_ROWS[r - 2][0];
    const sp = looks.get(`${r - 1},14`);
    const s = sparkShapes(sp, ex.w, ex.h);
    assert.ok(sparkPaths(s).length, label);
    if (sp.type !== 'line') {
      // Bars, left to right: as many, each edge within a pixel of Excel's, the same colour. (Hairlines, a picture
      // pixel tall, are left out: too faint to be found reliably in Excel's picture.)
      const ours = s.bars.filter((b) => b.h > 1).sort((a, b) => a.x - b.x);
      const theirs = ex.bars.filter((b) => b[3] - b[2] > 1).sort((a, b) => a[0] - b[0]);
      assert.equal(ours.length, theirs.length, label);
      ours.forEach((b, i) => {
        const [x0, x1, y0, y1, color] = theirs[i];
        const d = Math.max(Math.abs(b.x - x0), Math.abs(b.x + b.w - x1), Math.abs(b.y - y0), Math.abs(b.y + b.h - y1));
        worst.barEdge = Math.max(worst.barEdge, d);
        sum.barEdge += d;
        assert.ok(d <= 1, `${label} bar ${i}: ${d.toFixed(2)}px`);
        if (b.w > 2) assert.equal(color, b.color, `${label} bar ${i} colour`);
        counted.bars++;
      });
    } else if (ex.main === sp.colors.series || label === 'All the same' || label === 'Thin line') {
      // The line: each segment's middle on Excel's trace of it (at a vertex the stroke's join pulls the trace away).
      for (const run of s.line.runs) {
        for (let i = 1; i < run.length; i++) {
          const xm = (run[i - 1][0] + run[i][0]) / 2;
          const ym = (run[i - 1][1] + run[i][1]) / 2;
          const tr = ex.trace[Math.round(xm / ex.step)];
          if (tr == null) continue;
          const d = Math.abs(tr - ym);
          worst.trace = Math.max(worst.trace, d);
          sum.trace += d;
          // (Steep segments' traces are a picture pixel or so off where Excel's antialiasing steps.)
          assert.ok(d <= 1.3, `${label} segment ${i}: ${d.toFixed(2)}px`);
          counted.segments++;
        }
      }
      // Markers of their own colour: each where Excel put it, its size within a pixel.
      for (const [x, y, w, , color] of ex.blobs) {
        if (color === sp.colors.axis && w > 50) continue;
        const near = s.markers.filter((m) => m.color === color).sort((a, b) => Math.hypot(a.x - x, a.y - y) - Math.hypot(b.x - x, b.y - y))[0];
        assert.ok(near, `${label}: no marker of ${color} near ${x}, ${y}`);
        const d = Math.hypot(near.x - x, near.y - y);
        worst.marker = Math.max(worst.marker, d);
        worst.markerSize = Math.max(worst.markerSize, Math.abs(near.size - w));
        assert.ok(d <= 1.2, `${label} marker at ${x}, ${y}: ${d.toFixed(2)}px`);
        counted.markers++;
      }
    }
    // The axis: at Excel's height, within a pixel.
    const exAxis = ex.main === sp.colors.axis ? { y: ex.trace.filter((v) => v != null)[0] } : ex.blobs.filter((b) => b[4] === sp.colors.axis && b[2] > 50).map((b) => ({ y: b[1] }))[0];
    assert.equal(!!s.axis, !!exAxis, `${label} axis`);
    if (exAxis) {
      worst.axis = Math.max(worst.axis, Math.abs(s.axis.y - exAxis.y));
      assert.ok(Math.abs(s.axis.y - exAxis.y) <= 1, `${label} axis at ${exAxis.y}`);
    }
  });
  t.diagnostic(`compared ${counted.bars} bars, ${counted.segments} line segments, ${counted.markers} markers; mean bar edge ${(sum.barEdge / counted.bars).toFixed(2)}px, mean line ${(sum.trace / counted.segments).toFixed(2)}px; worst: ${JSON.stringify(Object.fromEntries(Object.entries(worst).map(([a, b]) => [a, +b.toFixed(2)])))}`);
});
