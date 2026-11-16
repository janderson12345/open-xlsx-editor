import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { SheetError } from '../src/ooxml/sheetops.js';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { doStep } from './refmoves-steps.js';

// test/slicers-excel.json (tools/excel-compare/excel-slicers.ps1 on slicer-cases.mjs, cut down by
// refmove-results.mjs): where Excel put the slicers and the timeline, what it kept of their caches and parts, the
// table and the pivot, after each insert or delete on the sheet test/fixtures/slicers.xlsx holds.
const EXCEL = JSON.parse(fs.readFileSync('test/slicers-excel.json', 'utf8'));
const OUT = 'test/out/slicers';
fs.mkdirSync(OUT, { recursive: true });
const book = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/slicers.xlsx')));

/** The workbook's state in refmove-results.mjs's terms. */
function stateOf(wb) {
  const files = Object.keys(wb.pkg.files);
  const text = (re) => files.filter((p) => re.test(p)).map((p) => wb.pkg.text(p));
  const anchors = {};
  for (const m of (wb.pkg.text('xl/drawings/drawing1.xml') || '').matchAll(/<xdr:twoCellAnchor editAs="(\w+)">([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
    const pt = (which) => /<xdr:col>(\d+)<\/xdr:col><xdr:colOff>(-?\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(-?\d+)<\/xdr:rowOff>/.exec(new RegExp(`<xdr:${which}>([\\s\\S]*?)</xdr:${which}>`).exec(m[2])[1]).slice(1).join(',');
    anchors[/<(?:sle|tsle):(?:slicer|timeslicer)\b[^>]*name="([^"]*)"/.exec(m[2])[1]] = `${m[1]} ${pt('from')} ${pt('to')}`;
  }
  const table = text(/^xl\/tables\//)[0];
  const pivot = text(/^xl\/pivotTables\/[^/]*\.xml$/)[0];
  return {
    anchors,
    caches: text(/^xl\/(slicerCaches|timelineCaches)\//).map((t) => {
      const tc = /<x15:tableSlicerCache tableId="(\d+)" column="(\d+)"/.exec(t);
      return `${/\sname="([^"]*)"/.exec(t)[1]}${tc ? ` table ${tc[1]} column ${tc[2]}` : ''}`;
    }).sort(),
    slicers: text(/^xl\/(slicers|timelines)\//).flatMap((t) => [...t.matchAll(/<(?:slicer|timeline) name="([^"]*)"/g)].map((m) => m[1])).sort(),
    table: table ? `${/ ref="([^"]*)"/.exec(table)[1]} ${[...table.matchAll(/<tableColumn id="(\d+)"[^>]*name="([^"]*)"/g)].map((m) => `${m[1]}:${m[2]}`).join(',')}` : null,
    pivot: pivot ? /<location ref="([^"]*)"/.exec(pivot)[1] : null,
  };
}

test('the probe\'s sheet: its slicers, timeline, table and pivot as Excel saved them', () => {
  const wb = book();
  fs.writeFileSync(path.join(OUT, '1.xlsx'), wb.toBytes());
  const got = stateOf(wb);
  const start = EXCEL[0];
  for (const k of ['anchors', 'caches', 'slicers', 'table', 'pivot']) assert.deepEqual(got[k], start[k], k);
});

for (const excel of EXCEL.slice(1)) {
  test(`${excel.name}: as Excel does it`, () => {
    const wb = book();
    let refused = null;
    try { doStep(wb, excel.step); } catch (e) { if (!(e instanceof SheetError)) throw e; refused = e.message; }
    // (Refused where Excel refuses: COM's own words differ from the box's, so only that it was.)
    assert.equal(!!refused, !!excel.refused, refused || excel.refused || 'not refused');
    const got = stateOf(wb);
    for (const k of ['anchors', 'caches', 'slicers', 'table', 'pivot']) assert.deepEqual(got[k], excel[k], k);
    // The saved workbook (checked by tools/check-ooxml.mjs below): the deleted slicer's name and lists gone.
    const bytes = wb.toBytes();
    fs.writeFileSync(path.join(OUT, `${EXCEL.indexOf(excel) + 1}.xlsx`), bytes);
    const book2 = Workbook.fromBytes(bytes);
    const names = book2.definedNames.map((d) => d.name).sort();
    const caches = excel.caches.map((c) => c.split(' ')[0]);
    assert.deepEqual(names, caches.sort());
    const lists = (book2.pkg.text('xl/workbook.xml').match(/<x14:slicerCache r:id=/g) || []).length;
    assert.equal(lists, caches.filter((c) => c.startsWith('Slicer_')).length);
  });
}

test('every saved workbook passes the structural check (parts, content types, relationships)', () => {
  const run = spawnSync(process.execPath, ['tools/check-ooxml.mjs', OUT], { encoding: 'utf8' });
  assert.equal(run.status, 0, run.stdout + run.stderr);
});
