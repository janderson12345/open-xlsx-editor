import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { readPivots } from '../src/ooxml/pivot.js';
import { movePivotFields, pivotLayout } from '../src/ooxml/pivotfields.js';
import { PIVOT, pivotBook } from './pivots.js';

// test/pivot-move-excel.json (tools/excel-compare/excel-pivot-move.ps1): Excel's pivot of one case's layout, its
// fields then moved to another case's, the three parts it saved and its cells. Here: Excel's own pivot of the first
// layout (test/pivots.js), its fields moved, and the parts and cells must be Excel's.
const MOVES = JSON.parse(fs.readFileSync('test/pivot-move-excel.json', 'utf8'));
const FIELDS = PIVOT.data[0];
const layoutOf = (c) => ({ rows: c.rows.map((f) => FIELDS.indexOf(f)), cols: c.cols.map((f) => FIELDS.indexOf(f)), values: c.values.map((v) => ({ field: FIELDS.indexOf(v.field), fn: v.fn })) });
const plain = (xml) => xml
  .replace(/ xr:uid="\{[^}]*\}"/g, '')
  .replace(/ refreshedDate="[^"]*"/, '')
  .replace(/ r:id="[^"]*"/, '')
  .replace(/ cacheId="\d+"/, '')
  .replace(/ sheet="(?:Data|Src)"/, ' sheet="Src"')
  .replace(/<ext uri="\{F4336242-56EE-4DD3-A5F7-E900642BF6E9\}"[\s\S]*?<\/ext>/, '');

MOVES.forEach((m) => {
  test(`a pivot's fields moved, ${PIVOT.cases[m.from].name} to ${PIVOT.cases[m.to].name}: Excel's parts and cells`, () => {
    const { wb, piv } = pivotBook(m.from, 'before');
    const p = readPivots(wb)[0];
    const out = movePivotFields(wb, p, layoutOf(PIVOT.cases[m.to]));
    assert.ok(!out.error, out.error);
    for (const ch of out.changes) if (ch.contents.length) wb.apply(ch.si, ch.contents);
    assert.equal(plain(wb.pkg.text(p.path)), plain(m.parts.table), 'table');
    assert.equal(plain(wb.pkg.text(p.cachePath)), plain(m.parts.cache), 'cache');
    assert.equal(plain(wb.pkg.text(p.recordsPath)), plain(m.parts.records), 'records');
    const cells = [];
    for (const [r, line] of wb.sheet(piv).cells) for (const [c, cell] of line) if (cell.v != null) cells.push([r, c, cell.v]);
    const want = m.cells.cells.map(([r, c, v]) => [r, c, v]);
    assert.deepEqual(cells.sort((a, b) => a[0] - b[0] || a[1] - b[1]), want.sort((a, b) => a[0] - b[0] || a[1] - b[1]));
  });
});

test('a pivot\'s layout read for the field list, and the layouts it refuses', () => {
  const { wb } = pivotBook(3, 'before');
  const p = readPivots(wb)[0];
  const l = pivotLayout(wb, p);
  assert.deepEqual({ rows: l.rows, cols: l.cols, values: l.values.map((v) => `${v.field}:${v.fn}:${v.name}`) }, { rows: [0], cols: [], values: ['3:sum:Sum of Units', '4:average:Average of Price'] });
  assert.match(movePivotFields(wb, p, { rows: [], cols: [], values: [{ field: 3, fn: 'sum' }] }).error, /Rows or Columns/);
});
