import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { createPivot, layoutRefusal } from '../src/ooxml/pivotnew.js';
import { readPivots, refreshPivot, PIVOT_MESSAGES } from '../src/ooxml/pivot.js';
import { PIVOT } from './pivots.js';

// test/pivot-excel.json: for each case, the parts Excel wrote for the PivotTable it had just made (excel-pivot.ps1),
// and its cells. Made here from the same data and fields, the parts and cells must be Excel's, ids and dates aside.
const FIELDS = PIVOT.data[0];

/** A workbook with the data on a sheet "Src" (Excel's was "Data") from A1 and an empty sheet "Pivot", as excel-pivot.ps1 makes them. */
function book() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const data = wb.addSheet(wb.sheets.length, 'Src');
  const piv = wb.addSheet(wb.sheets.length, 'Pivot');
  wb.apply(data, PIVOT.data.flatMap((row, r) => row.map((value, c) => ({ r, c, value }))));
  return { wb, data, piv, src: { si: data, r0: 0, c0: 0, r1: PIVOT.data.length - 1, c1: FIELDS.length - 1 } };
}
const layoutOf = (c) => ({ rows: c.rows.map((f) => FIELDS.indexOf(f)), cols: c.cols.map((f) => FIELDS.indexOf(f)), values: c.values.map((v) => ({ field: FIELDS.indexOf(v.field), fn: v.fn })) });

/** A part with what differs from one making to the next taken out: ids, dates, relationship ids, the cache's id,
 *  and Excel 2026's data-source extension (written only by it). */
const plain = (xml) => xml
  .replace(/ xr:uid="\{[^}]*\}"/g, '')
  .replace(/ refreshedDate="[^"]*"/, '')
  .replace(/ r:id="[^"]*"/, '')
  .replace(/ cacheId="\d+"/, '')
  .replace(/ sheet="(?:Data|Src)"/, ' sheet="Src"')
  .replace(/<ext uri="\{F4336242-56EE-4DD3-A5F7-E900642BF6E9\}"[\s\S]*?<\/ext>/, '');

PIVOT.cases.forEach((c) => {
  test(`a new PivotTable, ${c.name}: Excel's parts and cells`, () => {
    const { wb, piv, src } = book();
    const out = createPivot(wb, src, { si: piv, r: 2, c: 0 }, layoutOf(c), { name: 'Pivot1' });
    assert.ok(!out.error, out.error);
    for (const ch of out.changes) wb.apply(ch.si, ch.contents);
    const p = readPivots(wb).find((x) => x.name === 'Pivot1');
    const want = c.parts.before;
    assert.equal(plain(wb.pkg.text(p.path)), plain(want.table), 'table');
    assert.equal(plain(wb.pkg.text(p.cachePath)), plain(want.cache), 'cache');
    assert.equal(plain(wb.pkg.text(p.recordsPath)), plain(want.records), 'records');
    const cells = [];
    for (const [r, line] of wb.sheet(piv).cells) for (const [col, cell] of line) if (cell.v != null) cells.push([r, col, cell.v]);
    assert.deepEqual(cells.sort((a, b) => a[0] - b[0] || a[1] - b[1]), [...c.before.cells].sort((a, b) => a[0] - b[0] || a[1] - b[1]));
  });
});

test('a new PivotTable saves, opens again and refreshes as one Excel made', () => {
  const { wb, piv, src } = book();
  const out = createPivot(wb, src, { si: piv, r: 2, c: 0 }, layoutOf(PIVOT.cases[1]));
  for (const ch of out.changes) wb.apply(ch.si, ch.contents);
  assert.equal(out.name, 'PivotTable1');
  const again = Workbook.fromBytes(wb.toBytes());
  const p = readPivots(again);
  assert.equal(p.length, 1);
  assert.ok(!refreshPivot(again, p[0]).error);
  assert.match(again.pkg.text('xl/workbook.xml'), /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
  // (A second one takes the next name, parts and cache id.)
  const two = createPivot(again, { ...src, si: again.sheetIndexOf('Src') }, { si: again.sheetIndexOf('Pivot'), r: 2, c: 8 }, layoutOf(PIVOT.cases[0]));
  assert.equal(two.name, 'PivotTable2');
  assert.equal(two.path, 'xl/pivotTables/pivotTable2.xml');
  assert.match(Workbook.fromBytes(again.toBytes()).pkg.text('xl/workbook.xml'), /<pivotCache cacheId="1" r:id="rId\d+"\/><pivotCache cacheId="2" r:id="rId\d+"\/>/);
});

test('what a new PivotTable refuses: an empty heading in Excel\'s words, and layouts not made here', () => {
  const { wb, piv, src } = book();
  wb.apply(src.si, [{ r: 0, c: 2, value: null }]);
  assert.equal(createPivot(wb, src, { si: piv, r: 2, c: 0 }, layoutOf(PIVOT.cases[0])).error, PIVOT_MESSAGES.fieldName);
  assert.equal(createPivot(wb, { ...src, r1: 0 }, { si: piv, r: 2, c: 0 }, layoutOf(PIVOT.cases[0])).error, PIVOT_MESSAGES.source);
  assert.match(layoutRefusal({ rows: [], cols: [], values: [{ field: 3, fn: 'sum' }] }), /Rows or Columns/);
  assert.match(layoutRefusal({ rows: [0], cols: [1, 2], values: [{ field: 3, fn: 'sum' }] }), /More than one column/);
  assert.match(layoutRefusal({ rows: [0], cols: [1], values: [{ field: 3, fn: 'sum' }, { field: 4, fn: 'sum' }] }), /Several values/);
  assert.equal(readPivots(wb).length, 0);
});
