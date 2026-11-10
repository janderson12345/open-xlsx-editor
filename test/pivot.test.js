import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PIVOT, pivotBook } from './pivots.js';
import { readPivots, refreshPivot, pivotAt } from '../src/ooxml/pivot.js';

// Each case of test/pivot-excel.json (tools/excel-compare/excel-pivot.ps1): Excel built a pivot table of made-up sales,
// then changed the source (row 2's units to 100, its region to a new one, Central) and refreshed it.
const cellsOf = (wb, si, area) => {
  const out = [];
  const sh = wb.sheet(si);
  for (let r = area.r0; r <= area.r1; r++) for (let c = area.c0; c <= area.c1; c++) {
    const v = sh.cells.get(r)?.get(c)?.v;
    if (v != null && v !== '') out.push([r, c, v]);
  }
  return out;
};
const excelCells = (cells) => cells.map(([r, c, v]) => [r, c, v]);
/** The elements a refresh rewrites, from a part's text. */
const pick = (text, ...names) => names.map((n) => (text.match(new RegExp(`<${n}\\b[^>]*?(?:/>|>[\\s\\S]*?</${n}>)`, 'g')) || []).join('')).join('\n');
const tableBits = (t) => pick(t, 'location', 'pivotFields', 'rowItems', 'colItems');
const cacheBits = (t) => `${/recordCount="\d+"/.exec(t)[0]}\n${pick(t, 'sharedItems')}`;
const records = (t) => t.replace(/^[\s\S]*?<pivotCacheRecords\b[^>]*>/, '');

PIVOT.cases.forEach((c, k) => {
  test(`pivot ${c.name}: refreshed unchanged, Excel's cells and parts`, () => {
    const { wb, piv } = pivotBook(k, 'before');
    const p = readPivots(wb)[0];
    const out = refreshPivot(wb, p);
    assert.ok(!out.error, out.error);
    for (const ch of out.changes) wb.apply(ch.si, ch.contents);
    assert.deepEqual(out.changes.flatMap((x) => x.contents), [], 'nothing to change in the cells');
    assert.equal(tableBits(wb.pkg.text(p.path)), tableBits(c.parts.before.table));
    assert.equal(cacheBits(wb.pkg.text(p.cachePath)), cacheBits(c.parts.before.cache));
    assert.equal(records(wb.pkg.text(p.recordsPath)), records(c.parts.before.records));
    assert.ok(pivotAt(wb, piv, 2, 0));
  });

  test(`pivot ${c.name}: refreshed after the source changed, as Excel refreshed it`, () => {
    const { wb, src, piv } = pivotBook(k, 'before');
    for (const ch of PIVOT.change) wb.apply(src, [{ r: ch.r, c: ch.c, value: ch.value }]);
    const p = readPivots(wb)[0];
    const out = refreshPivot(wb, p);
    assert.ok(!out.error, out.error);
    for (const ch of out.changes) wb.apply(ch.si, ch.contents);
    const area = { r0: 0, c0: 0, r1: 40, c1: 10 };
    assert.deepEqual(cellsOf(wb, piv, area), excelCells(c.after.cells));
    assert.equal(tableBits(wb.pkg.text(p.path)), tableBits(c.parts.after.table));
    assert.equal(cacheBits(wb.pkg.text(p.cachePath)), cacheBits(c.parts.after.cache));
    assert.equal(records(wb.pkg.text(p.recordsPath)), records(c.parts.after.records));
  });
});

test("a pivot whose source lost a heading is refused in Excel's words; renamed headings and report filters aren't refreshed here", async () => {
  const { PIVOT_MESSAGES } = await import('../src/ooxml/pivot.js');
  const one = pivotBook(0);
  one.wb.apply(one.src, [{ r: 0, c: 0, value: null }]);
  const p = readPivots(one.wb)[0];
  const cache = one.wb.pkg.text(p.cachePath);
  assert.deepEqual(refreshPivot(one.wb, p), { error: PIVOT_MESSAGES.fieldName });
  assert.equal(one.wb.pkg.text(p.cachePath), cache, 'nothing changed');
  one.wb.apply(one.src, [{ r: 0, c: 0, value: 'Area' }]);
  assert.deepEqual(refreshPivot(one.wb, p), { error: PIVOT_MESSAGES.unsupported, why: 'a heading of its source renamed' });
  const two = pivotBook(0);
  const q = readPivots(two.wb)[0];
  two.wb.pkg.setText(q.path, two.wb.pkg.text(q.path).replace('<dataFields', '<pageFields count="1"><pageField fld="2" hier="-1"/></pageFields><dataFields'));
  assert.deepEqual(refreshPivot(two.wb, q), { error: PIVOT_MESSAGES.unsupported, why: 'report filters' });
});
