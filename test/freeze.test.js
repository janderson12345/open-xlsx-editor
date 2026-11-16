import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
import { freezeFor } from '../src/editor/freeze.js';

// Excel's Freeze Panes on a scrolled, split or partly hidden window (tools/excel-compare/excel-freeze.ps1, its ribbon's
// own commands on a new workbook in an invisible Excel, 64 rows by 53 columns in view; test/freeze-excel.json): the same
// steps here, and the sheet view written compared with Excel's.
const EXCEL = JSON.parse(fs.readFileSync('test/freeze-excel.json', 'utf8'));
const at = (ref) => { const p = parseRef(ref); return { r: p.r, c: p.c }; };
const fresh = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  return { wb, si: wb.addSheet(wb.sheets.length) };
};
/** The sheet view's topLeftCell and what is inside it, as written. */
const written = (wb, si) => {
  const xml = strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]);
  const m = /<sheetView\b([^>]*)>([\s\S]*?)<\/sheetView>/.exec(xml) || [null, /<sheetView\b([^>]*)\/>/.exec(xml)[1], ''];
  return { topLeft: /topLeftCell="([^"]+)"/.exec(m[1])?.[1] || null, inside: m[2] };
};
const excel = (name) => {
  const x = EXCEL.find((e) => e.name === name);
  assert.ok(x, name);
  const m = /<sheetView\b([^>]*)>([\s\S]*?)<\/sheetView>/.exec(x.sheetViews) || [null, /<sheetView\b([^>]*)\/>/.exec(x.sheetViews)[1], ''];
  return { topLeft: /topLeftCell="([^"]+)"/.exec(m[1])?.[1] || null, inside: m[2] };
};
/** A grid as freezeFor sees it: Excel's default sizes (20 pixel rows, 64 pixel columns, headings 20 high and 26 wide),
 *  scrolled to `view`, `hidden` rows and columns of no size, Excel's window of 64 rows and 53 columns. */
const gridAt = (view = 'A1', { rows: hr = [], cols: hc = [] } = {}) => ({
  view: () => ({ topLeft: at(view) }),
  cellsInView: () => ({ rows: 64, cols: 53 }),
  geometry: () => ({ rows: { size: (r) => (hr.includes(r) ? 0 : 20) }, cols: { size: (c) => (hc.includes(c) ? 0 : 64) }, rowHeadW: 26, colHeadH: 20 }),
});
/** Freeze `kind` with `active` on a sheet of workbook `wb` scrolled to `view`, as the View tab's command does. */
function freeze(wb, si, kind, active, view = 'A1', hidden = {}) {
  const sheet = wb.sheet(si);
  const frozen = !!(sheet.freeze.rows || sheet.freeze.cols);
  const f = freezeFor(kind, { sheet, frozen, grid: gridAt(view, hidden), active: at(active) });
  wb.setView(si, { freeze: f, active: at(active) });
  return f;
}
const same = (name, wb, si) => assert.deepEqual(written(wb, si), excel(name), name);

test('Freeze Panes counts from the cell first in view: scrolled, the middle of the window with A1 or a cell out of view', () => {
  const cases = [
    ['C5, not scrolled', 'C5', 'A1'],
    ['E15, scrolled to row 10 and column C', 'E15', 'C10'],
    ['B3, scrolled to row 10 and column C (active cell out of view)', 'B3', 'C10'],
    ['A1, not scrolled', 'A1', 'A1'],
    ['C1, not scrolled', 'C1', 'A1'],
    ['A5, not scrolled', 'A5', 'A1'],
    ['A10 active, scrolled to row 10', 'A10', 'A10'],
  ];
  for (const [name, active, view] of cases) {
    const { wb, si } = fresh();
    freeze(wb, si, 'panes', active, view);
    same(name, wb, si);
  }
});

test('Freeze Top Row and First Column take the row and column at the window’s top and left', () => {
  for (const [name, kind, active, view] of [
    ['Top Row, scrolled to row 10', 'topRow', 'D20', 'A10'],
    ['Top Row, scrolled to row 10 and column D', 'topRow', 'F20', 'D10'],
    ['First Column, scrolled to column D', 'firstColumn', 'F3', 'D1'],
    ['First Column, scrolled to row 10 and column D', 'firstColumn', 'F20', 'D10'],
  ]) {
    const { wb, si } = fresh();
    freeze(wb, si, kind, active, view);
    same(name, wb, si);
  }
});

test('hidden rows and columns: counted in the freeze; a hidden active row or column moves it on to the next shown', () => {
  for (const [name, kind, active, hidden] of [
    ['rows 2 to 3 hidden, C5', 'panes', 'C5', { rows: [1, 2] }],
    ['columns B to C hidden, E5', 'panes', 'E5', { cols: [1, 2] }],
    ['rows 4 to 6 hidden, C5 (active cell hidden)', 'panes', 'C5', { rows: [3, 4, 5] }],
    ['columns C to D hidden, D5 (active column hidden)', 'panes', 'D5', { cols: [2, 3] }],
    ['row 1 hidden, Top Row', 'topRow', 'C5', { rows: [0] }],
  ]) {
    const { wb, si } = fresh();
    freeze(wb, si, kind, active, 'A1', hidden);
    same(name, wb, si);
  }
});

test('a split window: Freeze Panes freezes it where it is split (frozenSplit); Top Row freezes the window’s top row', () => {
  // (Excel's own split sheet view, saved split after 3 rows and 2 columns, put in a sheet and read as a file.)
  const split = excel('split at row 3 and column 2, saved split');
  for (const [name, kind] of [['split at row 3 and column 2, then Freeze Panes (ribbon), C5 active', 'panes'], ['split at row 3 and column 2, then Window.FreezePanes', 'panes'], ['split, then Top Row', 'topRow']]) {
    let { wb, si } = fresh();
    const path = wb.sheets[si].path;
    const xml = strFromU8(unzipSync(wb.toBytes())[path]).replace(/<sheetView\b[^>]*\/>|<sheetView\b[^>]*>[\s\S]*?<\/sheetView>/, `<sheetView topLeftCell="${split.topLeft}" workbookViewId="0">${split.inside}</sheetView>`);
    const files = unzipSync(wb.toBytes());
    files[path] = new TextEncoder().encode(xml);
    wb = Workbook.fromBytes(zipSync(files));
    assert.deepEqual(wb.sheet(si).split, { x: 2310, y: 1200 });
    assert.deepEqual(wb.sheet(si).scrolledTo, at('E7'), 'shown unsplit, from the split window’s first cell');
    freeze(wb, si, kind, 'E8', 'E7');
    same(name, wb, si);
  }
});

test('Unfreeze leaves the window where the freeze started; frozen again from there; a frozen sheet’s Top Row', () => {
  let { wb, si } = fresh();
  freeze(wb, si, 'panes', 'E15', 'C10');
  freeze(wb, si, 'none', 'E15');
  same('E15 frozen scrolled, then unfrozen', wb, si);
  ({ wb, si } = fresh());
  freeze(wb, si, 'panes', 'C5');
  freeze(wb, si, 'panes', 'C5');
  same('C5 frozen, then Freeze Panes again (the menu turned to Unfreeze)', wb, si);
  same('C5 frozen, bottom pane scrolled to row 40, then unfrozen', wb, si);
  ({ wb, si } = fresh());
  freeze(wb, si, 'panes', 'C5');
  freeze(wb, si, 'topRow', 'C5');
  same('C5 frozen, then Top Row', wb, si);
  ({ wb, si } = fresh());
  freeze(wb, si, 'panes', 'C5');
  freeze(wb, si, 'none', 'E8');
  freeze(wb, si, 'panes', 'E8');
  same('C5 frozen, then E8 selected and Freeze Panes on the frozen sheet with Window.FreezePanes', wb, si);
});

test('a frozen sheet’s view read as Excel saved it: where the freeze starts and where the scrolled pane was left', () => {
  const { wb, si } = fresh();
  const path = wb.sheets[si].path;
  const files = unzipSync(wb.toBytes());
  const view = (name) => { const v = excel(name); return `<sheetView ${v.topLeft ? `topLeftCell="${v.topLeft}" ` : ''}workbookViewId="0">${v.inside}</sheetView>`; };
  const load = (name) => {
    files[path] = new TextEncoder().encode(strFromU8(files[path]).replace(/<sheetView\b[^>]*\/>|<sheetView\b[^>]*>[\s\S]*?<\/sheetView>/, view(name)));
    return Workbook.fromBytes(zipSync(files)).sheet(si);
  };
  let s = load('C5 frozen, then the bottom pane scrolled to row 40 and column H');
  assert.deepEqual(s.freeze, { rows: 4, cols: 2, top: 0, left: 0 });
  assert.deepEqual(s.scrolledTo, at('H40'));
  s = load('E15, scrolled to row 10 and column C');
  assert.deepEqual(s.freeze, { rows: 5, cols: 2, top: 9, left: 2 });
  assert.deepEqual(s.scrolledTo, at('E15'));
  s = load('E15 frozen scrolled, then unfrozen');
  assert.deepEqual([s.freeze.rows, s.freeze.cols, s.scrolledTo], [0, 0, at('C10')]);
});
