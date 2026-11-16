import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readCases, area, build } from './steps.js';
import { refName } from '../src/ooxml/refs.js';
import { lineSpans, shiftOrder, overlapping, MULTI_RANGE_MESSAGE, OVERLAP_MESSAGE } from '../src/ooxml/multiarea.js';
import { shiftCells } from '../src/ooxml/rowcol.js';
import { fillCopyAreas } from '../src/ooxml/fillops.js';
import { mergeAreasChange } from '../src/ooxml/mergeops.js';
import { findAll, replaceChanges } from '../src/ooxml/findops.js';
import { specialCells } from '../src/ooxml/special.js';
import { selectionStats } from '../src/ooxml/stats.js';
import { chartSeriesAreas } from '../src/ooxml/chartnew.js';
import { namesFromSelection } from '../src/ooxml/createnames.js';
import { selectionRefers, areasSelection, withinSelection } from '../src/editor/selection.js';

// Commands over a selection of several areas, as Excel did them (excel-multirange.ps1 on multirange-cases.json,
// giving test/multirange-excel.json): A1:F12 filled with row * 10 + column, the areas selected, the command run.
// Each case is done here as the editor does it (commands.js, commands-data.js, commands-lines.js: each area, all of
// them, or Excel's refusal) and A1:H14 compared with Excel's.
const cases = readCases('tools/excel-compare/multirange-cases.json', 'test/multirange-excel.json');
const byName = (name) => { const c = cases.find((x) => x.name === name); assert.ok(c?.excel, `Excel's answer for ${name}`); return c; };
const areas = (text) => text.split(',').map((t) => { const a = area(t.replace(/^([A-Z]+):([A-Z]+)$/, '$11:$21048576')); return { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }; });

/** The case's sheet: the grid of numbers, then its own cells and merges. */
function setUp(c) {
  const cells = {};
  for (let r = 1; r <= 12; r++) for (let k = 1; k <= 6; k++) cells[refName(r - 1, k - 1)] = String(r * 10 + k);
  const built = build({ cells: { ...cells, ...(c.setup?.cells || {}) } });
  // (Merged as Excel's Range.Merge merges: the other values gone.)
  if (c.setup?.merge) {
    const m = mergeAreasChange(built.wb, built.si, c.setup.merge.map((t) => area(t)), 'cells');
    built.wb.apply(built.si, m.contents);
    built.wb.setMerges(built.si, m.merges);
  }
  return built;
}
/** A1:H14 as values, { A1: 11, ... }, empty cells left out. */
function grid(wb, si) {
  const out = {};
  const sh = wb.sheet(si);
  for (let r = 0; r < 14; r++) for (let c = 0; c < 8; c++) { const v = sh.cells.get(r)?.get(c)?.v; if (v != null && v !== '') out[refName(r, c)] = v; }
  return out;
}
const excelGrid = (answer) => Object.fromEntries(Object.entries(answer.cells).filter(([, v]) => v.v != null && v.v !== '').map(([k, v]) => [k, v.v]));
const boxText = (answer) => (answer.boxes || []).map((b) => b.text).join(' | ');

const RUNS = {
  /** Fill Down or Right (commands-data.js fillEach). */
  fill(wb, si, list, dir) { for (const st of fillCopyAreas(wb, si, list, dir)) { wb.apply(si, st.contents); wb.setStyles(si, st.styles); } },
  /** Insert or Delete Sheet Rows and Columns (commands.js lines). */
  lines(wb, si, list, axis, dir) { for (const [a, b] of lineSpans(list, axis, dir > 0)) wb.changeLines(si, axis, a, dir * (b - a + 1)); },
  /** Insert or Delete cells, shifted (commands.js cells). */
  cells(wb, si, list, dir, axis) { if (overlapping(list)) return OVERLAP_MESSAGE; for (const a of shiftOrder(list, axis)) shiftCells(wb, si, a, dir, axis); return null; },
};
const GRID_CASES = [
  ['fill down, apart', (wb, si, l) => RUNS.fill(wb, si, l, 'down')],
  ['fill down, same rows', (wb, si, l) => RUNS.fill(wb, si, l, 'down')],
  ['fill right, apart', (wb, si, l) => RUNS.fill(wb, si, l, 'right')],
  ['fill down, same columns', (wb, si, l) => RUNS.fill(wb, si, l, 'down')],
  ['fill down, overlapping', (wb, si, l) => RUNS.fill(wb, si, l, 'down')],
  ['insert rows, two', (wb, si, l) => RUNS.lines(wb, si, l, 'row', 1)],
  ['insert rows, overlapping', (wb, si, l) => RUNS.lines(wb, si, l, 'row', 1)],
  ['delete rows, two', (wb, si, l) => RUNS.lines(wb, si, l, 'row', -1)],
  ['delete rows, overlapping', (wb, si, l) => RUNS.lines(wb, si, l, 'row', -1)],
  ['insert columns, two', (wb, si, l) => RUNS.lines(wb, si, l, 'col', 1)],
  ['delete columns, two', (wb, si, l) => RUNS.lines(wb, si, l, 'col', -1)],
  ['delete cells up, apart', (wb, si, l) => RUNS.cells(wb, si, l, -1, 'row')],
  ['delete cells up, same rows', (wb, si, l) => RUNS.cells(wb, si, l, -1, 'row')],
  ['delete cells up, same columns', (wb, si, l) => RUNS.cells(wb, si, l, -1, 'row')],
  ['delete cells up, columns partly shared', (wb, si, l) => RUNS.cells(wb, si, l, -1, 'row')],
  ['delete cells left, rows partly shared', (wb, si, l) => RUNS.cells(wb, si, l, -1, 'col')],
  ['insert cells down, apart', (wb, si, l) => RUNS.cells(wb, si, l, 1, 'row')],
  ['insert cells down, same rows', (wb, si, l) => RUNS.cells(wb, si, l, 1, 'row')],
  ['insert cells down, columns partly shared', (wb, si, l) => RUNS.cells(wb, si, l, 1, 'row')],
  ['insert cells right, same columns', (wb, si, l) => RUNS.cells(wb, si, l, 1, 'col')],
];
for (const [name, run] of GRID_CASES) {
  test(`several areas: ${name}, as Excel`, () => {
    const c = byName(name);
    const { wb, si } = setUp(c);
    run(wb, si, areas(c.select));
    assert.deepEqual(grid(wb, si), excelGrid(c.excel));
  });
}

test("several areas: cells over overlapping areas refused in Excel's words", () => {
  const c = byName('delete cells up, overlapping');
  assert.equal(c.excel.error, OVERLAP_MESSAGE);
  const { wb, si } = setUp(c);
  assert.equal(RUNS.cells(wb, si, areas(c.select), -1, 'row'), OVERLAP_MESSAGE);
  assert.deepEqual(grid(wb, si), excelGrid(c.excel));
});

test("several areas: Excel's refusals are the editor's words", () => {
  const refused = ['sort ascending, apart', 'sort ascending, same rows', 'sort box', 'custom sort', 'sort with whole columns', 'filter, autofilter id', 'filter, Filter id', 'remove duplicates', 'fill justify', 'group rows', 'subtotals'];
  for (const name of refused) assert.equal(boxText(byName(name).excel), MULTI_RANGE_MESSAGE, name);
  // (Text to Columns: COM's TextToColumns gives the same words; the ribbon's button has no idMso to run.)
  assert.equal(byName('text to columns').excel.error, MULTI_RANGE_MESSAGE);
  // (Insert Table, Format as Table and PivotTable greyed out: nothing to say.)
  for (const name of ['table insert', 'pivot insert', 'enabled FormatAsTableGallery']) assert.equal(byName(name).excel.enabled, false, name);
  // (And nothing changed by them.)
  for (const name of refused) assert.deepEqual(excelGrid(byName(name).excel), excelGrid(byName('borders outside').excel), name);
});

test('several areas: hide and size the lines of each, as Excel', () => {
  const lines = (list, axis) => lineSpans(list, axis).flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, k) => a + k + 1)).sort((x, y) => x - y);
  const hid = (name, axis) => { const c = byName(name); assert.deepEqual(lines(areas(c.select), axis), axis === 'row' ? c.excel.hiddenRows : c.excel.hiddenCols, name); };
  hid('hide rows', 'row');
  hid('hide columns', 'col');
  const rows = byName('row height set');
  assert.deepEqual(lines(areas(rows.select), 'row'), rows.excel.heights.map((h, k) => (h === 30 ? k + 1 : 0)).filter(Boolean));
  const cols = byName('column width set');
  assert.deepEqual(lines(areas(cols.select), 'col'), cols.excel.widths.map((w, k) => (w === 20 ? k + 1 : 0)).filter(Boolean));
  const fit = byName('autofit columns ribbon');
  assert.deepEqual(lines(areas(fit.select), 'col'), fit.excel.widths.map((w, k) => (w !== fit.excel.widths[1] ? k + 1 : 0)).filter(Boolean));
});

test('several areas: lines to insert and delete, as Excel took them', () => {
  const two = [{ r0: 1, c0: 0, r1: 3, c1: 1 }, { r0: 2, c0: 3, r1: 4, c1: 4 }];
  assert.deepEqual(lineSpans(two, 'row'), [[1, 4]]);
  assert.deepEqual(lineSpans(two, 'row', true), [[2, 4], [1, 1]]);
  // (Touching areas: each inserted above its own first row.)
  assert.deepEqual(lineSpans([{ r0: 1, c0: 0, r1: 2, c1: 0 }, { r0: 3, c0: 0, r1: 4, c1: 0 }], 'row', true), [[3, 4], [1, 2]]);
  assert.deepEqual(shiftOrder([{ r0: 1, c0: 0, r1: 2, c1: 1 }, { r0: 5, c0: 1, r1: 5, c1: 2 }], 'row').map((a) => a.r0), [5, 1]);
});

test('several areas: merge each, as Excel', () => {
  for (const [name, kind] of [['merge center, apart', 'center'], ['merge cells, apart', 'cells'], ['merge across, apart', 'across'], ['unmerge, two', 'unmerge'], ['merge, overlapping', 'cells']]) {
    const c = byName(name);
    const { wb, si } = setUp(c);
    const change = mergeAreasChange(wb, si, areas(c.select), kind);
    if (change) {
      wb.apply(si, change.contents);
      wb.setMerges(si, change.merges);
      wb.setStyles(si, change.styles);
    }
    const theirs = Object.values(c.excel.cells).map((x) => x.merge).filter(Boolean).sort();
    assert.deepEqual(wb.sheet(si).merges.map((m) => `${refName(m.r0, m.c0)}:${refName(m.r1, m.c1)}`).sort(), theirs, name);
    assert.deepEqual(grid(wb, si), excelGrid(c.excel), name);
  }
});

test('several areas: Find and Replace look in every area, as Excel', () => {
  const find = byName('find within');
  const { wb, si } = setUp(find);
  const opts = { areas: areas(find.select), areasSheet: si };
  assert.deepEqual(findAll(wb, si, '6', opts).map((x) => refName(x.r, x.c)).join(','), find.excel.result);
  const repl = byName('replace within');
  const built = setUp(repl);
  const steps = replaceChanges(built.wb, findAll(built.wb, built.si, '2', { areas: areas(repl.select), areasSheet: built.si }), '2', 'x', {});
  for (const st of steps) built.wb.apply(built.si, st.contents);
  assert.deepEqual(grid(built.wb, built.si), excelGrid(repl.excel));
  // (One cell selected, or a merged range alone: the whole sheet.)
  const sheet = wb.sheet(si);
  assert.equal(withinSelection(sheet, { ar: 0, ac: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, si), null);
  assert.deepEqual(withinSelection(sheet, areasSelection(areas(find.select)), si).areas, areas(find.select));
});

test('several areas: Go To Special looks in every area, as Excel', () => {
  for (const [name, choice] of [['special constants within', { kind: 'constants', types: 23 }], ['special blanks within', { kind: 'blanks' }]]) {
    const c = byName(name);
    const { wb, si } = setUp(c);
    const got = specialCells(wb, si, areas(c.select), { r: 1, c: 0 }, choice);
    const cellsOf = (list) => list.flatMap((a) => { const out = []; for (let r = a.r0; r <= a.r1; r++) for (let k = a.c0; k <= a.c1; k++) out.push(refName(r, k)); return out; }).sort();
    assert.deepEqual(cellsOf(got.areas), cellsOf(areas(c.excel.result)), name);
  }
});

test('several areas: the status bar sums every area, a shared cell twice, as Excel', () => {
  const c = byName('status sum');
  const { wb, si } = setUp(c);
  const s = selectionStats(wb, si, areasSelection(areas(c.select)));
  assert.equal(`sum ${s.sum} count ${s.count} avg ${s.average}`, c.excel.result);
});

test('several areas: a chart of every area, its series as Excel took them', () => {
  for (const name of ['chart, two columns apart', 'chart, apart']) {
    const c = byName(name);
    const { wb, si } = setUp(c);
    const got = chartSeriesAreas(wb, si, areas(c.select));
    const theirs = c.excel.charts[0].map((f) => f.match(/^=SERIES\([^,]*,[^,]*,([^,]+),\d+\)$/)[1].replace('Sheet1!', ''));
    assert.deepEqual(got.series.map((x) => x.val.f.replace(/^.*!/, '')), theirs, name);
  }
});

test('several areas: names refer to every area; Create from Selection names each, as Excel', () => {
  const c = byName('create names with text');
  const { wb, si } = setUp(c);
  const made = areas(c.select).flatMap((a) => namesFromSelection(wb, si, a, { top: true }));
  const theirs = c.excel.names.map((x) => x.split(' ')).map(([n, ref]) => [n, ref.replace('=Sheet1!', '')]).sort();
  assert.deepEqual(made.map((x) => [x.name, x.refersTo.replace(/^=.*!/, '')]).sort(), theirs);
  const sel = areasSelection(areas('A2:B3,D6:E7'));
  assert.equal(selectionRefers(wb, si, sel).replace(/=?[^=,!]*!/g, (m) => (m.startsWith('=') ? '=' : '')), '=$A$2:$B$3,$D$6:$E$7');
});

test('several areas: conditional formatting applies to every area, as Excel', () => {
  const c = byName('cond format om');
  assert.deepEqual(c.excel.conditions, [c.select]);
});

test('several areas: Print Selection prints each area on pages of its own, as Excel', () => {
  // (Excel's PDF of A2:B3 and D6:E7: two pages; print.js lays several areas out so, editing-more.mjs checks it.)
  assert.equal(Number(byName('print selection pages').excel.result), areas(byName('print selection pages').select).length);
});
