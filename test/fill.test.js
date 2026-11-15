import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fillCopy, autoFill, fillWay } from '../src/ooxml/fillops.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';

// Fill and the fill handle, each case as Excel did it over COM (Range.FillDown, FillRight, FillUp, FillLeft and
// Range.AutoFill: tools/excel-compare/excel-steps.ps1 on fill-cases.json, from fill-cases.mjs, giving
// test/fill-excel.json): the same cells filled here, and every cell of A1:H12 (formula and value, number format,
// bold, fill) compared with Excel's.
const cases = readCases('tools/excel-compare/fill-cases.json', 'test/fill-excel.json');

function run(c, { wb, si }) {
  for (const st of c.steps) {
    const steps = st.do === 'fill' ? fillCopy(wb, si, area(st.range), st.dir) : autoFill(wb, si, area(st.from), area(st.to), st.type || 'default');
    for (const x of steps) {
      if (x.contents.length) wb.apply(x.si, x.contents);
      if (x.styles.length) wb.setStyles(x.si, x.styles);
    }
  }
  wb.calculate();
}

// (Auto Fill Options for dates, Fill Days ... Fill Years: fill2-cases.json, test/fill2-excel.json.)
cases.push(...readCases('tools/excel-compare/fill2-cases.json', 'test/fill2-excel.json'));

for (const c of cases) {
  test(`fill: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    run(c, built);
    saveCase(c, built);
    assert.deepEqual(cellsIn(built.wb, built.si, c.read, { results: true }), excelCells(c.excel, { results: true }));
  });
}

test('fill: the way a fill handle was dragged', () => {
  const from = area('B2:C3');
  assert.equal(fillWay(from, area('B2:C9')), 'down');
  assert.equal(fillWay(from, area('B1:C3')), 'up');
  assert.equal(fillWay(from, area('B2:F3')), 'right');
  assert.equal(fillWay(from, area('A2:C3')), 'left');
  assert.equal(fillWay(from, area('B2:D9')), null);
});

test("fill: a double click on the handle fills down beside the next column's data, as Excel's documentation has it", async () => {
  const fillops = await import('../src/ooxml/fillops.js');
  const box = (a) => (a ? { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 } : null);
  const doubleClickFill = (sheet, sel) => fillops.doubleClickFill(sheet, sel);
  const { wb, si } = build({ cells: { A1: '1', A2: '2', A3: '3', A4: '4', B1: '=A1*2' } });
  const sheet = wb.sheet(si);
  assert.deepEqual(doubleClickFill(sheet, area('B1')), box(area('B1:B4')), 'down as far as column A goes');
  assert.equal(doubleClickFill(sheet, area('A1:A4')), null, 'nothing beside below it: nothing');
  const right = build({ cells: { B1: 'x', C1: '1', C2: '2', C3: '3' } });
  assert.deepEqual(doubleClickFill(right.wb.sheet(right.si), area('B1')), box(area('B1:B3')), 'the column right of it when the left is empty');
  const stop = build({ cells: { A1: '1', A2: '2', A3: '3', A4: '4', A5: '5', B1: 'a', B4: 'kept' } });
  assert.deepEqual(doubleClickFill(stop.wb.sheet(stop.si), area('B1')), box(area('B1:B3')), 'stopping before a cell already filled');
});

test('fill: custom lists (fillLists) go round as the built-in day and month names do', () => {
  const { wb, si } = build({ cells: { A1: 'east', B1: 'North', B2: 'South' } });
  wb.fillLists = [['North', 'East', 'South', 'West']];
  const run = (from, to) => { for (const x of autoFill(wb, si, area(from), area(to))) wb.apply(x.si, x.contents); };
  run('A1', 'A1:A4');
  assert.deepEqual([1, 2, 3].map((r) => wb.sheet(si).cells.get(r).get(0).v), ['south', 'west', 'north']);
  run('B1:B2', 'B1:B5');
  assert.deepEqual([2, 3, 4].map((r) => wb.sheet(si).cells.get(r).get(1).v), ['North', 'South', 'North']);
});
