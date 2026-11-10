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
