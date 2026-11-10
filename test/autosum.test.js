import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { rangeName } from '../src/ooxml/refs.js';
import { autoSum } from '../src/ooxml/autosumops.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';

// AutoSum, each case as Excel did it (the ribbon's AutoSum, Average, Count Numbers, Max and Min run with
// CommandBars.ExecuteMso in a visible Excel: tools/excel-compare/excel-autosum.ps1 on autosum-cases.json, from
// autosum-cases.mjs, giving test/autosum-excel.json): every cell of the case's range (formula, value, number format)
// and the selection after compared with Excel's.
const cases = readCases('tools/excel-compare/autosum-cases.json', 'test/autosum-excel.json');

for (const c of cases) {
  test(`autosum: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si } = built;
    const step = autoSum(wb, si, area(c.select), c.fn);
    assert.ok(!step.edit, 'committed, not left being edited');
    wb.apply(si, step.contents);
    if (step.styles.length) wb.setStyles(si, step.styles);
    wb.calculate();
    saveCase(c, built);
    assert.equal(rangeName(step.area), c.excel.selection, 'the selection after');
    assert.deepEqual(cellsIn(wb, si, c.read, { results: true }), excelCells(c.excel, { results: true }));
  });
}

test('autosum: one cell, or nothing to sum, left being edited with Excel\'s proposal', () => {
  const { wb, si } = build({ cells: { B1: 'Qty', B2: '1', B3: '2', C1: 'When', C2: '1/1/2024', C3: '1/2/2024' } });
  assert.deepEqual(autoSum(wb, si, area('B4'), 'sum').edit, { r: 3, c: 1, text: '=SUM(B2:B3)' });
  assert.deepEqual(autoSum(wb, si, area('C4:D4'), 'average').edit, { r: 3, c: 2, text: '=AVERAGE()' }, 'dates are not taken');
});
