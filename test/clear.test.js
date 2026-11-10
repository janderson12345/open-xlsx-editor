import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { rangeName } from '../src/ooxml/refs.js';
import { clearChanges, CLEAR_REFUSAL } from '../src/ooxml/clearops.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';

// Home > Clear, each case as Excel did it over COM (Range.Clear, ClearContents, ClearFormats:
// tools/excel-compare/excel-steps.ps1 on clear-cases.json, from clear-cases.mjs, giving test/clear-excel.json): every
// cell of A1:F10 and the merged ranges compared with Excel's, or Excel's refusal.
const cases = readCases('tools/excel-compare/clear-cases.json', 'test/clear-excel.json');

for (const c of cases) {
  test(`clear: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si } = built;
    for (const st of c.steps) {
      const x = clearChanges(wb, si, area(st.range), st.what);
      if (c.excel.refused) { assert.equal(x.error, CLEAR_REFUSAL); assert.equal(c.excel.error, CLEAR_REFUSAL); continue; }
      if (x.contents.length) wb.apply(si, x.contents);
      wb.setMerges(si, x.merges);
      if (x.styles.length) wb.setStyles(si, x.styles);
    }
    wb.calculate();
    saveCase(c, built);
    assert.deepEqual(cellsIn(wb, si, c.read, { results: true }), excelCells(c.excel, { results: true }));
    assert.deepEqual(wb.sheet(si).merges.map(rangeName).sort(), [...(c.excel.merges || [])].sort());
  });
}
