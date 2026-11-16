import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { autoSum, autoSumAreas } from '../src/ooxml/autosumops.js';
import { applyAutoFilter } from '../src/ooxml/filterops.js';
import { createTable, tableAt } from '../src/ooxml/tableedit.js';
import { applyTableFilter } from '../src/ooxml/tablefilter.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';
import { rangeName } from '../src/ooxml/refs.js';

// More of AutoSum as Excel did it (excel-autosum2.ps1 on autosum2-cases.json, from autosum2-cases.mjs, giving
// test/autosum2-excel.json): a filtered list, a filtered and an unfiltered table, rows hidden by hand, several areas,
// text and a blank among the numbers selected, a blank column beside a block. Each cell's formula, value and format
// and the selection after compared with Excel's.
const cases = readCases('tools/excel-compare/autosum2-cases.json', 'test/autosum2-excel.json');
const range = (text) => area(text);

/** The case's sheet with its filter, table or hidden rows set up as Excel's was. */
function setUp(setup) {
  const built = build(setup);
  const { wb, si } = built;
  const f = setup.autofilter;
  if (f) {
    const a = range(f.range);
    // (Excel's Range.AutoFilter(field, "Apple"): that column's values, Apple alone.)
    const columns = f.field ? new Map([[a.c0 + f.field - 1, { values: [f.criteria], blank: false, other: false }]]) : new Map();
    applyAutoFilter(wb, si, { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1, columns });
  }
  if (setup.table) {
    const a = range(setup.table.range);
    createTable(wb, si, a, { headers: true });
    if (setup.tableFilter) {
      const t = tableAt(wb, si, a.r0, a.c0);
      applyTableFilter(wb, t, { r0: t.r0, c0: t.c0, r1: t.r1, c1: t.c1, columns: new Map([[t.c0 + setup.tableFilter.field - 1, { values: [setup.tableFilter.criteria], blank: false, other: false }]]) });
    }
  }
  if (setup.hideRows) {
    const [a, b = a] = setup.hideRows.split(':').map(Number);
    wb.setLines(si, 'row', a - 1, b - 1, { hidden: true });
  }
  return built;
}

for (const c of cases) {
  test(`autosum: ${c.name}, as Excel`, () => {
    assert.ok(c.excel && !c.excel.error, `Excel's answer for ${c.name}`);
    const built = setUp(c.setup);
    const { wb, si } = built;
    const areas = c.select.split(',').map(range);
    const step = autoSumAreas(wb, si, areas, c.fn);
    assert.ok(!step.edit, 'committed, not left being edited');
    wb.apply(si, step.contents);
    if (step.styles.length) wb.setStyles(si, step.styles);
    wb.calculate();
    saveCase(c, built);
    // (Excel lists the areas from the active one: B2:B7,E2:E5 for B2:B6 then E2:E4 selected.)
    const after = (step.areas || [step.area]).map(rangeName);
    assert.deepEqual([...after].sort(), c.excel.selection.split(',').sort(), 'the selection after');
    assert.deepEqual(cellsIn(wb, si, c.read, { results: true }), excelCells(c.excel, { results: true }));
  });
}

test('autosum on one cell in a filtered list or under a table: Excel’s SUBTOTAL proposal left being edited', () => {
  const built = setUp(cases.find((c) => c.name === 'filtered list: the blank row under it').setup);
  assert.deepEqual(autoSum(built.wb, built.si, range('B7'), 'sum').edit, { r: 6, c: 1, text: '=SUBTOTAL(9,B2:B6)' });
  assert.deepEqual(autoSum(built.wb, built.si, range('B7'), 'max').edit, { r: 6, c: 1, text: '=MAX(B2:B6)' });
  const t = setUp(cases.find((c) => c.name === 'a table under it, not filtered').setup);
  assert.deepEqual(autoSum(t.wb, t.si, range('C7'), 'sum').edit, { r: 6, c: 2, text: '=SUBTOTAL(109,C2:C6)' });
});

test('several blank rows under the numbers: the first opened for editing with the proposal, as Excel left B7:C9', () => {
  const built = setUp({ cells: cases[0].setup.cells });
  assert.deepEqual(autoSum(built.wb, built.si, range('B7:C9'), 'sum').edit, { r: 6, c: 1, text: '=SUM(B2:B6)' });
});
