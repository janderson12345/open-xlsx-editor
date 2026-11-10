import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseRef, rangeName } from '../src/ooxml/refs.js';
import { shiftCells, shapeShift, ROWCOL_MESSAGES } from '../src/ooxml/rowcol.js';
import { SheetError } from '../src/ooxml/sheetops.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';

// Inserting and deleting cells, each case as Excel did it over COM (tools/excel-compare/excel-steps.ps1 on
// shift-cases.json, from shift-cases.mjs, giving test/shift-excel.json): the same sheet built here, the same step
// done, and every cell of A1:L26 (formula, value, format, fill, bold, border), the merged ranges, the next sheet's
// formulas and the defined names compared with Excel's.
const cases = readCases('tools/excel-compare/shift-cases.json', 'test/shift-excel.json');
// (One reference Excel rewrites its own way, not copied: =SUM(B3:B3), B3 hidden in A3:B3, becomes =SUM(A3:A3) when
// cells are inserted at B3 and the merged range widens. Here it moves with B3: =SUM(D3:D3). docs/ROWCOL.md.)
const KNOWN = { 'merge across the edge, insert right': ['H12'] };

/** The steps of a case done here. */
function run(c, built) {
  const { wb, si } = built;
  for (const st of c.steps) {
    const sel = area(st.range);
    const axis = st.shift ? (st.shift === 'down' || st.shift === 'up' ? 'row' : 'col') : shapeShift(sel);
    shiftCells(wb, si, sel, st.do === 'insert' ? 1 : -1, axis, { unmerge: true });
  }
}

for (const c of cases) {
  const want = c.excel;
  test(`cells: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si, s2, name } = built;
    if (want.refused) {
      assert.throws(() => run(c, built), SheetError);
      return;
    }
    run(c, built);
    saveCase(c, built);
    const mine = cellsIn(wb, si, c.read);
    const theirs = excelCells(want);
    for (const ref of KNOWN[c.name] || []) { delete mine[ref]; delete theirs[ref]; }
    assert.deepEqual(mine, theirs);
    assert.deepEqual(wb.sheet(si).merges.map(rangeName).sort(), [...(want.merges || [])].sort());
    for (const [ref, text] of Object.entries(want.sheet2 || {})) {
      const { r, c: col } = parseRef(ref);
      assert.equal(`=${wb.sheet(s2).cells.get(r)?.get(col)?.f}`, text.replace(/^=@/, '=').replace(/Sheet1!/g, `${name}!`), ref);
    }
    for (const [n, text] of Object.entries(want.names || {})) {
      if (!c.setup.names[n]) continue;
      assert.equal(wb.definedNames.find((d) => d.name === n).text, text.slice(1).replace(/Sheet1!/g, `${name}!`), n);
    }
  });
}

test('cells: Excel\'s question before a shift that splits merged cells, its refusals in its words', () => {
  const setup = (name) => build(cases.find((x) => x.name === name).setup);
  const merged = setup('merge across the edge, insert down');
  assert.throws(() => shiftCells(merged.wb, merged.si, area('B3:C4'), 1, 'row'), (e) => e.ask && e.message === ROWCOL_MESSAGES.unmerge);
  const full = setup('filled last row, insert down');
  assert.throws(() => shiftCells(full.wb, full.si, area('B3'), 1, 'row'), (e) => e.message === ROWCOL_MESSAGES.offSheet);
  const arr = setup('array across the rows, insert right');
  assert.throws(() => shiftCells(arr.wb, arr.si, area('B3:C4'), 1, 'col'), (e) => e.message === ROWCOL_MESSAGES.array);
});
