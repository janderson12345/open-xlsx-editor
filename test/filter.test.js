import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { serializeXml, kid } from '../src/ooxml/xml.js';
import { filterArea, filterChoices } from '../src/ooxml/filterops.js';
import { readCases, area, build, saveCase } from './steps.js';

// AutoFilter, each case as Excel did it (Range.AutoFilter over COM, then the workbook Excel saved:
// tools/excel-compare/excel-filter.ps1 on filter-cases.json, from filter-cases.mjs, giving test/filter-excel.json): the
// rows hidden, whether the filter is on and its range, the autoFilter element and the _xlnm._FilterDatabase name as
// Excel writes them (its random xr:uid aside).
const cases = readCases('tools/excel-compare/filter-cases.json', 'test/filter-excel.json');

/** A case's steps done here, as the filter buttons and Ctrl+Shift+L would do them. */
function run(c, { wb, si }) {
  for (const st of c.steps) {
    const sheet = wb.sheet(si);
    if (st.do === 'on') {
      if (sheet.autoFilter) wb.setAutoFilter(si, null);
      else wb.setAutoFilter(si, { ...filterArea(sheet, area(st.cell)).area, columns: new Map() });
      continue;
    }
    const f = sheet.autoFilter;
    const columns = new Map(f.columns);
    const col = f.c0 + st.field - 1;
    if (st.do === 'clear') columns.delete(col);
    else columns.set(col, { values: st.values.filter((v) => v !== '='), blank: st.values.includes('='), other: false });
    wb.setAutoFilter(si, { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1, columns });
  }
}
const tidy = (xml) => xml && xml.replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/ xr:uid="[^"]*"/g, '');

for (const c of cases) {
  test(`filter: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si, name } = built;
    run(c, built);
    saveCase(c, built);
    const sheet = wb.sheet(si);
    const hidden = [...sheet.rows].filter(([, row]) => row.hidden).map(([r]) => r + 1).sort((a, b) => a - b);
    assert.deepEqual(hidden, c.excel.hidden, 'the rows hidden');
    assert.equal(!!sheet.autoFilter, c.excel.filtering, 'on or off');
    const el = kid(wb._part(si).doc.documentElement, 'autoFilter');
    assert.equal(tidy(el ? serializeXml(el) : null), tidy(c.excel.autoFilter), 'the autoFilter element');
    const db = wb.definedNames.find((d) => d.name === '_xlnm._FilterDatabase' && d.sheet === si);
    const want = /hidden="1">([^<]*)</.exec(c.excel.filterName)?.[1].replace('Sheet1!', `${name}!`);
    assert.equal(db?.text, want, 'the _xlnm._FilterDatabase name');
    assert.equal(db?.el.getAttribute('hidden'), '1');
  });
}

test('filter: the list of a column, as the dropdown shows it', () => {
  const c = cases.find((x) => x.name === 'turned on from one cell');
  const { wb, si } = build(c.setup);
  wb.setAutoFilter(si, { ...filterArea(wb.sheet(si), area('B3')).area, columns: new Map() });
  const f = wb.sheet(si).autoFilter;
  assert.deepEqual(filterChoices(wb, si, f, 0).values.map((v) => v.text), ['apple', 'fig', 'kiwi', 'pear', 'plum'], 'one apple for apple and Apple, A to Z');
  assert.deepEqual(filterChoices(wb, si, f, 1).values.map((v) => v.text), ['1', '2', '2.5', '3', '10'], 'numbers by value');
  const teams = filterChoices(wb, si, f, 2);
  assert.deepEqual([teams.values.map((v) => v.text), teams.blanks], [['blue', 'green', 'red'], true], 'blanks offered last');
  wb.setAutoFilter(si, { ...f, columns: new Map([[2, { values: ['red'], blank: false, other: false }]]) });
  assert.deepEqual(filterChoices(wb, si, wb.sheet(si).autoFilter, 0).values.map((v) => v.text), ['apple'], 'only what the other columns let through');
  assert.deepEqual(filterChoices(wb, si, wb.sheet(si).autoFilter, 2).values.map((v) => v.checked), [false, false, true]);
});
