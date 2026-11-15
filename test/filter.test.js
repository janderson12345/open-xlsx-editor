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

test('sorting while rows are filtered out: only the rows shown are sorted, as Excel (filter2-excel.json)', async () => {
  const fs = await import('node:fs');
  const { sortChanges } = await import('../src/ooxml/sortops.js');
  const EXCEL = JSON.parse(fs.readFileSync('test/filter2-excel.json', 'utf8').replace(/^﻿/, '')).cases;
  const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/filter2-cases.json', 'utf8'));
  for (const name of ['red rows sorted A to Z while filtered', 'red rows sorted by Qty, largest first, while filtered']) {
    const c = CASES.find((x) => x.name === name);
    const { wb, si } = build({ cells: c.setup.cells });
    wb.setAutoFilter(si, { r0: 0, c0: 0, r1: 6, c1: 2, columns: new Map([[2, { values: ['red'], blank: false, other: false }]]) });
    const sort = c.steps.find((s) => s.do === 'sort');
    const step = sortChanges(wb, si, { r0: 0, c0: 0, r1: 6, c1: 2 }, { keys: [{ at: sort.key === 'B1' ? 1 : 0, desc: sort.order === 'desc' }], header: true });
    wb.apply(si, step.contents);
    const values = {};
    for (const [r, line] of wb.sheet(si).cells) for (const [col, cell] of line) values[`${'ABC'[col]}${r + 1}`] = String(cell.v);
    assert.deepEqual(values, EXCEL.find((x) => x.name === name).values, name);
  }
});

test('dates filtered by year, month or day, written and applied as Excel 365 does (filter2-excel.json)', async () => {
  const fs = await import('node:fs');
  const EXCEL = JSON.parse(fs.readFileSync('test/filter2-excel.json', 'utf8').replace(/^﻿/, '')).cases;
  const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/filter2-cases.json', 'utf8'));
  for (const c of CASES.filter((x) => x.name.startsWith('dates'))) {
    const { wb, si } = build({ cells: c.setup.cells });
    const st = c.steps.find((s) => s.do === 'dates');
    const dates = st.groups.map(([level, text]) => { const [m, d, y] = text.split('/').map(Number); return { year: y, ...(level > 0 ? { month: m } : {}), ...(level > 1 ? { day: d } : {}) }; });
    wb.setAutoFilter(si, { r0: 0, c0: 0, r1: 6, c1: 1, columns: new Map([[0, { values: [], blank: false, dates, other: false }]]) });
    const x = EXCEL.find((e) => e.name === c.name);
    const hidden = [...wb.sheet(si).rows].filter(([, row]) => row.hidden).map(([r]) => r + 1).sort((a, b) => a - b);
    assert.deepEqual(hidden, x.hidden, `${c.name}: the rows hidden`);
    const el = kid(wb._part(si).doc.documentElement, 'autoFilter', wb._part(si).doc.documentElement.namespaceURI);
    const mine = serializeXml(el).replace(/ xmlns="[^"]*"/, '');
    assert.equal(mine, x.autoFilter.replace(/ xr:uid="[^"]*"/, ''), `${c.name}: Excel's XML`);
    // (And read back: the same groups.)
    assert.deepEqual(wb.sheet(si).autoFilter.columns.get(0).dates, dates);
  }
});
