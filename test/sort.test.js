import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { rangeName, parseRef } from '../src/ooxml/refs.js';
import { sortArea, sortChanges, guessHeader, compareText, dataBeside, SORT_MESSAGES } from '../src/ooxml/sortops.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { entryChanges } from '../src/ooxml/entry.js';

// Sort, each case as Excel did it over COM (Worksheet.Sort as its Sort box applies one; Range.Sort with xlGuess for
// the cases marked legacy, as its A to Z button guesses headers: tools/excel-compare/excel-steps.ps1 on
// sort-cases.json, from sort-cases.mjs, giving test/sort-excel.json): the range sorted, every cell read (formula,
// value, format), the merged ranges, the other sheet's formula and the defined name compared with Excel's, or
// Excel's refusal in its words.
const cases = readCases('tools/excel-compare/sort-cases.json', 'test/sort-excel.json');
// (Sort On Cell Color and Font Color, and a custom list as the order: sort2-cases.json, test/sort2-excel.json.)
cases.push(...readCases('tools/excel-compare/sort2-cases.json', 'test/sort2-excel.json'));
const argbOf = (n) => `FF${[n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
// (Cells left out: the array's second cell, whose result the sheets built here keep only in its first.)
const KNOWN = { 'part of an array': ['B3'] };

/** A case's sort done here, as the Sort box or the A to Z button would do it. */
function run(c, built) {
  const { wb, si } = built;
  for (const st of c.steps) {
    const sheet = wb.sheet(si);
    const a = st.region ? sortArea(sheet, area(st.range)) : area(st.range);
    assert.equal(rangeName(a), c.excel.region, 'the range sorted');
    // (The Sort object's own guess was seen never to take headers; Range.Sort's is the A to Z button's.)
    const header = st.header === 'yes' || (st.header !== 'no' && st.legacy && guessHeader(wb, si, a));
    const keys = st.keys.map((k) => {
      const p = parseRef(k.key);
      return { at: st.leftToRight ? p.r : p.c, desc: k.order === 'desc', ...(k.on ? { on: k.on, color: argbOf(k.color) } : {}), ...(k.customOrder ? { customOrder: k.customOrder.split(',') } : {}) };
    });
    const step = sortChanges(wb, si, a, { keys, header, matchCase: !!st.matchCase, leftToRight: !!st.leftToRight, textAsNumbers: st.keys.some((k) => k.textAsNumbers) });
    if (step.error) return step.error;
    if (step.contents.length) wb.apply(si, step.contents);
    if (step.styles.length) wb.setStyles(si, step.styles);
  }
  wb.calculate();
  return null;
}

for (const c of cases) {
  test(`sort: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si, s2, name } = built;
    const refused = run(c, built);
    assert.equal(refused, c.excel.refused ? c.excel.error : null);
    saveCase(c, built);
    const mine = cellsIn(wb, si, c.read, { results: true });
    const theirs = excelCells(c.excel, { results: true });
    for (const ref of KNOWN[c.name] || []) { delete mine[ref]; delete theirs[ref]; }
    assert.deepEqual(mine, theirs);
    assert.deepEqual(wb.sheet(si).merges.map(rangeName).sort(), [...(c.excel.merges || [])].sort());
    for (const [ref, text] of Object.entries(c.excel.sheet2 || {})) {
      const { r, c: col } = parseRef(ref);
      assert.equal(`=${wb.sheet(s2).cells.get(r)?.get(col)?.f}`, text.replace(/Sheet1!/g, `${name}!`), ref);
    }
    for (const [n, text] of Object.entries(c.excel.names || {})) {
      assert.equal(wb.definedNames.find((d) => d.name === n).text, text.slice(1).replace(/Sheet1!/g, `${name}!`), n);
    }
  });
}

test('sort: text compared as Excel compares it, the refusals in its words', () => {
  assert.ok(compareText('coop', 'co-op') < 0 && compareText('co-op', 'cop') < 0, 'hyphens passed over, after the word without');
  assert.equal(compareText('ab', 'AB'), 0, 'capitals alike unless asked');
  assert.ok(compareText('ab', 'Ab', { matchCase: true }) < 0, 'small letters first when asked');
  assert.ok(compareText('e', 'é') < 0 && compareText('é', 'f') < 0, 'an accent after the plain letter');
  assert.ok(compareText('~x', '10') < 0 && compareText('10', '2') < 0 && compareText('2', 'a') < 0, 'signs, then digits one by one, then letters');
  assert.equal(SORT_MESSAGES.merged, 'To do this, all the merged cells need to be the same size.');
  const { wb, si } = build({ cells: { A1: 'Name', B1: 'Qty', A2: 'c', B2: '3', A3: 'a', B3: '1', D1: 'apart' } });
  const sheet = wb.sheet(si);
  assert.equal(sortArea(sheet, area('F9')), null, 'an empty cell with nothing around it: nothing to sort');
  assert.equal(rangeName(dataBeside(sheet, area('A1:A3'))), 'A1:B3', 'one column of a table selected: the table beside it found');
  assert.equal(dataBeside(sheet, area('A1:B3')), null, 'the whole table selected: nothing left behind');
});

test("a sort through a dynamic array's spill is refused, as Excel refuses it (measured): any part, its formula's cell in or not", () => {
  // (Excel, C1 =SEQUENCE(6,1,10,-1) spilling to C6 beside A1:B6: A1:C6, A2:C6, C1:C6, C2:C4 by any column, "You can't
  // change part of an array.", nothing changed; A1:B6 sorts.)
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Spill');
  for (let r = 0; r < 6; r++) { wb.setCell(si, { r, c: 0 }, 6 - r); wb.setCell(si, { r, c: 1 }, `w${r + 1}`); }
  wb.apply(si, [entryChanges(wb, si, 0, 2, '=SEQUENCE(6,1,10,-1)').content]);
  assert.deepEqual(wb.sheet(si).cells.get(0).get(2).spill, { r0: 0, c0: 2, r1: 5, c1: 2 });
  const sort = (ref, at) => sortChanges(wb, si, area(ref), { keys: [{ at, desc: false }] }).error ?? 'sorted';
  assert.deepEqual([sort('A1:C6', 0), sort('A2:C6', 0), sort('C1:C6', 2), sort('C2:C4', 2), sort('A1:C6', 2)], Array(5).fill(SORT_MESSAGES.array));
  assert.equal(sort('A1:B6', 0), 'sorted');
});
