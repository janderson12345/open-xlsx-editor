import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { specialCells, cellsToAreas, firstShown, NO_CELLS } from '../src/ooxml/special.js';
import { parseRange, refName } from '../src/ooxml/refs.js';

// test/special-excel.json (tools/excel-compare/excel-special.ps1 on special-cases.mjs): what Excel's Go To Special
// found on test/fixtures/special.xlsx for each case: its ranges, or nothing ("No cells were found.").
const EXCEL = JSON.parse(fs.readFileSync('test/special-excel.json', 'utf8').replace(/^\uFEFF/, ''));
const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/special-cases.json', 'utf8')).cases;
const book = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/special.xlsx')));

/** The cells of an address ("A1:B2,D4"), as a sorted list of names; a whole row or column band kept as its range. */
function cells(address) {
  const out = [];
  for (const part of address.split(',')) {
    const a = parseRange(part);
    if ((a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) > 5000) { out.push(`${refName(a.r0, a.c0)}:${refName(a.r1, a.c1)}`); continue; }
    for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) out.push(refName(r, c));
  }
  return out.sort();
}
const address = (areas) => areas.map((a) => (a.r0 === a.r1 && a.c0 === a.c1 ? refName(a.r0, a.c0) : `${refName(a.r0, a.c0)}:${refName(a.r1, a.c1)}`)).join(',');

test('Go To Special finds the cells Excel finds, in each of its choices', () => {
  const wb = book();
  assert.equal(CASES.length, EXCEL.length);
  CASES.forEach((k, i) => {
    const want = EXCEL[i];
    assert.equal(want.name, k.name);
    const areas = k.select.split(',').map(parseRange);
    const active = parseRange(k.active);
    const got = specialCells(wb, 0, areas, { r: active.r0, c: active.c0 }, { kind: k.kind, types: k.types, all: k.levels === 'all' });
    if (want.address == null) assert.equal(got.error, NO_CELLS, k.name);
    else assert.deepEqual(cells(address(got.areas || [])), cells(want.address), k.name);
  });
});

test('cells are gathered into ranges: runs of a row, joined down when they span the same columns', () => {
  assert.equal(address(cellsToAreas(new Set(['0,0', '0,1', '1,0', '1,1', '1,3', '2,3', '4,0']))), 'A1:B2,D2:D3,A5');
});

test('the cell made active is the first found that is in sight: not in hidden column F or row 4', () => {
  const wb = book();
  const blanks = specialCells(wb, 0, [{ r0: 0, c0: 0, r1: 0, c1: 0 }], { r: 0, c: 0 }, { kind: 'blanks' });
  assert.deepEqual(firstShown(wb.sheet(0), blanks.areas), { r: 0, c: 6 }); // G1 (F1:F3 hidden)
  const e = specialCells(wb, 0, [parseRange('A1:E8')], { r: 0, c: 0 }, { kind: 'blanks' });
  assert.deepEqual(firstShown(wb.sheet(0), e.areas), { r: 4, c: 4 }); // E5 (E4 in hidden row 4)
});

test('Visible cells only on a whole sheet comes back as bands, not cells', () => {
  const got = specialCells(book(), 0, [{ r0: 0, c0: 0, r1: 0, c1: 0 }], { r: 0, c: 0 }, { kind: 'visible' });
  assert.equal(address(got.areas), 'A1:E3,G1:XFD3,A5:E1048576,G5:XFD1048576');
});
