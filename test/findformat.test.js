import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { build, area } from './steps.js';
import { formatEach } from '../src/ooxml/formatops.js';
import { findAll, replaceChanges } from '../src/ooxml/findops.js';
import { formatFromCell, resolvedFormat } from '../src/ooxml/findformat.js';
import { refName } from '../src/ooxml/refs.js';

// Excel's cells (tools/excel-compare/excel-findformat.ps1) made here with the same formats, and Find Format's matches
// compared with what Range.Find found there (test/find-format2-excel.json).
const EXCEL = JSON.parse(fs.readFileSync('test/find-format2-excel.json', 'utf8'));
const ACCENT1 = { theme: 4 };
const ACCENT2 = { theme: 5 };

function sheetOf(column) {
  const cells = {};
  const n = { A: 7, B: 6, C: 6, D: 10, E: 6 }[column];
  for (let i = 1; i <= n; i++) cells[`${column}${i}`] = column === 'C' ? '5' : column === 'E' ? '7' : 'x';
  const { wb, si } = build({ cells });
  const fmt = (ref, change) => wb.setStyles(si, formatEach(wb, si, area(ref), change));
  if (column === 'A') {
    fmt('A2', { font: { color: { rgb: 'FF000000' } } });
    fmt('A4', { font: { color: { rgb: 'FFFF0000' } } });
    fmt('A5', { font: { color: ACCENT1 } });
    fmt('A6', { font: { color: { rgb: 'FF156082' } } });
    fmt('A7', { font: { color: { theme: 4, tint: 0.39997558519241921 } } });
  } else if (column === 'B') {
    fmt('B2', { fill: { rgb: 'FFFFFF00' } });
    fmt('B3', { fill: ACCENT2 });
    fmt('B4', { fill: { rgb: 'FFE97132' } });
    fmt('B5', { fill: { theme: 5, tint: 0.59999389629810485 } });
    fmt('B6', { fill: { pattern: 'darkGrid', fg: { rgb: 'FFFF0000' } } });
  } else if (column === 'C') {
    fmt('C2', { numFmt: '0.00' });
    fmt('C3', { numFmt: '#,##0.00' });
    fmt('C4', { font: { name: 'Arial' } });
    fmt('C5', { font: { size: 14 } });
    fmt('C6', { font: { name: 'Arial', size: 14 } });
  } else if (column === 'D') {
    fmt('D2', { border: { bottom: { style: 'thin' } } });
    fmt('D3', { border: { bottom: { style: 'thick' } } });
    fmt('D4', { border: { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } } });
    fmt('D5', { border: { bottom: { style: 'thin', color: { rgb: 'FFFF0000' } } } });
    fmt('D6', { align: { horizontal: 'center' } });
    fmt('D7', { align: { wrap: true } });
    fmt('D8', { align: { horizontal: 'center', wrap: true } });
    fmt('D9', { align: { horizontal: 'left', indent: 2 } });
    fmt('D10', { protection: { locked: false } });
  }
  return { wb, si };
}

const found = (wb, si, f) => findAll(wb, si, '', { findFormat: f }).map((h) => refName(h.r, h.c));
/** Excel's list starts after A1, so A1 comes last there: the same cells in sheet order. */
const excel = (name) => {
  const x = EXCEL.find((e) => e.name === name);
  assert.ok(x, name);
  return [...x.found].sort((a, b) => area(a).r0 - area(b).r0 || area(a).c0 - area(b).c0);
};

test('Find Format: colours resolved, a theme colour and its RGB alike, a tint not; automatic ink is black', () => {
  const A = sheetOf('A');
  assert.deepEqual(found(A.wb, A.si, { color: 'FF000000' }), excel('ink rgb black'));
  assert.deepEqual(found(A.wb, A.si, { color: 'FFFF0000' }), excel('ink red'));
  assert.deepEqual(found(A.wb, A.si, { color: 'FF156082' }), excel('ink accent 1 rgb'));
  assert.deepEqual(excel('ink theme accent 1'), excel('ink accent 1 rgb'));
  const B = sheetOf('B');
  assert.deepEqual(found(B.wb, B.si, { fill: 'FFFFFF00' }), excel('fill yellow'));
  assert.deepEqual(found(B.wb, B.si, { fill: 'FFE97132' }), excel('fill accent 2 rgb'));
  assert.deepEqual(found(B.wb, B.si, { fill: 'none' }), excel('fill none'));
});

test('Find Format: number formats, font name and size, alignment and protection as Excel matches them', () => {
  const C = sheetOf('C');
  assert.deepEqual(found(C.wb, C.si, { numFmt: '0.00' }), excel('numfmt 0.00'));
  assert.deepEqual(found(C.wb, C.si, { numFmt: 'General' }), excel('numfmt general'));
  assert.deepEqual(found(C.wb, C.si, { fontName: 'Arial' }), excel('font arial'));
  assert.deepEqual(found(C.wb, C.si, { fontSize: 14 }), excel('font size 14'));
  const D = sheetOf('D');
  assert.deepEqual(found(D.wb, D.si, { horizontal: 'center' }), excel('align centre'));
  assert.deepEqual(found(D.wb, D.si, { wrap: true }), excel('align wrap'));
  assert.deepEqual(found(D.wb, D.si, { indent: 2 }), excel('align indent 2'));
  assert.deepEqual(found(D.wb, D.si, { locked: false }), excel('unlocked'));
  assert.deepEqual(found(D.wb, D.si, { border: { bottom: { style: 'thin', color: 'FFFF0000' } } }), excel('border bottom red'));
  // (COM's FindFormat borders without a colour found nothing at all in Excel, not even its own thin line: not a
  // measurement of the dialog. A line asked for without a colour takes it in any colour here.)
  assert.deepEqual(found(D.wb, D.si, { border: { bottom: { style: 'thin' } } }), ['D2', 'D4', 'D5']);
});

test('Choose Format From Cell: every property of the cell, so only cells that look the same', () => {
  const { wb, si } = build({ cells: { E1: '7', E2: '7', E3: '7', E4: '7', E5: '7', E6: '7' } });
  const whole = { font: { b: true, name: 'Arial', size: 12, color: { rgb: 'FFFF0000' } }, fill: { rgb: 'FFFFFF00' }, numFmt: '0.00', align: { horizontal: 'center' }, border: { bottom: { style: 'thin' } } };
  for (const ref of ['E1', 'E2', 'E3', 'E4', 'E6']) wb.setStyles(si, formatEach(wb, si, area(ref), whole));
  wb.setStyles(si, formatEach(wb, si, area('E3'), { font: { i: true } }));
  wb.setStyles(si, formatEach(wb, si, area('E4'), { border: { bottom: null } }));
  wb.setStyles(si, formatEach(wb, si, area('E6'), { font: { color: { rgb: 'FF0000FF' } } }));
  const f = formatFromCell(wb, si, 0, 4);
  assert.deepEqual(found(wb, si, f), ['E1', 'E2']);
  assert.equal(f.fontName, 'Arial');
  assert.equal(f.border.bottom, 'thin:000000');
  // (A plain cell's format: the plain cells, whatever they hold.)
  assert.deepEqual(found(wb, si, formatFromCell(wb, si, 4, 4)), ['E5']);
  // (Replace Format from a cell: the cell takes that format whole, theme colours and all.)
  const steps = replaceChanges(wb, findAll(wb, si, '', { findFormat: formatFromCell(wb, si, 4, 4) }), '', '', { findFormat: formatFromCell(wb, si, 4, 4), replaceFormat: formatFromCell(wb, si, 0, 4) });
  assert.equal(steps.count, 1);
  wb.setStyles(si, steps[0].styles);
  assert.deepEqual(resolvedFormat(wb, wb.sheet(si).cells.get(4).get(4).s), resolvedFormat(wb, wb.sheet(si).cells.get(0).get(4).s));
});
