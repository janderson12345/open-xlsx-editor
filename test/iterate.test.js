import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
import { iteration, setIteration, calculateNow } from '../src/ooxml/calcmode.js';

// (Circular references as a hidden Excel worked them out: probe iter.ps1 and iter2.ps1, 30 cases and an 11-step
// sequence, all the same here. Values are Excel's, to the last bit.)
function book(on, count = 100) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'It');
  if (on) setIteration(wb, { on: true, count });
  return { wb, si };
}
const v = (wb, si, ref) => { const { r, c } = parseRef(ref); return wb.sheet(si).cells.get(r)?.get(c)?.v; };

test('iterative calculation off: a circle is left, the formula that closes it shows 0, as in Excel', () => {
  const { wb, si } = book(false);
  wb.setFormula(si, 'A1', 'B1+1');
  const { circular } = wb.apply(si, [{ ...parseRef('B1'), formula: 'A1*0.5' }]);
  assert.deepEqual([v(wb, si, 'A1'), v(wb, si, 'B1'), circular.length], [1, 0, 2]);
  wb.setFormula(si, 'D1', 'D1+1');
  calculateNow(wb);
  assert.equal(v(wb, si, 'D1'), 0);
});

test('iterative calculation on: passes row by row until nothing changes by 0.001, or the most iterations', () => {
  const { wb, si } = book(true);
  wb.setFormula(si, 'A1', '1+A1/2');
  assert.equal(v(wb, si, 'A1'), 1.9990234375);
  calculateNow(wb);
  assert.equal(v(wb, si, 'A1'), 1.99951171875);
  // (B1 before A2: rows first; B1 = A2+1, A2 = B1*0.5.)
  const rows = book(true);
  rows.wb.setFormula(rows.si, 'B1', 'A2+1');
  rows.wb.setFormula(rows.si, 'A2', 'B1*0.5');
  assert.deepEqual([v(rows.wb, rows.si, 'B1'), v(rows.wb, rows.si, 'A2')], [1.9990234375, 0.99951171875]);
  const few = book(true, 10);
  few.wb.setFormula(few.si, 'A1', 'A1+1');
  assert.equal(v(few.wb, few.si, 'A1'), 10);
  // (Any change goes round again: =A1+1 grows by the most iterations each time.)
  few.wb.setCell(few.si, 'Z9', 7);
  assert.equal(v(few.wb, few.si, 'A1'), 20);
  const cash = book(true);
  cash.wb.setCell(cash.si, 'B1', 100);
  cash.wb.setFormula(cash.si, 'C1', 'A1*0.1');
  cash.wb.setFormula(cash.si, 'A1', 'B1+C1');
  assert.deepEqual([v(cash.wb, cash.si, 'A1'), v(cash.wb, cash.si, 'C1')], [111.1111, 11.11111]);
});

test('iterative calculation in the file: iterate="1", iterateCount and iterateDelta only when not the defaults', () => {
  const { wb, si } = book(true);
  const calcPr = (b) => strFromU8(unzipSync(b.toBytes())['xl/workbook.xml']).match(/<calcPr[^>]*\/>/)?.[0];
  assert.equal(calcPr(wb), '<calcPr calcId="191029" iterate="1"/>');
  setIteration(wb, { on: true, count: 10, delta: 0.01 });
  assert.equal(calcPr(wb), '<calcPr calcId="191029" iterate="1" iterateCount="10" iterateDelta="0.01"/>');
  // (Opened again, it is on: a circle typed there goes round.)
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(iteration(again), { on: true, count: 10, delta: 0.01 });
  again.setFormula(si, 'A1', 'A1+1');
  assert.equal(v(again, si, 'A1'), 10);
  setIteration(again, { on: false });
  assert.equal(calcPr(again), '<calcPr calcId="191029" iterateCount="10" iterateDelta="0.01"/>');
});

test("iterative calculation turned on works the circles out, as Excel's did; changed while on, round again; off, left", () => {
  // (Excel over COM: =A1+1 left out at 0; Iteration on, MaxIterations 50: 50; MaxIterations 20: 70; off: 70.)
  const { wb, si } = book(false);
  wb.setFormula(si, 'A1', 'A1+1');
  assert.equal(v(wb, si, 'A1'), 0);
  setIteration(wb, { on: true, count: 50 });
  assert.equal(v(wb, si, 'A1'), 50);
  setIteration(wb, { on: true, count: 20 });
  assert.equal(v(wb, si, 'A1'), 70);
  setIteration(wb, { on: false });
  assert.equal(v(wb, si, 'A1'), 70);
});
