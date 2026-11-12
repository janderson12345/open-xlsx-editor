import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { calcMode, setCalcMode, setCalcOnSave, calcOnSave, needsCalculation, calculateNow, calculateSheet, calculateBeforeSave } from '../src/ooxml/calcmode.js';
import { parseRef } from '../src/ooxml/refs.js';

const book = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
const value = (wb, si, ref) => { const { r, c } = parseRef(ref); return wb.sheet(si).cells.get(r)?.get(c)?.v; };
const calcPr = (wb) => strFromU8(unzipSync(wb.toBytes())['xl/workbook.xml']).match(/<calcPr[^>]*\/>/)?.[0] ?? null;

test('Manual calculation, as Excel does it (measured): a change waits, a formula entered is worked out, F9 and Shift+F9', () => {
  const wb = book();
  const calc = wb.sheetIndexOf('Calc');
  const my = wb.sheetIndexOf('My Data');
  wb.setCell(calc, 'H1', 5);
  wb.setFormula(calc, 'H2', 'H1*2');
  wb.setFormula(my, 'C1', 'Calc!H1*10');
  assert.equal(calcMode(wb), 'auto');
  setCalcMode(wb, 'manual');
  // (Excel: A1 5 to 7 in Manual leaves =A1*2 at 10 and another sheet's =Sheet1!A1*10 at 50.)
  wb.setCell(calc, 'H1', 7);
  assert.deepEqual([value(wb, calc, 'H2'), value(wb, my, 'C1'), needsCalculation(wb)], [10, 50, true]);
  // (A formula entered is worked out, from what the cells hold now: =A1*3 is 21, =A2+1 11, A2 still 10.)
  wb.setFormula(calc, 'H3', 'H1*3');
  wb.setFormula(calc, 'H4', 'H2+1');
  assert.deepEqual([value(wb, calc, 'H3'), value(wb, calc, 'H4')], [21, 11]);
  // (Calculate Sheet: that sheet's formulas, A2 14 and A4 15; the other sheet's waits, 50. Then F9: 70.)
  calculateSheet(wb, calc);
  assert.deepEqual([value(wb, calc, 'H2'), value(wb, calc, 'H4'), value(wb, my, 'C1')], [14, 15, 50]);
  calculateNow(wb);
  assert.deepEqual([value(wb, my, 'C1'), needsCalculation(wb)], [70, false]);
  // (Back to Automatic: what waits is worked out, 9 to 18.)
  wb.setCell(calc, 'H1', 9);
  setCalcMode(wb, 'auto');
  assert.equal(value(wb, calc, 'H2'), 18);
});

test('Calculation Options saved as Excel saves them: calcMode="manual", "autoNoTable", none; calcOnSave="0"', () => {
  const wb = book();
  setCalcMode(wb, 'manual');
  assert.equal(calcPr(wb), '<calcPr calcId="191029" calcMode="manual"/>');
  setCalcOnSave(wb, false);
  assert.equal(calcPr(wb), '<calcPr calcId="191029" calcMode="manual" calcOnSave="0"/>');
  setCalcOnSave(wb, true);
  setCalcMode(wb, 'autoNoTable');
  assert.equal(calcPr(wb), '<calcPr calcId="191029" calcMode="autoNoTable"/>');
  setCalcMode(wb, 'auto');
  assert.equal(calcPr(wb), '<calcPr calcId="191029"/>');
  // (Read back from the file.)
  setCalcMode(wb, 'manual');
  setCalcOnSave(wb, false);
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual([calcMode(again), calcOnSave(again)], ['manual', false]);
});

test('Saving a manual workbook works it out first, unless Recalculate workbook before saving is off', () => {
  const wb = book();
  const calc = wb.sheetIndexOf('Calc');
  wb.setCell(calc, 'H1', 5);
  wb.setFormula(calc, 'H2', 'H1*2');
  setCalcMode(wb, 'manual');
  wb.setCell(calc, 'H1', 6);
  setCalcOnSave(wb, false);
  assert.equal(calculateBeforeSave(wb), false);
  assert.equal(value(wb, calc, 'H2'), 10);
  setCalcOnSave(wb, true);
  assert.equal(calculateBeforeSave(wb), true);
  assert.equal(value(wb, calc, 'H2'), 12);
});
