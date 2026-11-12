import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { errorFlags } from '../src/ooxml/errorcheck.js';
import { parseRef } from '../src/ooxml/refs.js';

// Excel's error indicators for the same cells: a hidden Excel's Range.Errors(1 to 6), its default rules on.
const fresh = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  return { wb, si: wb.addSheet(wb.sheets.length) };
};
const flagsAt = (wb, si) => {
  const f = errorFlags(wb, si);
  return (ref) => { const { r, c } = parseRef(ref); return (f.get(`${r},${c}`) || []).join(','); };
};

test('numbers stored as text and two-digit-year text dates, as Excel flags them', () => {
  const { wb, si } = fresh();
  const texts = { A1: '123', A2: '456', A3: 'abc', A4: '12.5', A5: '1,234', A6: '1e3', A7: ' 7', A8: '50%', A10: '(5)', A11: '-3',
    A12: '1/2/2024', A13: '1/2/24', A14: 'Jan 2, 24', A15: 'TRUE', A17: '2-Jan-24', A18: '12:30', A19: '0012' };
  for (const [ref, v] of Object.entries(texts)) wb.setCell(si, ref, v);
  wb.setFormula(si, 'A16', '"5"');
  const at = flagsAt(wb, si);
  const want = { A1: 'numberStoredAsText', A2: 'numberStoredAsText', A3: '', A4: 'numberStoredAsText', A5: 'numberStoredAsText', A6: 'numberStoredAsText',
    A7: 'numberStoredAsText', A8: 'numberStoredAsText', A10: 'numberStoredAsText', A11: 'numberStoredAsText', A12: '', A13: 'twoDigitTextYear',
    A14: 'twoDigitTextYear', A15: '', A16: '', A17: 'twoDigitTextYear', A18: '', A19: 'numberStoredAsText' };
  assert.deepEqual(Object.fromEntries(Object.keys(want).map((ref) => [ref, at(ref)])), want);
});

test('formulas that evaluate to errors: all but #N/A from NA() itself', () => {
  const { wb, si } = fresh();
  const formulas = { A1: 'NA()', A2: 'VLOOKUP(99,B1:B2,1,FALSE)', A3: '"a"+1', A5: 'NOSUCH()', A7: 'IFERROR(1/0,2)', A8: 'SQRT(-1)', A9: 'A1',
    A11: 'NA()+1', A12: 'MATCH(9,{1,2},0)', A13: '1/0+NA()' };
  for (const [ref, f] of Object.entries(formulas)) wb.setFormula(si, ref, f);
  const at = flagsAt(wb, si);
  const want = { A1: '', A2: 'evalError', A3: 'evalError', A5: 'evalError', A7: '', A8: 'evalError', A9: 'evalError', A11: '', A12: 'evalError', A13: 'evalError' };
  assert.deepEqual(Object.fromEntries(Object.keys(want).map((ref) => [ref, at(ref)])), want);
});

test('inconsistent formulas and formulas that omit adjacent cells, as Excel flags them', () => {
  const { wb, si } = fresh();
  for (let r = 1; r <= 5; r++) { wb.setCell(si, `E${r}`, r); wb.setCell(si, `F${r}`, r * 10); }
  ['E1+F1', 'E2+F2', 'E3*F3', 'E4+F4', 'E5+F5'].forEach((f, k) => wb.setFormula(si, `G${k + 1}`, f));
  wb.setFormula(si, 'H1', 'E1*2'); wb.setFormula(si, 'I1', 'F1*2'); wb.setFormula(si, 'J1', 'E1*3'); wb.setFormula(si, 'K1', 'H1*2');
  // (Omitted: A1:A4 numbers, =SUM(A1:A3) in A5; C2:C4 with text above; I1:I3 summed in I4 and 9 in I5; K1:N1, SUM(K1:M1);
  // a 2-D range; U1, a gap, U3:U4; W1 a date; Y on the sheet, read from another.)
  for (let r = 1; r <= 4; r++) wb.setCell(si, `M${r}`, r);
  wb.setFormula(si, 'M5', 'SUM(M1:M3)');
  wb.setCell(si, 'N1', 'hdr'); for (let r = 2; r <= 4; r++) wb.setCell(si, `N${r}`, r);
  wb.setFormula(si, 'N5', 'SUM(N2:N4)');
  for (let r = 1; r <= 3; r++) wb.setCell(si, `O${r}`, r);
  wb.setFormula(si, 'O4', 'SUM(O1:O3)'); wb.setCell(si, 'O5', 9);
  for (const c of ['Q', 'R', 'S', 'T']) wb.setCell(si, `${c}1`, 1);
  wb.setFormula(si, 'Q3', 'SUM(Q1:S1)');
  for (let r = 1; r <= 4; r++) { wb.setCell(si, `V${r}`, r); wb.setCell(si, `W${r}`, r); }
  wb.setFormula(si, 'V6', 'SUM(V1:W3)');
  wb.setCell(si, 'Y1', 1); wb.setCell(si, 'Y3', 2); wb.setCell(si, 'Y4', 3);
  wb.setFormula(si, 'Y5', 'SUM(Y3:Y4)');
  wb.setFormula(si, 'Z1', 'Z2*2');
  const at = flagsAt(wb, si);
  const want = { G1: '', G2: '', G3: 'formula', G4: '', G5: '', H1: '', I1: '', J1: 'formula', K1: '', M5: 'formulaRange', N5: '',
    O4: '', Q3: 'formulaRange', V6: '', Y5: '' };
  assert.deepEqual(Object.fromEntries(Object.keys(want).map((ref) => [ref, at(ref)])), want);
});

test('an unlocked formula cell; and a sheet\'s ignoredErrors turn a rule off for its cells', () => {
  const { wb, si } = fresh();
  wb.setFormula(si, 'S1', '1+1');
  wb.setStyles(si, [{ r: 0, c: 18, s: wb.styleWith(0, { protection: { locked: false } }) }]);
  wb.setCell(si, 'A1', '123');
  wb.setCell(si, 'A2', '456');
  assert.equal(flagsAt(wb, si)('S1'), 'unlockedFormula');
  const doc = wb._part(si).doc;
  const root = doc.documentElement;
  const ig = doc.createElementNS(root.namespaceURI, 'ignoredErrors');
  const one = doc.createElementNS(root.namespaceURI, 'ignoredError');
  one.setAttribute('sqref', 'A1 S1');
  one.setAttribute('numberStoredAsText', '1');
  one.setAttribute('unlockedFormula', '1');
  ig.appendChild(one);
  root.appendChild(ig);
  const at = flagsAt(wb, si);
  assert.deepEqual([at('A1'), at('A2'), at('S1')], ['', 'numberStoredAsText', '']);
});

test('Ignore Error written as Excel writes it: one ignoredError a rule, its cells joined (A1:A2 A4), after pageMargins', async () => {
  const { ignoreErrors } = await import('../src/ooxml/errorfix.js');
  const { strFromU8, unzipSync } = await import('fflate');
  const { wb, si } = fresh();
  for (let r = 1; r <= 4; r++) wb.setCell(si, `A${r}`, String(r));
  ignoreErrors(wb, si, [{ r: 0, c: 0 }], 'numberStoredAsText');
  ignoreErrors(wb, si, [{ r: 1, c: 0 }], 'numberStoredAsText');
  ignoreErrors(wb, si, [{ r: 3, c: 0 }], 'numberStoredAsText');
  ignoreErrors(wb, si, [{ r: 0, c: 2 }], 'evalError');
  ignoreErrors(wb, si, [{ r: 2, c: 6 }], 'formula');
  ignoreErrors(wb, si, [{ r: 0, c: 7 }], 'twoDigitTextYear');
  const xml = strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]);
  assert.match(xml, /<ignoredErrors><ignoredError sqref="A1:A2 A4" numberStoredAsText="1"\/><ignoredError sqref="C1" evalError="1"\/><ignoredError sqref="G3" formula="1"\/><ignoredError sqref="H1" twoDigitTextYear="1"\/><\/ignoredErrors>/);
  assert.deepEqual([flagsAt(wb, si)('A1'), flagsAt(wb, si)('A3')], ['', 'numberStoredAsText']);
});
