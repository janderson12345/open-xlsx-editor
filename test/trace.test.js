import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { defineName } from '../src/ooxml/names.js';
import { precedents, dependents } from '../src/ooxml/trace.js';
import { traceLevel } from '../src/editor/tracing.js';

// Excel's arrows for the same cells (a hidden Excel's ShowPrecedents, ShowDependents and NavigateArrow): C5 =
// A1+A2+SUM(B1:B3)+Sheet2!A1+Rate+A1, Rate naming D1: arrows from A1, A2, B1:B3 and D1, and one from the other sheet;
// C6 =C5*2, C7 =C5+C6 and Sheet2!B1 =Sheet1!C5: C5's dependents C6, C7 and one arrow to the other sheet.
const book = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.sheetIndexOf('Calc');
  defineName(wb, { name: 'Tax', refersTo: 'Calc!$L$1' });
  wb.setFormula(si, 'K5', "K1+K2+SUM(J1:J3)+'My Data'!A1+Tax+K1");
  wb.setFormula(si, 'K6', 'K5*2');
  wb.setFormula(si, 'K7', 'K5+K6');
  wb.setFormula(wb.sheetIndexOf('My Data'), 'B1', 'Calc!K5');
  return { wb, si };
};

test('precedents as Excel traces them: each range once, a name as what it names, other sheets apart', () => {
  const { wb, si } = book();
  const p = precedents(wb, si, 4, 10);
  const name = (d) => `${d.r0}:${d.c0}:${d.r1}:${d.c1}`;
  assert.deepEqual(p.here.map(name).sort(), ['0:10:0:10', '0:11:0:11', '0:9:2:9', '1:10:1:10'].sort());
  assert.deepEqual(p.elsewhere.map((d) => [wb.sheets[d.si].name, name(d)]), [['My Data', '0:0:0:0']]);
  assert.equal(precedents(wb, si, 0, 10), null);
});

test('dependents as Excel traces them; a level at a time; the arrows of each kind', () => {
  const { wb, si } = book();
  const d = dependents(wb, si, 4, 10);
  assert.deepEqual(d.here.map((x) => `${x.r}:${x.c}`).sort(), ['5:10', '6:10']);
  assert.deepEqual(d.elsewhere.map((x) => wb.sheets[x.si].name), ['My Data']);
  // (K1's dependents: K5; the next level K6 and K7 (twice from K5's, once from K6), then nothing more.)
  const trace = { arrows: [], precedents: null, dependents: null };
  const at = { r: 0, c: 10 };
  assert.equal(traceLevel(wb, si, trace, 'dependents', at), 1);
  assert.equal(traceLevel(wb, si, trace, 'dependents', at), 3);
  assert.equal(traceLevel(wb, si, trace, 'dependents', at), 1);
  assert.equal(traceLevel(wb, si, trace, 'dependents', at), 0);
  assert.ok(trace.arrows.every((a) => a.kind === 'dependents'));
  assert.equal(trace.arrows.filter((a) => a.to === null).length, 1);
});
