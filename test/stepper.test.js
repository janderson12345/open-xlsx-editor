import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { Stepper } from '../src/calc/stepper.js';
import { parseRef } from '../src/ooxml/refs.js';

/** Every step of cell `ref`'s formula, each as its text with the next part [underlined] and the last value _italic_. */
function steps(wb, si, ref) {
  const { r, c } = parseRef(ref);
  wb.calc.build();
  const e = wb.calc.entries.get(`${si}:${r}:${c}`);
  const s = new Stepper(wb.sheet(si).cells.get(r).get(c).f, wb.calc.context(e));
  const out = [];
  const show = () => s.pieces().map((p) => { const t = p.last ? `_${p.text}_` : p.text; return p.under ? `[${t}]` : t; }).join('').replace(/\]\[/g, '');
  out.push(show());
  while (s.step()) out.push(show());
  return out;
}

test("Evaluate Formula's steps, as Microsoft's example has them: a function of a range at once, IF only down its branch", () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  [55, 35, 45, 25].forEach((v, k) => wb.setCell(si, `D${k + 2}`, v));
  [10, 20, 30, 40].forEach((v, k) => wb.setCell(si, `E${k + 2}`, v));
  wb.setFormula(si, 'F1', 'IF(AVERAGE(D2:D5)>50,SUM(E2:E5),0)');
  assert.deepEqual(steps(wb, si, 'F1'), ['=IF([AVERAGE(D2:D5)]>50,SUM(E2:E5),0)', '=IF([_40_>50],SUM(E2:E5),0)', '=[IF(_FALSE_,SUM(E2:E5),0)]', '=_0_']);
  wb.setFormula(si, 'F2', 'IF(D2>50,SUM(E2:E5)*2,"no")');
  assert.deepEqual(steps(wb, si, 'F2'), ['=IF([D2]>50,SUM(E2:E5)*2,"no")', '=IF([_55_>50],SUM(E2:E5)*2,"no")', '=IF(_TRUE_,[SUM(E2:E5)]*2,"no")',
    '=IF(TRUE,[_100_*2],"no")', '=[IF(TRUE,_200_,"no")]', '=_200_']);
  wb.setFormula(si, 'F3', '(D2+D3)/2&" each"');
  assert.deepEqual(steps(wb, si, 'F3'), ['=([D2]+D3)/2&" each"', '=(_55_+[D3])/2&" each"', '=([55+_35_])/2&" each"', '=[_90_/2]&" each"',
    '=[_45_&" each"]', '=_"45 each"_']);
});

test("Step In: the part underlined a cell holding a formula, that cell's formula to work out; Step Out puts its value back", async () => {
  const { build } = await import('./steps.js');
  const { Stepper } = await import('../src/calc/stepper.js');
  const { wb, si } = build({ cells: { A1: '4', B1: '=A1*2', C1: '=B1+1', D1: '=5+1' } });
  wb.calc.build();
  const at = (r, c) => { const e = wb.calc.entries.get(`${si}:${r}:${c}`); return new Stepper(wb.sheet(si).cells.get(r).get(c).f, wb.calc.context(e)); };
  const outer = at(0, 2);
  assert.deepEqual(outer.stepInTarget(), { si, r: 0, c: 1 }, 'B1, holding =A1*2');
  const inner = at(0, 1);
  assert.deepEqual(inner.stepInTarget(), { si, r: 0, c: 0 }, 'A1 is a cell too (a constant: no formula to step into)');
  while (!inner.done) inner.step();
  outer.step();
  assert.equal(outer.pieces().map((p) => p.text).join(''), '=8+1', 'back out with 8 in B1\'s place');
  assert.equal(at(0, 3).stepInTarget(), null, 'nothing to step into in =5+1');
});
