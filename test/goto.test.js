import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { goToTarget, GOTO_MESSAGES } from '../src/ooxml/goto.js';
import { rangeName, MAX_ROWS, MAX_COLS } from '../src/ooxml/refs.js';
import { build } from './steps.js';

// The Name Box: what typing a reference or a name and Enter goes to (Excel's documented behaviour).
test('the Name Box goes to cells, ranges, whole lines, other sheets and defined names', () => {
  const { wb, si, s2 } = build({ cells: { A1: 'x' }, names: { Totals: '=Sheet1!$B$2:$B$9' } });
  const other = wb.sheets[s2].name;
  const go = (text) => { const t = goToTarget(wb, si, text); return t.error ? t : { si: t.si, at: rangeName(t.sel), active: [t.sel.ar, t.sel.ac] }; };
  assert.deepEqual(go('b5'), { si, at: 'B5', active: [4, 1] });
  assert.deepEqual(go(' D10:B5 '), { si, at: 'B5:D10', active: [4, 1] }, 'corners put in order, the active cell the first');
  assert.deepEqual(go('$C$3'), { si, at: 'C3', active: [2, 2] });
  assert.equal(go('B:D').at, `B1:D${MAX_ROWS}`);
  assert.equal(go('7:5').at, `A5:${'XFD'}7`);
  assert.equal(MAX_COLS, 16384);
  assert.deepEqual(go(`${other}!C4`), { si: s2, at: 'C4', active: [3, 2] });
  assert.deepEqual(go(`'${other}'!A1:B2`), { si: s2, at: 'A1:B2', active: [0, 0] });
  assert.deepEqual(go('totals'), { si, at: 'B2:B9', active: [1, 1] }, 'a defined name, in any case');
  assert.deepEqual(goToTarget(wb, si, 'Later'), { error: GOTO_MESSAGES.defining, name: true }, 'a name not there yet: Excel would name the selection');
  for (const bad of ['1abc', 'Nowhere!A1', 'B5:', 'A1:B2:C3', '']) assert.equal(goToTarget(wb, si, bad).error, GOTO_MESSAGES.invalid, bad);
});

test('the Name Box and Go To reach tables: by name their data rows, or a table reference', async () => {
  const fs = await import('node:fs');
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const { goToNames } = await import('../src/ooxml/goto.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const at = (text) => { const t = goToTarget(wb, 0, text); return t.error || `${wb.sheets[t.si].name}!${rangeName(t.sel)}`; };
  assert.deepEqual(['Fruit', 'fruit[Qty]', 'Fruit[#All]', 'Fruit[[#Headers],[Price]]'].map(at), ['Data!A2:E9', 'Data!B2:B9', 'Data!A1:E9', 'Data!C1']);
  assert.equal(at('Fruit[Nope]'), GOTO_MESSAGES.invalid);
  assert.deepEqual(goToNames(wb, 0), ['Fruit', 'Items', 'Rate'], 'Go To lists the names and the tables');
});
