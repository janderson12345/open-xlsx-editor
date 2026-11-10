import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parseInput, formatAfterEntry, editText } from '../src/ooxml/input.js';

// How Excel read 141 typed entries, some into cells formatted first (test/entry-excel.json, from
// tools/excel-compare/excel-entries.ps1; United States settings; `today` gives the year of a date typed without one).
const { today, entries } = JSON.parse(fs.readFileSync('test/entry-excel.json', 'utf8').replace(/^\uFEFF/, ''));
const [y, m, d] = today.split('-').map(Number);
const asOf = new Date(y, m - 1, d);
/** A format code as Excel's object model shows it: without the quotes and backslashes of the file's form. */
const shown = (code) => code.replace(/\\(.)/g, '$1').replace(/"([^"]*)"/g, '$1');
// (Excel gives "1E+400", too big for a number, the scientific format although it stays text: not copied.)
const ODD = new Set(['1E+400']);

test('typed entries read as Excel reads them: value, type and the format the cell gets', () => {
  const wrong = [];
  for (const e of entries) {
    if (ODD.has(e.text)) continue;
    const p = parseInput(e.text, { today: asOf, asText: e.before === '@' });
    const type = p.value == null ? 'blank' : p.value.code ? 'error' : typeof p.value === 'number' ? 'number' : typeof p.value === 'boolean' ? 'boolean' : 'text';
    const before = e.before || 'General';
    const format = formatAfterEntry(before, p) ?? before;
    const sameValue = type === 'number' ? Math.abs(p.value - e.value) <= 1e-9 * Math.max(1, Math.abs(e.value)) : String(p.value?.code ?? p.value) === String(e.value) || (type === 'boolean' && p.value === e.value);
    if (type !== e.type || !sameValue || shown(format) !== e.format) {
      wrong.push(`${JSON.stringify(e.text)}${e.before ? ` in [${e.before}]` : ''}: Excel ${e.type} ${e.value} [${e.format}], ours ${type} ${p.value} [${shown(format)}]`);
    }
  }
  assert.deepEqual(wrong, []);
});

test('the formula bar\'s text types back to the same value', () => {
  const cases = [[45293, 'm/d/yyyy'], [45293.4375, 'm/d/yyyy h:mm'], [0.4375, 'h:mm'], [0.505, '0.00%'], [1.0625, '[h]:mm:ss'], [1234.5, '#,##0.00']];
  for (const [v, code] of cases) {
    const text = editText({ t: 'n', v }, code);
    const back = parseInput(text, { today: asOf }).value;
    assert.ok(Math.abs(back - v) < 1e-9, `${v} [${code}] shown "${text}", read back ${back}`);
  }
  assert.equal(editText({ t: 'n', v: 45293.9375 }, 'm/d/yyyy h:mm'), '1/2/2024 10:30:00 PM');
  assert.equal(editText({ t: 'n', v: 0.5 }, '0%'), '50%');
});

test('a typed error value is that error in the cell: formulas see it', async () => {
  const { build } = await import('./steps.js');
  const { entryChanges } = await import('../src/ooxml/entry.js');
  const { wb, si } = build({ cells: { B1: '=ISNA(A1)', B2: '=ISTEXT(A2)' } });
  wb.apply(si, [entryChanges(wb, si, 0, 0, '#n/a').content, entryChanges(wb, si, 1, 0, '#N/A ').content]);
  wb.calculate();
  const at = (r, c) => wb.sheet(si).cells.get(r).get(c);
  assert.deepEqual([at(0, 0).t, at(0, 0).v, at(0, 1).v], ['e', '#N/A', true]);
  assert.deepEqual([at(1, 0).v, at(1, 1).v], ['#N/A ', true], 'with a space after it, text');
});
