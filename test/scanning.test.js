// The code scanning findings fixed (GitHub CodeQL): inputs that made a regular expression slow now run in linear time,
// and a rich text run's font attributes can't break out of their quotes. With the expressions they replaced, the AM/PM
// and trailing-zeros cases below took about 6 seconds each (Node 20); the date-time split and quoted sheet names were
// flagged for inputs like theirs. Each must now finish well inside a second.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseInput } from '../src/ooxml/input.js';
import { parseFormula } from '../src/calc/parse.js';
import { richRunsXml } from '../src/ooxml/strings.js';

const quick = (label, fn) => {
  const t0 = performance.now();
  const result = fn();
  const ms = performance.now() - t0;
  assert.ok(ms < 500, `${label}: ${ms.toFixed(0)} ms`);
  return result;
};
const tryParse = (f) => { try { return parseFormula(f); } catch (e) { return e; } };

test('typed entries: long runs of spaces read in linear time, and read as before', () => {
  const spaces = ' '.repeat(100000);
  assert.deepEqual(quick('date-time split', () => parseInput(`! ${spaces}x`)).value, `! ${spaces}x`.trim());
  assert.equal(typeof quick('AM/PM', () => parseInput(`1${spaces}1pm`)).value, 'string');
  // (What the slices read: AM and PM after a space, the hour or time before it; a date and a time.)
  assert.equal(parseInput('10 am').value, 10 / 24);
  assert.equal(parseInput('10:30   PM').value, (22 * 60 + 30) / 1440);
  assert.equal(typeof parseInput('10:30pm').value, 'string');
  assert.equal(parseInput('1/2/2024 10:30').value, parseInput('1/2/2024').value + (10 * 60 + 30) / 1440);
  assert.equal(parseInput('1/2/2024   10:30 PM').value, parseInput('1/2/2024').value + (22 * 60 + 30) / 1440);
});

test('formulas: a long run of quotes and a long number tokenize in linear time; quoted sheet names as before', () => {
  quick('quotes', () => tryParse(`'${"''&".repeat(30000)}`));
  quick('quotes, closed', () => tryParse(`'${"''&".repeat(30000)}'!A1`));
  quick('zeros', () => tryParse(`1${'0'.repeat(100000)}1`));
  const ref = (f) => JSON.stringify(parseFormula(f));
  assert.match(ref("'Q1 ''24'!A1"), /Q1 '24/);
  assert.match(ref("'a''b'!B2+1"), /a'b/);
  assert.ok(tryParse("'unclosed!A1") instanceof Error);
  assert.ok(tryParse("''!A1") instanceof Error);
});

test('rich text: a run\'s font attributes are escaped, so a value with quotes stays inside its attribute', () => {
  const xml = richRunsXml([{ text: 'x', font: { u: 'double"/><evil/>', vertAlign: '"', color: 'FF0000"/>', name: 'A"B' } }],
    { name: 'Calibri', size: 11 });
  assert.ok(!xml.includes('<evil'), xml);
  assert.ok(!/val="[^"]*"[^ />]/.test(xml), xml);
  assert.match(xml, /<u val="double&quot;\/&gt;&lt;evil\/&gt;"\/>/);
});
