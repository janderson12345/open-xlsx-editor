import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { formatValue, BUILTIN_FORMATS, dateParts } from '../src/ooxml/format.js';
import { PROBE_NUMFMT_BUILTIN, PROBE_NUMFMT_CUSTOM, PROBE_NUMFMT_VALUES } from './fixtures/make-fixtures.mjs';

// What Excel showed for each format of test/fixtures/numfmt.xlsx with each value (its PDF, text after each row's label,
// whitespace runs read as one space), and what formatValue gives.
test('number formats: all 798 format and value pairs of the probe as Excel shows them', () => {
  const excel = JSON.parse(fs.readFileSync('test/numfmt-excel.json', 'utf8')).shown;
  const codes = [...PROBE_NUMFMT_BUILTIN.map((id) => BUILTIN_FORMATS[id]), ...PROBE_NUMFMT_CUSTOM];
  const wrong = [];
  codes.forEach((code, f) => PROBE_NUMFMT_VALUES.forEach((v, k) => {
    const r = formatValue(v, code, { width: 40 });
    const ours = r.hashes ? '#'.repeat(40) : r.text.replace(/\s+/g, ' ').trim();
    if (ours !== excel[f][k]) wrong.push(`${code} ${v}: Excel ${excel[f][k]}, ours ${ours}`);
  }));
  assert.deepEqual(wrong, []);
});

test('format colours, padding and fill come apart from the text', () => {
  assert.equal(formatValue(-5, '[Red]0;[Blue]0').color, '#0000FF');
  assert.deepEqual(formatValue(5, '[Color10]0').color, { indexed: 17 });
  const acct = formatValue(1234.5, BUILTIN_FORMATS[44]);
  assert.deepEqual(acct.parts, [{ pad: '(' }, '$', { fill: ' ' }, '1,234.50', { pad: ')' }]);
});

test('serial dates: 1900 system with its 29 February, 1904 system', () => {
  assert.deepEqual(['y', 'm', 'd'].map((k) => dateParts(60, false)[k]), [1900, 2, 29]);
  assert.deepEqual(['y', 'm', 'd'].map((k) => dateParts(61, false)[k]), [1900, 3, 1]);
  assert.deepEqual(['y', 'm', 'd'].map((k) => dateParts(0, true)[k]), [1904, 1, 1]);
  assert.equal(formatValue(45678.75, 'yyyy-mm-dd', { date1904: true }).text, '2029-01-22');
});
