import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';

// (Excel 365, English, over COM: what MIDB and REPLACEB give when they cut an emoji in two, read back a UTF-16 unit at a
// time. The probe's JSON can't carry half an emoji, so these are here.)
test('MIDB and REPLACEB cut an emoji between its two UTF-16 units, as Excel does', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const units = (f) => [...wb.calc.formulaAt(0, 50, 0, f)].flatMap((ch) => ch.split('').map((u) => u.charCodeAt(0).toString(16)));
  const E = 'UNICHAR(128512)';
  assert.deepEqual(units(`MIDB("a"&${E}&"b",3,1)`), ['de00']);
  assert.deepEqual(units(`MIDB("a"&${E}&"b",2,1)`), ['d83d']);
  assert.deepEqual(units(`MIDB("a"&${E}&"b",2,2)`), ['d83d', 'de00']);
  assert.deepEqual(units(`REPLACEB("a"&${E}&"b",2,1,"X")`), ['61', '58', 'de00', '62']);
});
