import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { refName } from '../src/ooxml/refs.js';

// Excel's answers for approximate lookups over 400 unsorted lists, some with blanks and text (test/lookup-excel.json,
// from tools/excel-compare/excel-lookups.ps1): on a list that isn't sorted, Excel's binary search gives its own answer.
const { cases } = JSON.parse(fs.readFileSync('test/lookup-excel.json', 'utf8').replace(/^\uFEFF/, ''));

test('MATCH 1, MATCH -1 and VLOOKUP TRUE give Excel\'s answers on unsorted lists', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  const changes = [];
  cases.forEach((c, k) => {
    c.list.forEach((v, r) => { if (v != null) changes.push({ r, c: k, value: v }); });
    const list = `${refName(0, k)}:${refName(c.list.length - 1, k)}`;
    changes.push({ r: 19, c: k, formula: `MATCH(${c.x},${list},1)` }, { r: 20, c: k, formula: `MATCH(${c.x},${list},-1)` },
      { r: 21, c: k, formula: `VLOOKUP(${c.x},${list},1,TRUE)` });
  });
  wb.apply(si, changes);
  const wrong = [];
  cases.forEach((c, k) => {
    for (const [r, key] of [[19, 'match1'], [20, 'matchDown'], [21, 'vlookup']]) {
      const cell = wb.sheet(si).cells.get(r).get(k);
      const ours = cell.v; // (an error's code, as the file has it: #N/A)
      if (ours !== c[key]) wrong.push(`${JSON.stringify(c.list)} x=${c.x} ${key}: Excel ${c[key]}, ours ${ours}`);
    }
  });
  assert.deepEqual(wrong.slice(0, 5), [], `${wrong.length} of ${cases.length * 3} differ`);
});
