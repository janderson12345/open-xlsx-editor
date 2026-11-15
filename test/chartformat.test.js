import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chartParts } from '../src/ooxml/chartedit.js';
import { excel, withChart, EDITS, bare } from './chartformat-cases.js';

test(`each format change written as Excel writes it (${Object.keys(EDITS).length} changes)`, () => {
  const wrong = [];
  for (const [name, edit] of Object.entries(EDITS)) {
    const wb = withChart(name);
    edit(wb);
    const got = bare(wb.pkg.text(chartParts(wb, 0)[0].path), name);
    const want = bare(excel.charts[name], name);
    if (got !== want) {
      let k = 0;
      while (got[k] === want[k]) k++;
      wrong.push(`${name} at ${k}: Excel …${want.slice(k - 80, k + 150)}… here …${got.slice(k - 80, k + 150)}…`);
    }
  }
  // (The unedited chart, as a check of the comparison itself.)
  assert.equal(bare(withChart('none').pkg.text(chartParts(withChart('none'), 0)[0].path), 'none'), bare(excel.charts.none, 'none'));
  assert.deepEqual(wrong, []);
});
