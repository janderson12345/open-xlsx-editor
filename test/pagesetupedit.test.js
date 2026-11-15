import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { serializeXml } from '../src/ooxml/xml.js';
import { insertBreak, removeBreak, resetBreaks, setPrintArea, clearPrintArea, printAreaOf } from '../src/ooxml/pagesetupedit.js';
import { build } from './steps.js';

// Excel's own XML for its Page Layout > Breaks commands (ExecuteMso in a hidden Excel; docs/PRINT.md) and the
// _xlnm.Print_Area it saved (PageSetup.PrintArea).
const breaksXml = (wb, si) => (serializeXml(wb._part(si).doc).match(/<(row|col)Breaks[^>]*>.*?<\/(row|col)Breaks>/g) || []);
const ROW = (ids) => `<rowBreaks count="${ids.length}" manualBreakCount="${ids.length}">${ids.map((i) => `<brk id="${i}" max="16383" man="1"/>`).join('')}</rowBreaks>`;
const COL = (ids) => `<colBreaks count="${ids.length}" manualBreakCount="${ids.length}">${ids.map((i) => `<brk id="${i}" max="1048575" man="1"/>`).join('')}</colBreaks>`;

test('Insert Page Break as Excel\'s: at C5 a row and a column break, in row 1 or column A one, at A1 none', () => {
  for (const [at, r, c, want] of [['C5', 4, 2, [ROW([4]), COL([2])]], ['C1', 0, 2, [COL([2])]], ['A5', 4, 0, [ROW([4])]], ['A1', 0, 0, []]]) {
    const { wb, si } = build({ cells: { A1: '1' } });
    assert.equal(insertBreak(wb, si, r, c), at !== 'A1', at);
    assert.deepEqual(breaksXml(wb, si), want, at);
  }
});

test('Remove Page Break and Reset All Page Breaks as Excel\'s', () => {
  const { wb, si } = build({ cells: { A1: '1' } });
  insertBreak(wb, si, 4, 2);
  insertBreak(wb, si, 8, 4);
  // (Excel: C5 then E9 inserted; two of each.)
  assert.deepEqual(breaksXml(wb, si), [ROW([4, 8]), COL([2, 4])]);
  // (Remove at C5: row 5's and column C's; at E3: column E's, row 3 having none; Reset All: none left.)
  removeBreak(wb, si, 4, 2);
  assert.deepEqual(breaksXml(wb, si), [ROW([8]), COL([4])]);
  removeBreak(wb, si, 2, 4);
  assert.deepEqual(breaksXml(wb, si), [ROW([8])]);
  assert.deepEqual(wb.sheet(si).page.breaks, { rows: [8], cols: [] });
  assert.equal(resetBreaks(wb, si), true);
  assert.deepEqual(breaksXml(wb, si), []);
  assert.equal(resetBreaks(wb, si), false);
});

test('the print area as Excel saves it: the sheet\'s own name, absolute, two ranges joined by a comma, cleared gone', () => {
  const { wb, si } = build({ cells: { A1: '1' } });
  wb.renameSheet(si, 'Q1 Plan');
  const area = () => wb.definedNames.filter((d) => d.name === '_xlnm.Print_Area').map((d) => `${d.sheet === si} ${d.text}`);
  setPrintArea(wb, si, [{ r0: 1, c0: 1, r1: 4, c1: 3 }]);
  assert.deepEqual(area(), ["true 'Q1 Plan'!$B$2:$D$5"]);
  setPrintArea(wb, si, [{ r0: 6, c0: 5, r1: 7, c1: 6 }], { add: true });
  assert.deepEqual(area(), ["true 'Q1 Plan'!$B$2:$D$5,'Q1 Plan'!$F$7:$G$8"]);
  assert.equal(printAreaOf(wb, si).length, 2);
  setPrintArea(wb, si, [{ r0: 2, c0: 2, r1: 2, c1: 2 }]);
  assert.deepEqual(area(), ["true 'Q1 Plan'!$C$3"]);
  assert.equal(clearPrintArea(wb, si), true);
  assert.deepEqual(area(), []);
  assert.equal(clearPrintArea(wb, si), false);
});
