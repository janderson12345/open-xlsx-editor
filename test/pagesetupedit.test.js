import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { serializeXml } from '../src/ooxml/xml.js';
import { insertBreak, removeBreak, resetBreaks, setPrintArea, clearPrintArea, printAreaOf, readTitles, setPrintTitles, printTitlesOf } from '../src/ooxml/pagesetupedit.js';
import { readArea } from '../src/editor/pagesetupbox.js';
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

test('print titles as Excel takes and saves them: columns first, then rows; loose rows read; a lone row refused', () => {
  // (Excel's PrintTitleRows: "1:2" and "$1:$2" are $1:$2, "B2:C4" is $2:$4, "$3" refused.)
  assert.deepEqual(readTitles('1:2', 'rows'), [0, 1]);
  assert.deepEqual(readTitles('$1:$2', 'rows'), [0, 1]);
  assert.deepEqual(readTitles('B2:C4', 'rows'), [1, 3]);
  assert.equal(readTitles('$3', 'rows'), undefined);
  assert.equal(readTitles('', 'rows'), null);
  assert.deepEqual(readTitles('$A:$B', 'cols'), [0, 1]);
  assert.deepEqual(readTitles("'Q1 Plan'!$A:$A", 'cols'), [0, 0]);
  const { wb, si } = build({ cells: { A1: '1' } });
  wb.renameSheet(si, 'Q1 Plan');
  const titles = () => wb.definedNames.filter((d) => d.name === '_xlnm.Print_Titles').map((d) => d.text);
  setPrintTitles(wb, si, { rows: [0, 1], cols: [0, 1] });
  assert.deepEqual(titles(), ["'Q1 Plan'!$A:$B,'Q1 Plan'!$1:$2"]);
  assert.deepEqual(printTitlesOf(wb, si), { rows: [0, 1], cols: [0, 1] });
  setPrintTitles(wb, si, { rows: null, cols: [0, 1] });
  assert.deepEqual(titles(), ["'Q1 Plan'!$A:$B"]);
  setPrintTitles(wb, si, { rows: [0, 0] });
  assert.deepEqual(titles(), ["'Q1 Plan'!$1:$1"]);
  setPrintTitles(wb, si, {});
  assert.deepEqual(titles(), []);
  // (The box's print area: ranges by commas, the sheet's name allowed; anything else refused.)
  assert.deepEqual(readArea("$B$2:$D$5,'Q1 Plan'!$F$7"), [{ r0: 1, c0: 1, r1: 4, c1: 3 }, { r0: 6, c0: 5, r1: 6, c1: 5 }]);
  assert.deepEqual(readArea(''), []);
  assert.equal(readArea('B2:'), null);
});
