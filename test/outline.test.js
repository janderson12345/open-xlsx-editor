import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml as serialize } from '../src/ooxml/xml.js';
import { typedWidth } from '../src/ooxml/sizes.js';
import { groupLines, ungroupLines, showLevels, hideDetail, showDetail, clearOutline, setSummary, outlineOf, autoOutline, setOutlineSymbols } from '../src/ooxml/outline.js';

// Excel's outlines: test/outline-excel.json holds what Excel wrote for each case (tools/excel-compare/excel-outline.ps1:
// Range.Group, Ungroup, Outline.ShowLevels, Range.ShowDetail, ClearOutline, Outline.SummaryRow and SummaryColumn, each on
// a new workbook of made-up numbers, saved and read back). The same steps here on a new sheet, and every row's and
// column's outlineLevel, collapsed, hidden, height and width compared with Excel's, with sheetFormatPr's
// outlineLevelRow and outlineLevelCol and sheetPr's outlinePr.
const excel = JSON.parse(fs.readFileSync('test/outline-excel.json', 'utf8'));
const byName = new Map(excel.cases.map((c) => [c.name, c]));

/** A new sheet with Excel's probe numbers: A1:A12 and B1:H1. */
function numbers() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Outline');
  wb.apply(si, [...Array.from({ length: 12 }, (_, r) => ({ r, c: 0, value: (r + 1) * 10 })), ...Array.from({ length: 7 }, (_, k) => ({ r: 0, c: k + 1, value: k + 2 }))]);
  return { wb, si };
}
const R = (a, b = a) => ['row', a - 1, b - 1];
const C = (a, b = a) => ['col', a.charCodeAt(0) - 65, b.charCodeAt(0) - 65];
const group = (wb, si, [axis, a, b]) => groupLines(wb, si, axis, a, b);
/** Each case's steps, as the probe's script has them (1-based rows and lettered columns, as Excel's). */
const STEPS = {
  'group rows 2-4': (wb, si) => group(wb, si, R(2, 4)),
  'group rows twice nested': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); },
  'group rows three levels': (wb, si) => { group(wb, si, R(2, 9)); group(wb, si, R(3, 7)); group(wb, si, R(4, 5)); },
  'group rows two groups side by side': (wb, si) => { group(wb, si, R(2, 3)); group(wb, si, R(5, 6)); },
  'group rows touching': (wb, si) => { group(wb, si, R(2, 3)); group(wb, si, R(4, 5)); },
  'group empty rows past the data': (wb, si) => group(wb, si, R(20, 22)),
  'group rows then collapse level 1': (wb, si) => { group(wb, si, R(2, 4)); showLevels(wb, si, 'row', 1); },
  'nested collapse level 1': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); showLevels(wb, si, 'row', 1); },
  'nested collapse level 2': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); showLevels(wb, si, 'row', 2); },
  'nested level 1 then level 3': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); showLevels(wb, si, 'row', 1); showLevels(wb, si, 'row', 3); },
  'hide detail of inner group': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); hideDetail(wb, si, 'row', 4); },
  'hide inner then outer then show outer': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); hideDetail(wb, si, 'row', 4); hideDetail(wb, si, 'row', 7); showDetail(wb, si, 'row', 7); },
  'hide outer then show inner': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); hideDetail(wb, si, 'row', 7); assert.equal(showDetail(wb, si, 'row', 4), false); },
  'hide detail then show detail': (wb, si) => { group(wb, si, R(2, 4)); hideDetail(wb, si, 'row', 4); showDetail(wb, si, 'row', 4); },
  'hide detail from a detail row': (wb, si) => { group(wb, si, R(2, 4)); hideDetail(wb, si, 'row', 2); },
  'ungroup inner': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(3, 4)); ungroupLines(wb, si, ...R(3, 4)); },
  'ungroup part of a group': (wb, si) => { group(wb, si, R(2, 6)); ungroupLines(wb, si, ...R(3, 4)); },
  'ungroup a collapsed group': (wb, si) => { group(wb, si, R(2, 4)); hideDetail(wb, si, 'row', 4); ungroupLines(wb, si, ...R(2, 4)); },
  'ungroup rows not grouped': (wb, si) => { group(wb, si, R(2, 4)); assert.equal(ungroupLines(wb, si, ...R(8, 9)), false); },
  'summary above': (wb, si) => { setSummary(wb, si, { below: false }); group(wb, si, R(3, 5)); },
  'summary above collapse': (wb, si) => { setSummary(wb, si, { below: false }); group(wb, si, R(3, 5)); showLevels(wb, si, 'row', 1); },
  'summary left columns': (wb, si) => { setSummary(wb, si, { right: false }); group(wb, si, C('C', 'D')); showLevels(wb, si, 'col', 1); },
  'summary below and right set again': (wb, si) => { setSummary(wb, si, { below: false }); setSummary(wb, si, { below: true }); group(wb, si, R(3, 5)); },
  'group columns B-D': (wb, si) => group(wb, si, C('B', 'D')),
  'group columns collapse': (wb, si) => { group(wb, si, C('B', 'D')); showLevels(wb, si, 'col', 1); },
  'group columns nested hide inner': (wb, si) => { group(wb, si, C('B', 'F')); group(wb, si, C('C', 'D')); hideDetail(wb, si, 'col', 4); },
  'group columns with widths': (wb, si) => { wb.setLines(si, 'col', 2, 2, { width: typedWidth(20) }); group(wb, si, C('B', 'D')); showLevels(wb, si, 'col', 1); },
  'group rows with a height': (wb, si) => { wb.setLines(si, 'row', 2, 2, { height: 30 }); group(wb, si, R(2, 4)); showLevels(wb, si, 'row', 1); },
  'group rows already hidden': (wb, si) => { wb.setLines(si, 'row', 2, 2, { hidden: true }); group(wb, si, R(2, 4)); },
  'eight levels': (wb, si) => { for (let k = 0; k < 8; k++) assert.equal(group(wb, si, R(2 + k, 20 - k)), k < 7); },
  'ninth level refused': (wb, si) => { for (let k = 0; k < 8; k++) group(wb, si, R(2 + k, 20 - k)); assert.equal(group(wb, si, R(10, 11)), false); },
  'clear outline': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, C('B', 'C')); showLevels(wb, si, 'row', 1); showLevels(wb, si, 'col', 1); clearOutline(wb, si, 'row'); clearOutline(wb, si, 'col'); },
  'auto outline': (wb, si) => {
    // (The probe's cells: 3 rows of numbers in A and B, SUMs under them in row 4 and beside them in C.)
    wb.apply(si, Array.from({ length: 12 }, (_, r) => ({ r, c: 0, value: null })).concat(Array.from({ length: 7 }, (_, n) => ({ r: 0, c: n + 1, value: null }))));
    const cells = [];
    for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) cells.push({ r, c, value: r + c + 3 });
    cells.push({ r: 3, c: 0, formula: 'SUM(A1:A3)' }, { r: 3, c: 1, formula: 'SUM(B1:B3)' });
    for (let r = 0; r < 4; r++) cells.push({ r, c: 2, formula: `SUM(A${r + 1}:B${r + 1})` });
    wb.apply(si, cells);
    assert.equal(autoOutline(wb, si), true);
  },
  'outline symbols off': (wb, si) => { group(wb, si, R(2, 4)); setOutlineSymbols(wb, si, false); },
  'outline symbols off without an outline': (wb, si) => setOutlineSymbols(wb, si, false),
  'outline symbols off then on': (wb, si) => { group(wb, si, R(2, 4)); setOutlineSymbols(wb, si, false); setOutlineSymbols(wb, si, true); },
  'clear outline of part': (wb, si) => { group(wb, si, R(2, 7)); group(wb, si, R(9, 10)); clearOutline(wb, si, ...R(2, 7)); },
};

const attrsOf = (tag) => Object.fromEntries([...tag.matchAll(/ ([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
/** A sheet's rows and columns as the outline sees them: { rows: { r: {...} }, cols: { c: {...} }, format, pr }. */
function lines(xml) {
  const pick = (a, keys) => Object.fromEntries(keys.filter((k) => a[k] != null).map((k) => [k, a[k]]));
  const rows = {};
  for (const m of xml.matchAll(/<row\b([^>]*?)\/?>/g)) {
    const a = attrsOf(m[1]);
    const p = pick(a, ['ht', 'hidden', 'customHeight', 'outlineLevel', 'collapsed']);
    if (Object.keys(p).length) rows[a.r] = p;
  }
  const cols = {};
  for (const m of xml.matchAll(/<col\b([^>]*?)\/>/g)) {
    const a = attrsOf(m[1]);
    const p = pick(a, ['width', 'hidden', 'customWidth', 'outlineLevel', 'collapsed']);
    for (let c = Number(a.min); c <= Number(a.max); c++) cols[c] = p;
  }
  const fmt = attrsOf((xml.match(/<sheetFormatPr\b[^>]*>/) || [''])[0]);
  const pr = (xml.match(/<outlinePr\b[^>]*\/>/) || [''])[0];
  const view = attrsOf((xml.match(/<sheetView\b[^>]*>/) || [''])[0]);
  return { rows, cols, format: pick(fmt, ['outlineLevelRow', 'outlineLevelCol']), outlinePr: attrsOf(pr), symbols: view.showOutlineSymbols ?? null };
}

for (const [name, steps] of Object.entries(STEPS)) {
  test(`outline as Excel writes it: ${name}`, () => {
    const want = byName.get(name);
    assert.ok(want, `no measurement for ${name}`);
    const { wb, si } = numbers();
    steps(wb, si);
    const got = lines(serialize(wb._part(si).doc));
    const exp = lines(want.sheetPr + want.sheetView + want.sheetFormatPr + want.cols + want.sheetData);
    assert.deepEqual(got, exp);
  });
}

test('every grouping case of the probe is checked', () => {
  const left = excel.cases.map((c) => c.name).filter((n) => !/subtotal|styles/.test(n) && !STEPS[n]);
  assert.deepEqual(left, []);
});

test('the outline the grid draws: groups, their summaries and the deepest level', () => {
  const { wb, si } = numbers();
  group(wb, si, R(2, 7));
  group(wb, si, R(3, 4));
  group(wb, si, C('B', 'C'));
  const o = outlineOf(wb.sheet(si));
  assert.equal(o.rows.level, 2);
  assert.deepEqual(o.rows.groups.map(({ level, start, end, summary }) => [level, start, end, summary]), [[1, 1, 6, 7], [2, 2, 3, 4]]);
  assert.deepEqual(o.cols.groups.map(({ level, start, end, summary }) => [level, start, end, summary]), [[1, 1, 2, 3]]);
  setSummary(wb, si, { below: false });
  assert.deepEqual(outlineOf(wb.sheet(si)).rows.groups.map((g) => g.summary), [0, 1]);
  assert.equal(wb.sheet(si).rows.get(2).level, 2);
});
