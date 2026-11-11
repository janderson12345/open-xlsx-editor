import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { spawnSync } from 'node:child_process';
import { Workbook } from '../src/ooxml/workbook.js';
import { BAR_COLORS } from '../src/ooxml/cfbarsedit.js';
import { conditionalLooks } from '../src/ooxml/condformat.js';
import { parseRange } from '../src/ooxml/refs.js';

// Data bars and icon sets as Home > Conditional Formatting adds them: test/cfbars-excel.json holds what Excel wrote
// for 13 data bars and 20 icon sets, one a column (tools/excel-compare/excel-cf-bars.ps1). Each is added here to the same
// column and its rule, and its rule in the extLst, compared with Excel's text, Excel's made-up ids aside.
const excel = JSON.parse(fs.readFileSync('test/cfbars-excel.json', 'utf8'));
const COLORS = { Blue: 'blue', Green: 'green', Red: 'red', Orange: 'orange', LightBlue: 'lightBlue', Purple: 'purple' };
const ruleOf = (name) => {
  if (name === 'barNegative') return { kind: 'dataBar', color: BAR_COLORS.blue };
  const bar = /^bar(\w+?)(Gradient|Solid)$/.exec(name);
  if (bar) return { kind: 'dataBar', color: BAR_COLORS[COLORS[bar[1]]], solid: bar[2] === 'Solid' };
  return { kind: 'iconSet', set: name.replace(/^icons/, '') };
};
const open = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
const bareIds = (x) => x.replace(/\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}/g, '{id}');
/** Each conditionalFormatting by its ranges (sqref, or an extLst one's xm:sqref). */
const bySqref = (list, re) => new Map(list.map((x) => [re.exec(x)[1], bareIds(x)]));

test('each data bar and icon set written as Excel writes it: the rule and its extLst part (33 rules)', () => {
  const wb = open();
  // (Added last column first: each new rule comes first, so column A ends with priority 1, as COM numbered them.)
  const cases = Object.entries(excel.columns).sort((a, b) => b[1].length - a[1].length || b[1].localeCompare(a[1]));
  for (const [name, col] of cases) wb.setConditional(0, [parseRange(`${col}1:${col}10`)], ruleOf(name));
  const sheet = strFromU8(unzipSync(wb.toBytes())['xl/worksheets/sheet1.xml']);
  const ours = bySqref(sheet.match(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g) || [], /sqref="([^"]+)"/);
  const theirs = bySqref(excel.conditionalFormatting, /sqref="([^"]+)"/);
  const extRe = /<x14:conditionalFormatting\b[\s\S]*?<\/x14:conditionalFormatting>/g;
  const oursX = bySqref(sheet.match(extRe) || [], /<xm:sqref>([^<]+)</);
  const theirsX = bySqref(excel.extLst.match(extRe) || [], /<xm:sqref>([^<]+)</);
  const wrong = [];
  for (const [sq, want] of theirs) if (ours.get(sq) !== want) wrong.push(`${sq}: Excel ${want}\n   here ${ours.get(sq)}`);
  for (const [sq, want] of theirsX) if (oursX.get(sq) !== want) wrong.push(`${sq} (extLst): Excel ${want}\n   here ${oursX.get(sq)}`);
  assert.equal(theirs.size + theirsX.size, 13 + 20 + 13);
  assert.deepEqual(wrong, []);
  // (Excel's own ext comes first in the sheet's extLst.)
  assert.match(sheet, /<extLst><ext uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}"/);
  fs.mkdirSync('test/out/cf', { recursive: true });
  fs.writeFileSync('test/out/cf/cfbars.xlsx', wb.toBytes());
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/cf/cfbars.xlsx'], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
  // Read back: each column's rule, with its extLst settings.
  const again = Workbook.fromBytes(wb.toBytes());
  const rules = again.sheet(0).condFormats.flatMap((cf) => cf.rules);
  assert.equal(rules.filter((r) => r.type === 'dataBar' && r.dataBar.minLength === 0 && r.dataBar.min.type === 'autoMin').length, 13);
  assert.equal(rules.filter((r) => r.type === 'iconSet').length, 20);
  // Clear Rules from column A and from the stars' column: their extLst parts go too.
  again.setConditional(0, [parseRange('A1:A10'), parseRange(`${excel.columns.icons3Stars}1:${excel.columns.icons3Stars}10`)], null);
  const left = strFromU8(unzipSync(again.toBytes())['xl/worksheets/sheet1.xml']);
  assert.doesNotMatch(left, /sqref="A1:A10"|<xm:sqref>A1:A10<|3Stars/);
});

test('what they draw, by Excel\'s documented rules: bars from 0 (or the axis), icons by thirds of the span', () => {
  const wb = open();
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  for (let r = 0; r < 10; r++) { wb.setCell(0, { r, c: 0 }, r + 1); wb.setCell(0, { r, c: 1 }, r - 3); wb.setCell(0, { r, c: 2 }, r + 1); }
  wb.setConditional(0, [parseRange('A1:A10')], { kind: 'dataBar', color: BAR_COLORS.blue });
  wb.setConditional(0, [parseRange('B1:B10')], { kind: 'dataBar', color: BAR_COLORS.green, solid: true });
  wb.setConditional(0, [parseRange('C1:C10')], { kind: 'iconSet', set: '3Arrows' });
  const again = Workbook.fromBytes(wb.toBytes());
  const looks = conditionalLooks(again, 0);
  const bar = (r, c) => { const b = looks.get(`${r},${c}`).bar; return [Number(b.from.toFixed(3)), Number(b.to.toFixed(3))]; };
  // (1 to 10, automatic min 0 and max 10: each bar its share of the width.)
  assert.deepEqual([bar(0, 0), bar(4, 0), bar(9, 0)], [[0, 0.1], [0, 0.5], [0, 1]]);
  // (-3 to 6 with an automatic axis: the axis a third in, -3 the whole left third, 6 the whole right.)
  assert.deepEqual([bar(0, 1), bar(3, 1), bar(9, 1)], [[0, 0.333], [0.333, 0.333], [0.333, 1]]);
  assert.equal(looks.get('0,1').bar.color, '#FF0000');
  assert.equal(looks.get('9,1').bar.color, '#63C384');
  // (Thresholds at 0%, 33% and 67% of 1 to 10: 3.97 and 7.03.)
  assert.deepEqual([...Array(10).keys()].map((r) => looks.get(`${r},2`).icon.index).join(''), '0001111222');
});

test("Excel's reading of them: each column's rule, its bar's colour, fill, border and automatic ends, or its icon set (test/cfbars-readback-excel.json)", () => {
  // (excel-cf-bars-check.ps1 opened test/out/cf/cfbars.xlsx, made by the first test, in Excel.)
  const back = JSON.parse(fs.readFileSync('test/cfbars-readback-excel.json', 'utf8')).rules;
  const SETS = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack',
    '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes'];
  const bgr = (argb) => parseInt(argb.slice(6, 8) + argb.slice(4, 6) + argb.slice(2, 4), 16);
  const wrong = [];
  for (const name of Object.keys(excel.columns)) {
    const want = ruleOf(name);
    const got = back[name];
    const ok = want.kind === 'dataBar'
      ? got?.type === 'dataBar' && got.color === bgr(want.color) && got.solid === !!want.solid && got.border === (want.solid ? 0 : 1) && got.min === 6 && got.max === 7
      : got?.type === 'iconSet' && got.set === SETS.indexOf(want.set) + 1;
    if (!ok) wrong.push(`${name}: ${JSON.stringify(got)}`);
  }
  assert.deepEqual(wrong, []);
});
