import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parseXml, serializeXml } from '../src/ooxml/xml.js';
import { readSparklines } from '../src/ooxml/sparklines.js';
import { sparkPieces, addSparklines, setGroupOptions, setGroupColor, clearSparklines, groupSparklines, ungroupSparklines, setGroupData, setSparklineData, SparklineError, applyStyle } from '../src/ooxml/sparkedit.js';
import { parseRange } from '../src/ooxml/refs.js';
import { Workbook } from '../src/ooxml/workbook.js';

// Sparklines made and edited as Excel does (docs/SPARKLINES.md): test/sparkedit-excel.json is what
// tools/excel-compare/excel-sparkedit.ps1 recorded: Range.SparklineGroups.Add for 16 pairs of location and data, and
// the Sparkline tab's edits (Group, Ungroup, Clear, a theme colour, Edit Data...) on the same start, each saved.
const excel = JSON.parse(fs.readFileSync('test/sparkedit-excel.json', 'utf8'));
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = (t) => parseRange(t);
const data = (t) => (t.includes('!') ? { sheet: t.split('!')[0].replace(/^'|'$/g, ''), range: R(t.split('!')[1]) } : { sheet: 'Sheet1', range: R(t) });

test('the sparklines a location and data make, or Excel\'s refusal', () => {
  for (const a of excel.adds) {
    if (!a.location || !a.data) continue;
    const name = `${a.location} <- ${a.data}`;
    if (a.error) { assert.throws(() => sparkPieces(R(a.location), data(a.data)), SparklineError, name); continue; }
    const got = sparkPieces(R(a.location), data(a.data)).map((p) => `${String.fromCharCode(65 + p.c)}${p.r + 1} <- ${p.source.replace(/^Sheet1!/, '')}`);
    assert.deepEqual(got, a.groups[0].cells, name);
  }
});

/** A sheet with O2:O4 a line over B2:M4 and O6 a red column over B6:M6, as the probe starts each edit. */
function start() {
  const doc = parseXml(`<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`);
  addSparklines(doc, 'line', R('O2:O4'), data('B2:M4'));
  addSparklines(doc, 'column', R('O6'), data('B6:M6'));
  setGroupColor(doc, 5, 14, 'series', { rgb: 'FFFF0000' });
  return doc;
}
const groupsXml = (doc) => serializeXml(doc).match(/<x14:sparklineGroups[\s\S]*<\/x14:sparklineGroups>/)?.[0] ?? '';
/** Excel's text without its made-up ids, and ours, for comparison. */
const bare = (x) => x.replace(/ xr2:uid="\{[0-9A-F-]+\}"/g, '').replace(/ xmlns:x14="[^"]*"/g, '');
const at = (ref) => { const a = R(ref); return [a.r0, a.c0]; };

const EDITS = {
  start: () => {},
  group: (d) => groupSparklines(d, [R('O3'), R('O6')], ...at('O6')),
  groupAtLine: (d) => groupSparklines(d, [R('O3'), R('O6')], ...at('O3')),
  ungroup: (d) => ungroupSparklines(d, [R('O3')]),
  ungroupTwo: (d) => ungroupSparklines(d, [R('O2:O3')]),
  addOver: (d) => addSparklines(d, 'column', R('O3'), data('B3:M3')),
  weight: (d) => setGroupOptions(d, ...at('O2'), { lineWeight: 1.5, minType: 'custom', min: -2, empty: 'zero' }),
  clear: (d) => clearSparklines(d, [R('O3')]),
  clearGroups: (d) => clearSparklines(d, [R('O3')], true),
  clearAll: (d) => clearSparklines(d, [R('O2:O6')]),
  themeColor: (d) => { setGroupColor(d, ...at('O2'), 'series', { theme: 4, tint: -0.249977111117893 }); setGroupColor(d, ...at('O2'), 'markers', { theme: 9 }); },
  modify: (d) => setGroupData(d, ...at('O2'), R('P2:P4'), data('C2:M4')),
  single: (d) => setSparklineData(d, ...at('O2'), data('C3:H3')),
};

test('each of the Sparkline tab\'s edits writes what Excel saved', () => {
  for (const [name, edit] of Object.entries(EDITS)) {
    const doc = start();
    edit(doc);
    assert.equal(bare(groupsXml(doc)), bare(excel.edits[name].ext), name);
  }
});

test('the options read back after writing', () => {
  const doc = start();
  setGroupOptions(doc, 1, 14, { type: 'stacked', markers: true, high: true, negative: true, axis: true, hidden: true, rightToLeft: true, empty: 'span', dates: 'Sheet1!B20:M20', maxType: 'group', minType: 'custom', min: 1 });
  const g = readSparklines(doc.documentElement, MAIN)[1];
  assert.deepEqual([g.type, g.markers, g.high, g.low, g.negative, g.axis, g.hidden, g.rightToLeft, g.empty, g.dates, g.maxType, g.minType, g.min], ['stacked', true, true, false, true, true, true, true, 'span', 'Sheet1!B20:M20', 'group', 'custom', 1]);
  assert.match(groupsXml(doc), /<x14:sparklineGroup manualMin="1" type="stacked" dateAxis="1" displayEmptyCellsAs="span" markers="1" high="1" negative="1" displayXAxis="1" displayHidden="1" minAxisType="custom" maxAxisType="group" rightToLeft="1"><x14:colorSeries/);
  assert.match(groupsXml(doc), /<x14:colorLow rgb="FFD00000"\/><xm:f>Sheet1!B20:M20<\/xm:f><x14:sparklines>/);
  setGroupOptions(doc, 1, 14, { dates: null });
  assert.doesNotMatch(groupsXml(doc), /B20/);
});

test('a sheet with no extLst gets one, last in the worksheet', () => {
  const doc = parseXml(`<worksheet xmlns="${MAIN}"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
  addSparklines(doc, 'line', R('O2'), { sheet: 'My data', range: R('B2:M2') });
  const x = serializeXml(doc);
  assert.match(x, /<pageMargins[^>]*\/><extLst><ext uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}" xmlns:x14="http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2009\/9\/main"><x14:sparklineGroups xmlns:xm=/);
  assert.match(x, /<xm:f>'My data'!B2:M2<\/xm:f><xm:sqref>O2<\/xm:sqref>/);
});

/** A workbook of sparklines written here, every option and both kinds of colour, for Excel to read back
 *  (test/out/sparklines/sparkline-insert.xlsx; tools/excel-compare/excel-sparkedit-check.ps1). */
function written() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/widths.xlsx')));
  const steps = [];
  for (let r = 1; r <= 3; r++) for (let c = 1; c <= 12; c++) steps.push({ r, c, value: ((r * 7 + c * 3) % 11) - 3 });
  for (let c = 1; c <= 12; c++) steps.push({ r: 5, c, value: 45658 + c * c });
  wb.apply(0, steps);
  const name = wb.sheets[0].name;
  const d = (t) => ({ sheet: name, range: R(t) });
  wb.changeSparklines(0, (doc) => {
    addSparklines(doc, 'line', R('O2'), d('B2:M2'));
    addSparklines(doc, 'column', R('O3:O4'), d('B3:M4'));
    addSparklines(doc, 'stacked', R('O5'), d('B2:M2'));
    setGroupOptions(doc, 1, 14, { markers: true, high: true, low: true, first: true, last: true, negative: true, lineWeight: 1.5, empty: 'span' });
    setGroupColor(doc, 1, 14, 'series', { theme: 4, tint: -0.249977111117893 });
    setGroupColor(doc, 1, 14, 'markers', { theme: 9 });
    setGroupColor(doc, 1, 14, 'negative', { rgb: 'FF00B050' });
    setGroupOptions(doc, 2, 14, { maxType: 'group', minType: 'custom', min: -1, axis: true, hidden: true, empty: 'zero' });
    setGroupOptions(doc, 4, 14, { dates: `${name}!B6:M6`, rightToLeft: true });
  });
  return wb;
}

test('sparklines written here, as Excel reads them back (test/sparkedit-readback-excel.json)', () => {
  const wb = written();
  fs.mkdirSync('test/out/sparklines', { recursive: true });
  fs.writeFileSync('test/out/sparklines/sparkline-insert.xlsx', wb.toBytes());
  const back = JSON.parse(fs.readFileSync('test/sparkedit-readback-excel.json', 'utf8'));
  const file = back.find((f) => f.file === 'sparkline-insert.xlsx');
  // (Each group found by its cells: COM's order of them isn't the file's.)
  const by = (cells) => file.groups.find((g) => g.cells.join() === cells);
  const [stacked, column, line] = ['O5 <- B2:M2', 'O3 <- B3:M3,O4 <- B4:M4', 'O2 <- B2:M2'].map(by);
  assert.deepEqual([file.groups.length, stacked.type, column.type, line.type], [3, 3, 2, 1]);
  assert.deepEqual([line.lineWeight, line.blanks, line.markers, line.high, line.low, line.first, line.last, line.negative], [1.5, 3, true, true, true, true, true, true]);
  assert.deepEqual([line.seriesTheme, +line.seriesTint.toFixed(2), line.markersTheme, line.negativeColor], [5, -0.25, 10, 0x50B000]);
  assert.deepEqual([column.minType, column.min, column.maxType, column.axis, column.hidden, column.blanks], [3, -1, 1, true, true, 2]);
  // (A date range on the sparklines' own sheet: COM gives it without the sheet's name.)
  assert.deepEqual([stacked.dateRange, stacked.rtl], ['B6:M6', true]);
  // (The Sparkline tab's own edits, in the browser suite's saved file: Win/Loss, High Point, red, the axis, -3.)
  const tab = back.find((f) => f.file === 'sparkline-edits.xlsx').groups.find((g) => g.cells.join() === 'Q2 <- B2:M2');
  assert.deepEqual([tab.type, tab.high, tab.seriesColor, tab.axis, tab.minType, tab.min], [3, true, 255, true, 3, -3]);
});

test('each of the 36 gallery styles writes the colours Excel wrote (test/sparkstyles-excel.json)', () => {
  const styles = JSON.parse(fs.readFileSync('test/sparkstyles-excel.json', 'utf8'));
  assert.equal(styles.length, 36);
  styles.forEach((st, k) => {
    const doc = start();
    applyStyle(doc, [R('O3')], 2, 14, k);
    const got = Object.fromEntries([...groupsXml(doc).split('<x14:sparklineGroup ')[2].matchAll(/<x14:color(\w+) ([^/]*)\/>/g)].map((m) => [m[1].toLowerCase(), m[2]]));
    assert.deepEqual(got, st.colors, st.name);
  });
  // (Only the groups with a sparkline selected: O6's keeps its red.)
  const doc = start();
  applyStyle(doc, [R('O3')], 2, 14, 0);
  assert.match(groupsXml(doc), /<x14:colorSeries rgb="FFFF0000"\/>/);
});
