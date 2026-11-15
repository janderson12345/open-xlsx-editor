import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { spawnSync } from 'node:child_process';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRange, parseRef } from '../src/ooxml/refs.js';
import { fillSeries, removeDuplicates, textToColumns } from '../src/ooxml/datatools.js';
import { flashFill } from '../src/ooxml/flashfill.js';
import { copySheet } from '../src/ooxml/sheetcopy.js';
import { setNote } from '../src/ooxml/noteedit.js';
import { insertChart } from '../src/ooxml/chartnew.js';
import { readNotes } from '../src/ooxml/notes.js';
import { readCharts } from '../src/ooxml/charts.js';
import { shownText } from '../src/ooxml/filterops.js';
import { parseInput } from '../src/ooxml/input.js';
import { isDateFormat } from '../src/ooxml/format.js';
import { defineName } from '../src/ooxml/names.js';
import { formatCodeOf } from '../src/ooxml/styles.js';

// Excel's data tools: test/datatools-excel.json holds 27 cases of Fill > Series, Remove Duplicates, Text to Columns and
// Flash Fill (tools/excel-compare/excel-datatools.ps1: each on a new workbook of made-up cells), and a sheet copied
// (Worksheet.Copy). The same cells here, the same tool, and every cell of A1:H12 compared with Excel's: its value
// (numbers to 1e-9), and whether it is a date.
const excel = JSON.parse(fs.readFileSync('test/datatools-excel.json', 'utf8'));
function blank() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  wb.setStyles(0, [...sheet.cells].flatMap(([r, line]) => [...line.keys()].map((c) => ({ r, c, s: 0 }))));
  return wb;
}
/** Cells typed as the probe typed them (dates and numbers read as typed). */
function typed(wb, cells) {
  for (const [ref, text] of Object.entries(cells)) {
    const at = parseRef(ref);
    const p = parseInput(String(text));
    wb.setCell(0, at, p.value ?? String(text));
    if (p.format) wb.setStyles(0, [{ ...at, s: wb.styleWith(0, { numFmt: p.format }) }]);
  }
}
const run = (wb, steps) => { for (const s of steps) { wb.apply(s.si, s.contents); if (s.styles?.length) wb.setStyles(s.si, s.styles); } };
const area = parseRange;
/** Each case's cells and tool, as the probe's script has them. */
const SETUP = {
  'series linear down': [{ A1: '1' }, (wb) => fillSeries(wb, 0, area('A1:A6'), { by: 'columns', step: 2 })],
  'series linear stop': [{ A1: '5' }, (wb) => fillSeries(wb, 0, area('A1:A10'), { by: 'columns', step: 3, stop: 15 })],
  'series growth': [{ A1: '2' }, (wb) => fillSeries(wb, 0, area('A1:A5'), { by: 'columns', type: 'growth', step: 3 })],
  'series across': [{ A1: '10' }, (wb) => fillSeries(wb, 0, area('A1:E1'), { by: 'rows', step: -2.5 })],
  'series months': [{ A1: '1/31/2026' }, (wb) => fillSeries(wb, 0, area('A1:A5'), { by: 'columns', type: 'date', unit: 'month', step: 1 })],
  'series weekdays': [{ A1: '11/13/2026' }, (wb) => fillSeries(wb, 0, area('A1:A6'), { by: 'columns', type: 'date', unit: 'weekday', step: 1 })],
  'series days by 7': [{ A1: '1/1/2026' }, (wb) => fillSeries(wb, 0, area('A1:A4'), { by: 'columns', type: 'date', unit: 'day', step: 7 })],
  'series years': [{ A1: '2/29/2024' }, (wb) => fillSeries(wb, 0, area('A1:A3'), { by: 'columns', type: 'date', unit: 'year', step: 1 })],
  'series trend': [{ A1: '1', A2: '3', A3: '4' }, (wb) => fillSeries(wb, 0, area('A1:A6'), { by: 'columns', trend: true })],
  'series two columns': [{ A1: '1', B1: '10' }, (wb) => fillSeries(wb, 0, area('A1:B4'), { by: 'columns', step: 5 })],
  'remove duplicates all columns': [{ A1: 'Name', B1: 'Team', A2: 'pear', B2: 'red', A3: 'Pear', B3: 'red', A4: 'fig', B4: 'blue', A5: 'pear', B5: 'blue', A6: 'fig', B6: 'blue', D2: 'beside' },
    (wb) => removeDuplicates(wb, 0, area('A1:B6'), [0, 1], true).steps],
  'remove duplicates one column': [{ A1: 'Name', B1: 'Team', A2: 'pear', B2: 'red', A3: 'fig', B3: 'red', A4: 'kiwi', B4: 'blue', A5: 'plum', B5: 'blue', A6: '1', B6: 'red', A7: '1.0', B7: 'red' },
    (wb) => removeDuplicates(wb, 0, area('A1:B7'), [1], true).steps],
  'remove duplicates numbers no header': [{ A1: '1', A2: '2', A3: '1', A4: '3', A5: '2', A6: '1' }, (wb) => removeDuplicates(wb, 0, area('A1:A6'), [0], false).steps],
  'text to columns comma': [{ A1: 'a,b,c', A2: '1,2.5,x', A3: 'one,,three', A4: 'solo' }, (wb) => textToColumns(wb, 0, area('A1:A4'), { delimiters: { comma: true } })],
  'text to columns space consecutive': [{ A1: 'John  Smith', A2: 'Mary Ann Lee', A3: ' lead' }, (wb) => textToColumns(wb, 0, area('A1:A3'), { delimiters: { space: true }, consecutive: true })],
  'text to columns tab and semicolon': [{ A1: 'a\tb;c', A2: '1;2' }, (wb) => textToColumns(wb, 0, area('A1:A2'), { delimiters: { tab: true, semicolon: true } })],
  'text to columns quoted': [{ A1: '"Smith, John",42', A2: 'x,"y"' }, (wb) => textToColumns(wb, 0, area('A1:A2'), { delimiters: { comma: true } })],
  'text to columns dates and zeros': [{ A1: '007,3/4/2026,TRUE,-5%' }, (wb) => textToColumns(wb, 0, area('A1:A1'), { delimiters: { comma: true } })],
  'text to columns elsewhere': [{ A1: 'a-b', A2: 'c-d' }, (wb) => textToColumns(wb, 0, area('A1:A2'), { delimiters: { other: '-' }, dest: { r: 0, c: 2 } })],
  'text to columns fixed width': [{ A1: 'ABC12345 x', A2: 'DE 678', A3: 'XYZ0000999Y', A4: 'ab   cd  e', A5: '1234567' }, (wb) => textToColumns(wb, 0, area('A1:A5'), { breaks: [3, 8] })],
  'text to columns fixed width types': [{ A1: '2026-01-05 12.5  TRUE', A2: '3/4/2026  007   -5%' }, (wb) => textToColumns(wb, 0, area('A1:A2'), { breaks: [10, 16] })],
  'text to columns fixed width elsewhere': [{ A1: 'abc', A2: 'abcdefghij' }, (wb) => textToColumns(wb, 0, area('A1:A2'), { breaks: [2, 5, 8], dest: { r: 0, c: 2 } })],
  'flash fill first names': [{ A1: 'Full', B1: 'First', A2: 'Avery Chen', B2: 'Avery', A3: 'Sam Lee', A4: 'Jo Park', A5: 'Rin Ito' }, (wb) => flashFill(wb, 0, { col: 1, r0: 1, r1: 4, sourceCols: [0] }).steps],
  'flash fill initials': [{ A1: 'Full', B1: 'Short', A2: 'Avery Chen', B2: 'A. Chen', A3: 'Sam Lee', A4: 'Jo Park' }, (wb) => flashFill(wb, 0, { col: 1, r0: 1, r1: 3, sourceCols: [0] }).steps],
  'flash fill upper': [{ A1: 'Code', B1: 'Up', A2: 'ab-12', B2: 'AB-12', A3: 'cd-34', A4: 'ef-56' }, (wb) => flashFill(wb, 0, { col: 1, r0: 1, r1: 3, sourceCols: [0] }).steps],
  'flash fill join': [{ A1: 'First', B1: 'Last', C1: 'Both', A2: 'Avery', B2: 'Chen', C2: 'Chen, Avery', A3: 'Sam', B3: 'Lee', A4: 'Jo', B4: 'Park' }, (wb) => flashFill(wb, 0, { col: 2, r0: 1, r1: 3, sourceCols: [0, 1] }).steps],
  'flash fill digits': [{ A1: 'Ref', B1: 'No', A2: 'INV-0042-X', B2: '0042', A3: 'INV-0107-Y', A4: 'INV-2300-Z' }, (wb) => flashFill(wb, 0, { col: 1, r0: 1, r1: 3, sourceCols: [0] }).steps],
};

test("each data tool's cells as Excel leaves them (27 cases)", () => {
  const wrong = [];
  for (const c of excel.cases) {
    const [cells, tool] = SETUP[c.name];
    const wb = blank();
    typed(wb, cells);
    run(wb, tool(wb));
    const got = {};
    for (const [r, line] of wb.sheet(0).cells) for (const [k, cell] of line) if (r < 12 && k < 8 && cell.v != null && cell.v !== '') got[`${String.fromCharCode(65 + k)}${r + 1}`] = cell;
    const want = Object.keys(c.cells);
    if (JSON.stringify(Object.keys(got).sort()) !== JSON.stringify([...want].sort())) { wrong.push(`${c.name}: Excel has ${want}, here ${Object.keys(got)}`); continue; }
    for (const ref of want) {
      const [text, type, fmt, raw] = c.cells[ref];
      const cell = got[ref];
      const ok = type === 'Double' ? typeof cell.v === 'number' && Math.abs(cell.v - Number(raw)) < 1e-9 : type === 'Boolean' ? cell.v === (raw === 'True') : shownText(wb, cell) === text;
      const dated = type === 'Double' && isDateFormat(fmt);
      const here = typeof cell.v === 'number' && isDateFormat(formatCodeOf(wb.styles, cell.s || 0));
      if (!ok || dated !== here) wrong.push(`${c.name} ${ref}: Excel ${raw} (${fmt}), here ${cell.v} (${here ? 'a date' : 'not a date'})`);
    }
  }
  assert.equal(excel.cases.length, 27);
  assert.deepEqual(wrong, []);
});

test('a sheet copied as Excel copies it: its name, its XML, its names, its notes and charts', () => {
  const wb = blank();
  wb.renameSheet(0, 'Data');
  typed(wb, { A1: 'Item', B1: 'Qty', A2: 'pear', B2: '3', A3: 'fig' });
  wb.apply(0, [{ r: 2, c: 1, formula: 'B2*2' }]);
  setNote(wb, 0, 1, 0, 'A note', 'Avery Chen');
  insertChart(wb, 0, area('A1:B3'), 'columnClustered');
  wb.setConditional(0, [area('B2:B3')], { kind: 'cellIs', operator: 'greaterThan', f1: '4', preset: 'lightRed' });
  const k = copySheet(wb, 0);
  assert.equal(wb.sheets[k].name, 'Data (2)');
  copySheet(wb, 0, 2);
  assert.equal(wb.sheets[2].name, 'Data (3)');
  const files = unzipSync(wb.toBytes());
  // (Excel's copy is the sheet's text but for its new uid and its tab not selected: datatools-excel.json's copy.)
  const bare = (x) => x.replace(/ xr:uid="\{[^}]*\}"/, '').replace(/ tabSelected="1"/, '').replace(/r:id="rId\d+"/g, 'r:id');
  assert.equal(bare(excel.copy.sheet2), bare(excel.copy.sheet1));
  assert.equal(bare(strFromU8(files[wb.sheets[1].path])), bare(strFromU8(files[wb.sheets[0].path])));
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(again.sheets.slice(0, 3).map((s) => s.name), ['Data', 'Data (2)', 'Data (3)']);
  assert.equal(again.sheet(1).cells.get(2).get(1).f, 'B2*2');
  assert.equal(readNotes(again, 1).get('1,0').text, readNotes(again, 0).get('1,0').text);
  const chart = readCharts(again, 1)[0];
  assert.match(chart.series[0].name ? JSON.stringify(chart) : '', /Qty/);
  const copyChart = Object.keys(files).filter((p) => /charts\/chart\d+\.xml$/.test(p)).map((p) => strFromU8(files[p])).find((x) => x.includes("'Data (2)'!"));
  assert.ok(copyChart, "the copy's chart refers to the copy");
  // (The VML boxes of the copies take blocks of their own.)
  const blocks = Object.keys(files).filter((p) => /\.vml$/.test(p)).map((p) => /data="(\d+)"/.exec(strFromU8(files[p]))[1]);
  assert.equal(new Set(blocks).size, blocks.length);
  fs.mkdirSync('test/out/datatools', { recursive: true });
  fs.writeFileSync('test/out/datatools/copy.xlsx', wb.toBytes());
  const check = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/datatools/copy.xlsx'], { encoding: 'utf8' });
  assert.equal(check.status, 0, check.stdout + check.stderr);
});

test("a copy's names: the source's local names, and a workbook-wide name referring to it, as the copy's own (Excel's)", () => {
  const wb = blank();
  wb.renameSheet(0, 'Data');
  typed(wb, { A1: 'x' });
  defineName(wb, { name: 'Local', refersTo: '=Data!$B$2', sheet: 0 });
  defineName(wb, { name: 'Everywhere', refersTo: '=Data!$A$1' });
  copySheet(wb, 0);
  // (basic.xlsx's own names aside: those of Excel's case.)
  const names = wb.definedNames.filter((d) => ['Local', 'Everywhere'].includes(d.name)).map((d) => `${d.name}@${d.sheet ?? '-'}=${d.text}`).sort();
  // (Excel's: excel-datatools.ps1's copy.book.)
  const want = excel.copy.book.match(/<definedName[^>]*>[^<]*<\/definedName>/g).map((x) => { const m = /name="([^"]+)"(?: localSheetId="(\d+)")?>([^<]*)</.exec(x); return `${m[1]}@${m[2] ?? '-'}=${m[3]}`; }).sort();
  assert.deepEqual(names, want);
});
