// Tables made and changed as Excel makes and changes them (src/ooxml/tableedit.js): each case of
// tools/excel-compare/table-cases.json replayed here on the same made-up lists, and the table parts (uids and namespace
// declarations aside), the sheet's cells (values, strings and formulas) and the table relationships and content types
// compared with what Excel saved (test/tablemake-excel.json, excel-tables.ps1).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { parseRange, parseRef, rangeName } from '../src/ooxml/refs.js';
import { readTables } from '../src/ooxml/tables.js';
import { createTable, tableEntry, setTableStyle, setTableOption, setTotalsRow, setTotalFunction, setHeaderRow, renameTable, resizeTable, tableAt } from '../src/ooxml/tableedit.js';
import { convertToRange } from '../src/ooxml/tableconvert.js';

const excel = JSON.parse(fs.readFileSync('test/tablemake-excel.json', 'utf8').trim());
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/table-cases.json', 'utf8'));
const DATA = {
  basic: [['Region', 'Item', 'Qty', 'Price'], ['North', 'kiwi', 30, 1.5], ['South', 'apple', 10, 0.5], ['East', 'pear', 20, 2], ['West', 'date', 50, 3], ['North', 'fig', 40, 2.5]],
  numbers: [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
  heads: [['Qty', 2024, '', 'Qty', '=1+1'], ['a', 1, 2, 3, 4], ['b', 5, 6, 7, 8]],
};
const CALC = { 0: 'none', 1: 'sum', 2: 'average', 3: 'count', 4: 'countNums', 5: 'min', 6: 'max', 7: 'stdDev', 8: 'var' };
const OPTION = { ShowTableStyleFirstColumn: 'first', ShowTableStyleLastColumn: 'last', ShowTableStyleRowStripes: 'rows', ShowTableStyleColumnStripes: 'cols', ShowAutoFilterDropDown: 'filterButton' };

function book(data) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  wb.setMerges(0, []);
  // (Excel's new workbook's sheet name, which Convert to Range writes into references.)
  wb.renameSheet(0, 'Sheet1');
  DATA[data].forEach((row, i) => row.forEach((v, j) => {
    if (v === '') return;
    if (typeof v === 'string' && v.startsWith('=')) wb.setFormula(0, { r: i + 1, c: j + 1 }, v.slice(1));
    else wb.setCell(0, { r: i + 1, c: j + 1 }, v);
  }));
  wb.calculate();
  return wb;
}

/** Run a case's steps; the steps Excel takes that this has no command for are named. */
function run(c) {
  const wb = book(c.data);
  let name = null;
  const t = () => readTables(wb).find((x) => x.name === name);
  for (const st of c.steps) {
    if (st.do === 'add') name = createTable(wb, 0, parseRange(st.range), { headers: st.headers !== 'no' && !(st.headers === 'guess' && typeof DATA[c.data][0][0] === 'number') });
    else if (st.do === 'style') setTableStyle(wb, t(), st.value || null);
    else if (st.do === 'set' && st.prop === 'ShowTotals') setTotalsRow(wb, t(), st.value);
    else if (st.do === 'set' && st.prop === 'ShowHeaders') setHeaderRow(wb, t(), st.value);
    else if (st.do === 'set' && OPTION[st.prop]) setTableOption(wb, t(), OPTION[st.prop], st.value);
    else if (st.do === 'name') { const r = renameTable(wb, t(), st.value); if (!r) name = st.value; }
    else if (st.do === 'resize') resizeTable(wb, t(), parseRange(st.range));
    else if (st.do === 'calc') setTotalFunction(wb, t(), st.col - 1, CALC[st.value]);
    else if (st.do === 'unlist') { convertToRange(wb, t()); name = null; }
    else if (st.do === 'set' && st.prop === 'ShowAutoFilter') setTableOption(wb, t(), 'filter', st.value);
    else if (st.do === 'type') {
      // (Typed as the editor enters it: the cell, then what that does to a table.)
      const at = parseRef(st.cell);
      if (st.value.startsWith('=')) wb.setFormula(0, at, st.value.slice(1)); else wb.setCell(0, at, st.value === '' ? null : st.value);
      tableEntry(wb, 0, at.r, at.c);
    }
    else return { skip: `${st.do} ${st.prop || ''}`.trim() };
  }
  return { wb };
}

/** A table part as compared: no declaration, namespace declarations, uids or mc attributes. */
const plain = (xml) => xml.replace(/<\?xml[^>]*>\s*/, '').replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/ mc:Ignorable="[^"]*"/, '').replace(/ xr\d*:uid="[^"]*"/g, '');

/** Excel's sheetData as { ref: value | "=formula" }. */
function excelCells(xml, strings) {
  const out = {};
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>(.*?)<\/c>)/g)) {
    const [, ref, attrs, inner = ''] = m;
    const f = /<f>(.*?)<\/f>/.exec(inner)?.[1];
    const v = /<v>(.*?)<\/v>/.exec(inner)?.[1];
    if (f != null) out[ref] = `=${f.replace(/&amp;/g, '&')}`;
    else if (v != null) out[ref] = / t="s"/.test(attrs) ? strings[Number(v)] : Number(v);
  }
  return out;
}
function ourCells(wb) {
  const out = {};
  for (const [r, line] of wb.sheet(0).cells) {
    for (const [c, cell] of line) {
      const ref = `${String.fromCharCode(65 + c)}${r + 1}`;
      if (cell.f != null) out[ref] = `=${cell.f}`;
      else if (cell.v != null) out[ref] = cell.v;
    }
  }
  return out;
}

// (Each case's file kept for Excel to read back: tools/excel-compare/excel-tables-check.ps1.)
fs.mkdirSync('test/out/tablemake', { recursive: true });

for (const c of cases) {
  const x = excel.find((e) => e.name === c.name);
  test(`table ${c.name}: as Excel writes it`, () => {
    const { wb, skip } = run(c);
    if (skip) return assert.fail(`no command for ${skip}`);
    if (c.name !== 'unlist') wb.setStyles(0, [{ r: 9, c: 1, s: wb.styleWith(0, { cellStyle: 'Good' }) }]);
    fs.writeFileSync(`test/out/tablemake/${c.name}.xlsx`, wb.toBytes());
    const ours = readTables(wb).map((t) => plain(wb.pkg.text(t.path))).sort();
    const theirs = (x.tables || []).map((t) => plain(t.xml)).sort();
    assert.deepEqual(ours, theirs);
    assert.deepEqual(ourCells(wb), excelCells(x.sheetData || '', x.strings || []));
  });
}

/** Each cell's format as compared after Convert to Range: its font's bold, italic and colour, its fill's colour and
 *  its border sides with their colours, from Excel's saved styles (`x`) or ours. */
function excelLooks(x) {
  const list = (xml, tag) => [...(xml || '').matchAll(new RegExp(`<${tag}(?: [^>]*)?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((m) => m[0]);
  const fonts = list(x.fonts.replace(/^<fonts[^>]*>/, ''), 'font');
  const fills = list(x.fills.replace(/^<fills[^>]*>/, ''), 'fill');
  const borders = list(x.borders.replace(/^<borders[^>]*>/, ''), 'border');
  const xfs = list(x.cellXfs.replace(/^<cellXfs[^>]*>/, ''), 'xf');
  const out = {};
  for (const m of x.sheetData.matchAll(/<c r="([A-Z]+\d+)"([^>]*)>/g)) {
    const xf = xfs[Number(/ s="(\d+)"/.exec(m[2])?.[1] || 0)];
    const id = (k) => Number(new RegExp(`${k}="(\\d+)"`).exec(xf)?.[1] || 0);
    out[m[1]] = lookKey(fonts[id('fontId')], fills[id('fillId')], borders[id('borderId')]);
  }
  return out;
}
const lookKey = (font, fill, border) => [
  /<b\/>/.test(font) ? 'b' : '', /<i\/>/.test(font) ? 'i' : '', /<color ([^/]*)\/>/.exec(font)?.[1] || '',
  /<fgColor ([^/]*)\/>/.exec(fill)?.[1] || '',
  ...['left', 'right', 'top', 'bottom'].map((sd) => new RegExp(`<${sd} style="(\\w+)"><color ([^/]*)/>`).exec(border)?.slice(1).join(' ') || ''),
].join('|');

test("Convert to Range keeps the table's look as the cells' own formats, as Excel writes them (unlist)", () => {
  const { wb } = run(cases.find((c) => c.name === 'unlist'));
  const x = excel.find((e) => e.name === 'unlist');
  wb.toBytes();
  const styles = wb.pkg.text('xl/styles.xml');
  const sheet = serializeXml(wb._part(0).doc);
  const ours = excelLooks({ fonts: /<fonts[\s\S]*?<\/fonts>/.exec(styles)[0], fills: /<fills[\s\S]*?<\/fills>/.exec(styles)[0], borders: /<borders[\s\S]*?<\/borders>/.exec(styles)[0],
    cellXfs: /<cellXfs[\s\S]*?<\/cellXfs>/.exec(styles)[0], sheetData: /<sheetData[\s\S]*?<\/sheetData>/.exec(sheet)[0] });
  const theirs = excelLooks(x);
  const diffs = Object.keys(theirs).filter((k) => /^[B-E][2-8]$/.test(k) && theirs[k] !== ours[k]).map((k) => `${k}: ours ${ours[k]} Excel ${theirs[k]}`);
  assert.deepEqual(diffs, []);
});

test('a table made is wired in as Excel wires it: the part, its relationship, content type and the sheet\'s tableParts', () => {
  const { wb } = run(cases.find((c) => c.name === 'add'));
  const x = excel.find((e) => e.name === 'add');
  const types = wb.pkg.text('[Content_Types].xml');
  for (const o of x.contentTypes) assert.ok(types.includes(o.replace(/^<Override /, '').replace(/\/>$/, '')), o);
  assert.match(wb.pkg.text('xl/worksheets/_rels/sheet1.xml.rels'), /Type="[^"]*\/table" Target="..\/tables\/table1.xml"/);
  assert.match(serializeXml(wb._part(0).doc), /<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts>/);
  assert.equal(tableAt(wb, 0, 3, 2)?.name, 'Table1');
});

test("names Excel refuses for a table, in Excel's words", () => {
  const { wb } = run(cases.find((c) => c.name === 'add'));
  for (const bad of ['A1', '1st', 'my table', 'R1C1', '']) assert.ok(renameTable(wb, readTables(wb)[0], bad)?.error, bad);
  assert.equal(renameTable(wb, readTables(wb)[0], 'Ok_name'), null);
  assert.equal(parseRef('B2').r, 1);
});

test("Excel reads back each table the editor wrote as it was written (test/tablemake-readback-excel.json)", () => {
  const back = JSON.parse(fs.readFileSync('test/tablemake-readback-excel.json', 'utf8').trim());
  const CALC_OF = { none: 0, sum: 1, average: 2, count: 3, countNums: 4, min: 5, max: 6, stdDev: 7, var: 8, custom: 9 };
  const diffs = [];
  for (const f of back) {
    if (f.error) { diffs.push(`${f.file}: ${f.error}`); continue; }
    const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/out/tablemake/${f.file}`)));
    const mine = readTables(wb).map((t) => ({ name: t.name, range: rangeName(t), style: t.style.name || '', headers: !!t.headers, totals: !!t.totals, rows: t.style.rows, cols: t.style.cols,
      first: t.style.first, last: t.style.last, filter: !!t.filter, buttons: !!t.filter && !t.filter.hidden.size, columns: t.columns, calc: t.totalsRow.map((x) => CALC_OF[x.fn || 'none']) }));
    const theirs = (f.tables || []).map(({ totalsText, ...t }) => t);
    if (JSON.stringify(mine) !== JSON.stringify(theirs)) diffs.push(`${f.file}: ours ${JSON.stringify(mine)} Excel ${JSON.stringify(theirs)}`);
    if (f.file !== 'unlist.xlsx' && f.b10 !== 'Good') diffs.push(`${f.file}: B10 ${f.b10}`);
  }
  assert.deepEqual(diffs, []);
});
