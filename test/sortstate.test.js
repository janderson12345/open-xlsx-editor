// A sort saved as Excel saves it (src/ooxml/sortstate.js): each of the 17 sorts of test/sortstate-excel.json (Excel over
// COM, tools/excel-compare/excel-sortstate.ps1) made here on the same made-up table, its sortState compared with
// Excel's, its place in the sheet (or in the autoFilter) and its colour's dxf too.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml, kid, kids } from '../src/ooxml/xml.js';
import { parseRange, parseRef } from '../src/ooxml/refs.js';
import { setSortState } from '../src/ooxml/sortstate.js';
import { readTables } from '../src/ooxml/tables.js';

const excel = JSON.parse(fs.readFileSync('test/sortstate-excel.json', 'utf8').trim());
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/sortstate-cases.json', 'utf8'));
const ROWS = [['Fruit', 'Count', 'Colour'], ['kiwi', 30, 'green'], ['Apple', 10, 'red'], ['pear', 20, 'green'], ['Date', 50, 'brown'], ['fig', 40, 'purple']];
const argb = (bgr) => `FF${[bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255].map((x) => x.toString(16).toUpperCase().padStart(2, '0')).join('')}`;

function table() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  ROWS.forEach((row, r) => row.forEach((v, c) => wb.setCell(0, { r, c }, v)));
  return wb;
}
/** A case's sort as the editor's Sort box gives it: each key's column (or, left to right, its row). */
function planOf(c) {
  const area = parseRange(c.range);
  return {
    area, header: c.header !== 'no', matchCase: !!c.matchCase, leftToRight: !!c.leftToRight,
    keys: c.keys.map((k) => {
      const at = parseRef(k.key.split(':')[0]);
      return { at: c.leftToRight ? at.r : at.c, desc: k.order === 'desc', on: k.on, color: k.on ? argb(k.color) : null,
        customOrder: k.customOrder ? k.customOrder.split(',') : undefined };
    }),
  };
}
const own = (s) => s.replace(/ xmlns="[^"]*"/, '');

test('sortState as Excel saves it: 17 of 17 sorts (the key as the column\'s data rows, as Excel\'s Sort box passes it)', () => {
  const wrong = [];
  for (const c of cases) {
    const want = excel.cases.find((x) => x.name === c.name);
    const wb = table();
    if (c.filter) wb.setAutoFilter(0, { r0: 0, c0: 0, r1: 5, c1: 2, columns: new Map() });
    const dxfsBefore = kids(kid(wb.stylesDoc.documentElement, 'dxfs', wb.stylesDoc.documentElement.namespaceURI), 'dxf', wb.stylesDoc.documentElement.namespaceURI).length;
    assert.equal(setSortState(wb, 0, planOf(c)), true, c.name);
    const doc = wb._part(0).doc;
    const ns = doc.documentElement.namespaceURI;
    const auto = kid(doc.documentElement, 'autoFilter', ns);
    const el = c.filter ? kid(auto, 'sortState', ns) : kid(doc.documentElement, 'sortState', ns);
    // (Excel's dxfId counts from its workbook's dxfs, none before; this workbook's count from its own.)
    let got = own(serializeXml(el)).replace(/dxfId="(\d+)"/, (m, n) => `dxfId="${Number(n) - dxfsBefore}"`);
    let expect = want.sortState;
    // (Where the case gave Excel one cell as the key, Excel wrote that cell; the editor writes the column's data, as the
    // cases giving Excel the column's data show. The rest compared as written.)
    if (!c.keys[0].key.includes(':')) { got = got.replace(/ ref="[A-Z]+\d+:[A-Z]+\d+"\/>/g, '/>').replace(/ ref="[A-Z]+\d+:[A-Z]+\d+"( customList| dxfId)/g, '$1'); expect = expect.replace(/ ref="[A-Z]+\d+"\/>/g, '/>').replace(/ ref="[A-Z]+\d+"( customList| dxfId)/g, '$1'); }
    if (got !== expect) wrong.push(`${c.name}\n  Excel ${expect}\n  here  ${got}`);
    if (c.filter && kid(doc.documentElement, 'sortState', ns)) wrong.push(`${c.name}: a sheet-level sortState beside the autoFilter's`);
    if (c.keys[0].on) {
      const dx = kids(kid(wb.stylesDoc.documentElement, 'dxfs', wb.stylesDoc.documentElement.namespaceURI), 'dxf', wb.stylesDoc.documentElement.namespaceURI).at(-1);
      const mine = own(serializeXml(dx));
      const theirs = want.dxfs.match(/<dxf>.*<\/dxf>/s)[0].replace(/<name val="[^"]*"\/><family val="\d+"\/><scheme val="\w+"\/>/, '').replace(/<sz val="\d+"\/>/, '');
      if (mine.replace(/<name val="[^"]*"\/>(<family val="\d+"\/>)?(<scheme val="\w+"\/>)?/, '').replace(/<sz val="[\d.]+"\/>/, '') !== theirs) wrong.push(`${c.name} dxf\n  Excel ${theirs}\n  here  ${mine}`);
    }
  }
  assert.deepEqual(wrong, []);
});

test('sortState: a new sort replaces the last; placed after the autoFilter, before mergeCells; not written in a table', () => {
  const wb = table();
  const plan = planOf(cases[0]);
  setSortState(wb, 0, plan);
  setSortState(wb, 0, { ...plan, keys: [{ at: 0, desc: true }] });
  const xml = serializeXml(wb._part(0).doc);
  assert.equal(xml.match(/<sortState\b/g).length, 1);
  assert.match(xml, /<sortCondition descending="1" ref="A2:A6"\/>/);
  assert.ok(xml.indexOf('<sortState') > xml.indexOf('</sheetData>') && xml.indexOf('<sortState') < xml.indexOf('<pageMargins'));
  // (Saved and opened again: the sortState is in the file as written.)
  const again = Workbook.fromBytes(wb.toBytes());
  assert.match(serializeXml(again._part(0).doc), /<sortState ref="A2:C6" xmlns:xlrd2="[^"]+"><sortCondition descending="1" ref="A2:A6"\/><\/sortState>/);

  // (A sort inside a table: Excel keeps that sortState in the table's part; nothing written here.)
  const t = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/anchors-table.xlsx')));
  const tb = readTables(t)[0];
  assert.equal(setSortState(t, tb.si, { area: { r0: tb.r0, c0: tb.c0, r1: tb.r1, c1: tb.c1 }, header: true, keys: [{ at: tb.c0 }] }), false);
  assert.ok(!serializeXml(t._part(tb.si).doc).includes('<sortState'));
});
