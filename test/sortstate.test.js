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
import { changeLines } from '../src/ooxml/rowcol.js';
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
  for (const c of cases.filter((x) => !x.after && !x.table)) {
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

test('sortState: a new sort replaces the last; placed after the autoFilter, before mergeCells', () => {
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

  // (A sort inside a table: in the table's part, the next test.)
});

test('a saved sort follows inserted and deleted rows and columns as Excel moves it: 10 of 11 (Excel keeps the first row when the header row above is deleted)', () => {
  const wrong = [];
  for (const c of cases.filter((x) => x.after && !x.table)) {
    const want = excel.cases.find((x) => x.name === c.name);
    const wb = table();
    if (c.filter) wb.setAutoFilter(0, { r0: 0, c0: 0, r1: 5, c1: 2, columns: new Map() });
    setSortState(wb, 0, planOf(c));
    for (const st of c.after) {
      const [a0, a1] = st.range.split(':');
      const rows = /^\d+$/.test(a0);
      const at = rows ? Number(a0) - 1 : parseRef(`${a0}1`).c;
      const last = rows ? Number(a1) - 1 : parseRef(`${a1}1`).c;
      changeLines(wb, 0, rows ? 'row' : 'col', at, (st.do === 'insert' ? 1 : -1) * (last - at + 1));
    }
    const xml = serializeXml(wb._part(0).doc);
    const got = (xml.match(/<sortState\b.*?<\/sortState>/) || [''])[0].replace(/ xmlns(:\w+)?="[^"]*"/g, '');
    const expect = (want.sortState || '').replace(/ xmlns(:\w+)?="[^"]*"/g, '');
    // (Excel's own quirk: the header row above deleted, its sortState kept starting on row 2, the levels moving up.)
    if (c.name === 'then-header-row-deleted') { assert.equal(expect, '<sortState ref="A2:C5"><sortCondition descending="1" ref="C1:C5"/><sortCondition ref="A1:A5"/></sortState>'); continue; }
    if (got !== expect) wrong.push(`${c.name}\n  Excel ${expect}\n  here  ${got}`);
  }
  assert.deepEqual(wrong, []);
});

test('a sort inside a table saved in the table\'s part, after its autoFilter, as Excel saves it, and following a row inserted', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/anchors-table.xlsx')));
  const tb = readTables(wb)[0];
  assert.equal(setSortState(wb, tb.si, { area: { r0: tb.r0, c0: tb.c0, r1: tb.r1, c1: tb.c1 }, header: true, matchCase: true, keys: [{ at: 2, desc: true }, { at: 1 }] }), true);
  const part = () => wb.pkg.text(tb.path).replace(/ xmlns(:\w+)?="[^"]*"/g, '');
  // (Excel's table-two-keys-case, on this table's place: caseSensitive, the data rows, a level per key.)
  assert.match(part(), /<autoFilter ref="B2:D12"\/><sortState caseSensitive="1" ref="B3:D12"><sortCondition descending="1" ref="C3:C12"\/><sortCondition ref="B3:B12"\/><\/sortState><tableColumns/);
  assert.ok(!serializeXml(wb._part(tb.si).doc).includes('<sortState'));
  // (Excel's table-then-row-inserted: the table, its autoFilter, sortState and levels a row longer.)
  changeLines(wb, tb.si, 'row', 5, 1);
  assert.match(part(), /<sortState caseSensitive="1" ref="B3:D13"><sortCondition descending="1" ref="C3:C13"\/><sortCondition ref="B3:B13"\/><\/sortState>/);
});
