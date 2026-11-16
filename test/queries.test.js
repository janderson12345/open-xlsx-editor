import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, refName } from '../src/ooxml/refs.js';
import { readQueries, readConnections, readMashup, parseSection, formulaProblem, setQueryFormula, refreshQueryTable, tableOfRows, queryStatus, QUERY_MESSAGES } from '../src/ooxml/queries.js';
import { hostRequest } from '../src/editor/queriespane.js';

// Power Query queries and connections as Excel makes them (tools/excel-compare/excel-queries.ps1: Workbook.Queries.Add,
// one loaded to a table through ListObjects.Add over the Mashup OLE DB provider, one a connection only, one to the
// data model; test/queries-excel.xlsx, and what COM reports, test/queries-excel.json). Excel's Refresh of the table
// after the query's result changed: excel-queries-refresh.ps1, test/queries-refresh-excel.json. The editor's M
// written back and read, refreshed and kept by Excel: queries-writeback.mjs then excel-queries-check.ps1,
// test/queries-writeback-excel.json.
const FIXTURE = 'test/queries-excel.xlsx';
const excel = JSON.parse(fs.readFileSync('test/queries-excel.json', 'utf8'));
const refreshes = JSON.parse(fs.readFileSync('test/queries-refresh-excel.json', 'utf8'));
const writeback = JSON.parse(fs.readFileSync('test/queries-writeback-excel.json', 'utf8'));
const wbCases = JSON.parse(fs.readFileSync('tools/excel-compare/queries-writeback-cases.json', 'utf8'));
const open = () => Workbook.fromBytes(fs.readFileSync(FIXTURE));

test('the queries and their M as Workbook.Queries reads them, each with where it loads', () => {
  const { queries } = readQueries(open());
  assert.deepEqual(queries.map((q) => [q.name, q.m]), excel.queries.map((q) => [q.name, q.formula]));
  assert.deepEqual(queries.map((q) => q.load), ['table', 'connection', 'model']);
  const sales = queries[0];
  assert.deepEqual({ name: sales.table.name, sheet: sales.table.sheet, ref: sales.table.ref }, { name: 'Sales', sheet: 'Sheet1', ref: excel.table.range });
  assert.equal(sales.connection.name, 'Query - Sales');
  assert.deepEqual(queries.map(queryStatus), ['4 rows loaded.', 'Connection only.', '4 rows loaded.']);
});

test('the connections as COM lists them: name, the query each reads, in the data model or not', () => {
  const conns = readConnections(open());
  const byName = (a) => [...a].sort((x, y) => x.name.localeCompare(y.name));
  assert.deepEqual(byName(conns).map((c) => [c.name, c.model]), byName(excel.connections).map((c) => [c.name, c.inModel]));
  assert.deepEqual(byName(conns).map((c) => c.query), ['Big', 'Model', 'Sales', null]);
  for (const c of excel.connections.filter((x) => x.command && !x.inModel)) assert.equal(conns.find((x) => x.name === c.name).command, c.command);
});

test('the DataMashup read as [MS-QDEFF] lays it out: its package, permissions, metadata and bindings', () => {
  const mash = readMashup(open());
  assert.equal(mash.path, 'customXml/item1.xml');
  assert.equal(mash.enc, 'utf-16le');
  assert.equal(mash.version, 0);
  assert.equal(mash.rest.length, 0);
  assert.deepEqual(Object.keys(mash.files).sort(), ['Config/Package.xml', 'Formulas/Section1.m', '[Content_Types].xml']);
  assert.equal(mash.meta.Sales.FillTarget, 'Sales');
  assert.equal(mash.meta.Big.FillObjectType, 'ConnectionOnly');
});

test('a workbook with queries saved untouched keeps every part byte for byte', () => {
  const bytes = fs.readFileSync(FIXTURE);
  const before = unzipSync(new Uint8Array(bytes));
  const wb = open();
  readQueries(wb);
  const after = unzipSync(wb.toBytes());
  assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort());
  for (const k of Object.keys(before)) assert.ok(Buffer.from(after[k]).equals(Buffer.from(before[k])), k);
});

test('M written back: only Section1.m changes, read again as written; the rest of the DataMashup byte for byte', () => {
  const wb = open();
  const was = readMashup(wb);
  const m = 'let\n    Source = Sales,\n    #"Filtered Rows" = Table.SelectRows(Source, each [Qty] > 25)\nin\n    #"Filtered Rows"';
  assert.equal(setQueryFormula(wb, 'Big', m), null);
  const now = readMashup(wb);
  assert.equal(now.queries.find((q) => q.name === 'Big').m, m);
  assert.deepEqual(now.queries.map((q) => q.name), ['Sales', 'Big', 'Model']);
  assert.equal(now.queries[0].m, was.queries[0].m);
  for (const k of ['permissions', 'metadata', 'bindings']) assert.ok(Buffer.from(now[k]).equals(Buffer.from(was[k])), k);
  for (const k of ['Config/Package.xml', '[Content_Types].xml']) assert.ok(Buffer.from(now.files[k]).equals(Buffer.from(was.files[k])), k);
  assert.equal(now.section, was.section.replace('[Qty] > 15', '[Qty] > 25'));
  // (The part's text around the base64, and its UTF-16 with its BOM, as they were.)
  assert.equal(now.before, was.before);
  assert.equal(now.after, was.after);
  assert.deepEqual([...wb.pkg.files[was.path].subarray(0, 2)], [0xff, 0xfe]);
  // (Saved and opened again, the same.)
  const back = Workbook.fromBytes(wb.toBytes());
  assert.equal(readQueries(back).queries[1].m, m);
});

test('M the editor wrote is what Excel reads, refreshes and keeps (excel-queries-check.ps1)', () => {
  for (const c of wbCases) {
    const wb = open();
    for (const [name, m] of Object.entries(c.set)) assert.equal(setQueryFormula(wb, name, m), null);
    const x = writeback.find((w) => w.name === c.name);
    assert.deepEqual(x.errors, [], c.name);
    const ours = Object.fromEntries(readQueries(Workbook.fromBytes(wb.toBytes())).queries.map((q) => [q.name, q.m]));
    assert.deepEqual(ours, x.formulas, c.name);
    assert.deepEqual(x.kept, x.formulas, `${c.name}: kept after Excel saved it`);
  }
  // (Excel's refresh of the changed M: the table's rows the new M's.)
  assert.equal(writeback.find((w) => w.name === 'rows').range, 'A1:C7');
  assert.equal(writeback.find((w) => w.name === 'steps').range, 'A1:C4');
});

test('a formula that would break the section is refused, and nothing changes', () => {
  const wb = open();
  const was = Buffer.from(wb.pkg.files['customXml/item1.xml']);
  for (const bad of ['', 'let a = 1; in a', 'let a = "open in a', 'let a = (1 in a', 'let a = 1 /* in a']) {
    const out = setQueryFormula(wb, 'Big', bad);
    assert.ok(out?.error?.startsWith("The query's formula can't be written"), bad);
  }
  assert.ok(Buffer.from(wb.pkg.files['customXml/item1.xml']).equals(was));
  assert.match(setQueryFormula(wb, 'Nope', 'let a = 1 in a').error, /no query named Nope/);
});

test('the section read past strings, quoted names and comments', () => {
  const text = 'section Section1;\r\n\r\n// shared Not = 1;\r\nshared #"My ""Query""" = let\r\n    a = "x;y", /* ; */ b = #"a b"\r\nin\r\n    a;\r\n\r\n[ Description = "d;" ]\r\nshared Two = {1; 2};\r\nshared Three = 3;';
  const qs = parseSection(text);
  assert.deepEqual(qs.map((q) => q.name), ['My "Query"', 'Two', 'Three']);
  assert.equal(qs[0].m, 'let\r\n    a = "x;y", /* ; */ b = #"a b"\r\nin\r\n    a');
  assert.equal(qs[1].m, '{1; 2}');
  assert.equal(formulaProblem('let a = 1 in a'), null);
  assert.equal(formulaProblem('let a = [b = 1; c = 2] in a'), null);
});

/** Excel's sheet as { A1: value }: its shared strings read, numbers as numbers. */
function excelCells(sheetData, strings) {
  const out = {};
  for (const m of sheetData.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>(.*?)<\/c>)/g)) {
    const v = /<v>([^<]*)<\/v>/.exec(m[3] || '')?.[1];
    if (v == null) continue;
    out[m[1]] = /t="s"/.test(m[2]) ? strings[Number(v)] : Number(v);
  }
  return out;
}
function ourCells(wb) {
  const out = {};
  for (const [r, line] of wb.sheet(0).cells) for (const [c, cell] of line) if (cell.v != null && cell.v !== '') out[refName(r, c)] = cell.v;
  return out;
}
const columnsOf = (xml) => [...xml.matchAll(/<tableColumn ([^>]*?)\/?>/g)].map((m) => {
  const a = (n) => new RegExp(`\\b${n}="([^"]*)"`).exec(m[1])?.[1];
  return { id: a('id'), uniqueName: a('uniqueName'), name: a('name'), field: a('queryTableFieldId') };
});
const fieldsOf = (xml) => ({ nextId: /nextId="(\d+)"/.exec(xml)[1], fields: [...xml.matchAll(/<queryTableField ([^>]*?)\/>/g)].map((m) => m[1]) });

test("Refresh writes the rows into the query's table as Excel's does: rows and columns, what's below moving", () => {
  for (const x of refreshes) {
    assert.deepEqual(x.errors, [], x.name);
    const wb = open();
    if (x.name === 'unbound') addUserColumn(wb);
    for (const [ref, v] of Object.entries(x.cells || {})) { const { r, c } = parseRef(ref); wb.apply(0, [{ r, c, value: v }]); }
    const rows = x.csv.split('|').map((line) => line.split(','));
    const q = readQueries(wb).queries[0];
    const out = refreshQueryTable(wb, q, rows);
    assert.equal(out.error, undefined, x.name);
    assert.equal(rangeOf(wb), x.range, x.name);
    assert.deepEqual(ourCells(wb), excelCells(x.sheetData, x.strings), x.name);
    assert.deepEqual(ourFormulas(wb), formulasOf(x.sheetData), x.name);
    assert.deepEqual(columnsOf(wb.pkg.text('xl/tables/table1.xml')), columnsOf(x.table), x.name);
    assert.deepEqual(fieldsOf(wb.pkg.text('xl/queryTables/queryTable1.xml')), fieldsOf(x.queryTable), x.name);
    const name = wb.definedNames.find((d) => d.name === 'ExternalData_1');
    assert.equal(`<definedNames><definedName name="ExternalData_1" localSheetId="0" hidden="1">${name.text}</definedName></definedNames>`, x.definedNames, x.name);
    // (Saved, the table part still a well-formed queryTable table.)
    const back = Workbook.fromBytes(wb.toBytes());
    assert.equal(readQueries(back).queries[0].table.ref, x.range);
  }
});
const rangeOf = (wb) => readQueries(wb).queries[0].table.ref;
/** The case 'unbound''s column of the user's, as ListColumns.Add left it (its parts as Excel's refresh kept them:
 *  tableColumn 1 with its calculated formula, queryTableField 4 not bound, one unbound column on the right). */
function addUserColumn(wb) {
  const t = wb.pkg.text('xl/tables/table1.xml')
    .replace(/ref="A1:C5"/g, 'ref="A1:D5"').replace('<tableColumns count="3">', '<tableColumns count="4">')
    .replace('</tableColumns>', '<tableColumn id="1" uniqueName="1" name="Double" queryTableFieldId="4"><calculatedColumnFormula>Sales[[#This Row],[Qty]]*2</calculatedColumnFormula></tableColumn></tableColumns>');
  wb.pkg.setText('xl/tables/table1.xml', t);
  const qt = wb.pkg.text('xl/queryTables/queryTable1.xml')
    .replace('<queryTableRefresh nextId="4">', '<queryTableRefresh nextId="5" unboundColumnsRight="1">').replace('<queryTableFields count="3">', '<queryTableFields count="4">')
    .replace('</queryTableFields>', '<queryTableField id="4" dataBound="0" tableColumnId="1"/></queryTableFields>');
  wb.pkg.setText('xl/queryTables/queryTable1.xml', qt);
  wb.apply(0, [{ r: 0, c: 3, value: 'Double' }, ...[1, 2, 3, 4].map((r) => ({ r, c: 3, formula: 'Sales[[#This Row],[Qty]]*2' }))]);
}
/** A sheet's formulas as { E2: text }. */
const formulasOf = (sheetData) => Object.fromEntries([...sheetData.matchAll(/<c r="([A-Z]+\d+)"[^>]*><f>([^<]*)<\/f>/g)].map((m) => [m[1], m[2]]));
const ourFormulas = (wb) => { const out = {}; for (const [r, line] of wb.sheet(0).cells) for (const [c, cell] of line) if (cell.f) out[refName(r, c)] = cell.f; return out; };

test('a host answer as objects, as { columns, rows }; a Date as its serial; none made one empty row', () => {
  assert.deepEqual(tableOfRows([{ a: 1, b: 'x' }, { b: 'y', c: true }]), { columns: ['a', 'b', 'c'], rows: [[1, 'x', undefined], [undefined, 'y', true]] });
  assert.deepEqual(tableOfRows({ columns: ['a'], rows: [[1], [2]] }), { columns: ['a'], rows: [[1], [2]] });
  assert.equal(tableOfRows('nope'), null);
  const wb = open();
  const q = readQueries(wb).queries[0];
  assert.equal(refreshQueryTable(wb, q, 'nope').error, QUERY_MESSAGES.rows);
  refreshQueryTable(wb, q, { columns: ['When', 'Qty'], rows: [[new Date(Date.UTC(2024, 0, 1)), 3]] });
  assert.equal(rangeOf(wb), 'A1:B2');
  assert.equal(wb.sheet(0).cells.get(1).get(0).v, 45292);
  refreshQueryTable(wb, readQueries(wb).queries[0], { columns: ['When', 'Qty'], rows: [] });
  assert.equal(rangeOf(wb), 'A1:B2');
  assert.equal(wb.sheet(0).cells.get(1)?.get(0)?.v ?? null, null);
  assert.match(refreshQueryTable(wb, readQueries(wb).queries[1], [['a'], [1]]).error, /isn't loaded to a table/);
});

test('what the host is asked: the query, its connection, its M and its table', () => {
  const q = readQueries(open()).queries[0];
  assert.deepEqual(hostRequest(q, 'query'), {
    name: 'Sales', kind: 'query', m: excel.queries[0].formula,
    connection: { name: 'Query - Sales', description: "Connection to the 'Sales' query in the workbook.", type: 'OLE DB', connection: 'Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Sales;Extended Properties=""', command: 'SELECT * FROM [Sales]' },
    target: { sheet: 'Sheet1', ref: 'A1:C5', table: 'Sales' },
  });
  assert.equal(QUERY_MESSAGES.cant(['Query - Sales']), 'Unable to refresh data for a data connection in the workbook. The following connections failed to refresh:\n\nQuery - Sales');
});

test('a workbook without queries or connections has none', () => {
  const wb = Workbook.fromBytes(fs.readFileSync('test/fixtures/tables.xlsx'));
  assert.deepEqual(readQueries(wb), { queries: [], connections: [], mashup: false });
});
