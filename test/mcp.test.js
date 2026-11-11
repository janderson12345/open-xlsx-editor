// The MCP server (bin/oxe-mcp.mjs) driven as an AI assistant drives it: JSON-RPC over stdin/stdout.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { Workbook } from '../src/ooxml/workbook.js';
import { readPivots } from '../src/ooxml/pivot.js';

function server(root) {
  const child = spawn(process.execPath, [fileURLToPath(new URL('../bin/oxe-mcp.mjs', import.meta.url)), '--root', root], { stdio: ['pipe', 'pipe', 'inherit'] });
  let buf = '';
  const waiting = new Map();
  child.stdout.on('data', (d) => {
    buf += d;
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const msg = JSON.parse(buf.slice(0, i));
      buf = buf.slice(i + 1);
      waiting.get(msg.id)?.(msg);
    }
  });
  let n = 0;
  const call = (method, params) => new Promise((resolve) => { const id = ++n; waiting.set(id, resolve); child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n'); });
  const tool = async (name, args) => { const r = await call('tools/call', { name, arguments: args }); return { text: r.result.content[0].text, error: !!r.result.isError }; };
  return { call, tool, close: () => child.kill() };
}
const folder = () => fs.mkdtempSync(path.join(os.tmpdir(), 'oxe-mcp-'));

test('an assistant reads, writes, reviews its changes, takes one back and saves a copy; nothing outside the folder', async () => {
  const root = folder();
  // (formulas.xlsx with its results worked out and saved, as a workbook Excel saved has them.)
  const fruit = Workbook.fromBytes(new Uint8Array(fs.readFileSync(new URL('./fixtures/formulas.xlsx', import.meta.url))));
  fruit.calculate();
  fs.writeFileSync(path.join(root, 'fruit.xlsx'), fruit.toBytes());
  const s = server(root);
  try {
    const init = await s.call('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } });
    assert.equal(init.result.serverInfo.name, 'open-xlsx-editor');
    assert.match(init.result.instructions, /review_changes/);
    const tools = (await s.call('tools/list', {})).result.tools.map((t) => t.name);
    for (const t of ['list_sheets', 'read_range', 'write_cells', 'set_values', 'insert_rows', 'create_pivot', 'define_name', 'delete_name', 'review_changes', 'revert_changes', 'save_workbook']) assert.ok(tools.includes(t), t);
    assert.equal((await s.tool('list_workbooks', {})).text, 'fruit.xlsx');
    const sheets = JSON.parse((await s.tool('list_sheets', { path: 'fruit.xlsx' })).text);
    assert.deepEqual(sheets.map((x) => x.name), ['Calc', 'Data', 'My Data']);
    assert.match((await s.tool('read_range', { path: 'fruit.xlsx', range: 'Fruit[#All]' })).text, /\| 2 \| Apple \| 3 \|/);
    // An edit, and the review: the cell entered, and the formulas worked out again because of it.
    const w = JSON.parse((await s.tool('write_cells', { path: 'fruit.xlsx', cells: { 'Data!B2': '7' } })).text);
    assert.deepEqual(w, { written: ['Data!B2'], refused: [] });
    const review = (await s.tool('review_changes', { path: 'fruit.xlsx' })).text;
    assert.match(review, /\| Data!B2 \| 3 \| 7 \|/);
    assert.match(review, /Results worked out again: .*Calc!B\d+/);
    const changes = JSON.parse((await s.tool('review_changes', { path: 'fruit.xlsx', format: 'json' })).text);
    assert.equal(changes.cells.filter((c) => c.kind === 'edited').length, 1);
    // Taken back: nothing left to review.
    assert.deepEqual(JSON.parse((await s.tool('revert_changes', { path: 'fruit.xlsx', cells: ['Data!B2'] })).text), { reverted: ['Data!B2'], refused: [] });
    assert.equal((await s.tool('review_changes', { path: 'fruit.xlsx' })).text, 'No changes.');
    // Rows moved on a sheet: its cells can't be taken back one by one, only all at once.
    await s.tool('write_cells', { path: 'fruit.xlsx', cells: { 'My Data!A1': '5' } });
    await s.tool('insert_rows', { path: 'fruit.xlsx', at: '1', sheet: 'My Data' });
    assert.match((await s.tool('review_changes', { path: 'fruit.xlsx' })).text, /1 row inserted before row 1 on My Data/);
    assert.match((await s.tool('revert_changes', { path: 'fruit.xlsx', cells: ["'My Data'!A2"] })).text, /use revertAll/);
    assert.equal((await s.tool('revert_changes', { path: 'fruit.xlsx', all: true })).text, 'Every change taken back.');
    assert.equal((await s.tool('review_changes', { path: 'fruit.xlsx' })).text, 'No changes.');
    // A defined name, used in a formula; the review lists it.
    assert.equal((await s.tool('define_name', { path: 'fruit.xlsx', name: 'Markup', refers_to: '=1.5' })).text, 'Defined Markup.');
    assert.ok((await s.tool('define_name', { path: 'fruit.xlsx', name: 'my markup', refers_to: '=2' })).error);
    await s.tool('write_cells', { path: 'fruit.xlsx', cells: { "'My Data'!B1": '=Markup*2' } });
    assert.equal(JSON.parse((await s.tool('read_cells', { path: 'fruit.xlsx', range: "'My Data'!B1" })).text).rows[0][0].value, 3);
    assert.match((await s.tool('review_changes', { path: 'fruit.xlsx' })).text, /Name Markup defined as =1.5/);
    await s.tool('revert_changes', { path: 'fruit.xlsx', all: true });
    // A PivotTable of the table, on a new sheet; then saved as a copy, which opens with it.
    const made = JSON.parse((await s.tool('create_pivot', { path: 'fruit.xlsx', source: 'Fruit', rows: ['Name'], values: [{ field: 'Qty', fn: 'sum' }] })).text);
    assert.equal(made.name, 'PivotTable1');
    assert.match((await s.tool('read_range', { path: 'fruit.xlsx', range: `'${made.sheet}'!A3:B4` })).text, /Row Labels \| Sum of Qty/);
    assert.match((await s.tool('review_changes', { path: 'fruit.xlsx' })).text, /PivotTable PivotTable1 made at/);
    assert.ok((await s.tool('save_workbook', { path: 'fruit.xlsx' })).error);
    assert.equal((await s.tool('save_workbook', { path: 'fruit.xlsx', save_as: 'fruit (edited).xlsx' })).text, 'Saved fruit (edited).xlsx.');
    assert.ok((await s.tool('save_workbook', { path: 'fruit.xlsx', save_as: 'fruit (edited).xlsx' })).error);
    assert.equal(readPivots(Workbook.fromBytes(new Uint8Array(fs.readFileSync(path.join(root, 'fruit (edited).xlsx'))))).length, 1);
    // Outside the folder, or not a workbook: refused, for reading and for saving.
    assert.ok((await s.tool('read_range', { path: '../escape.xlsx' })).error);
    assert.ok((await s.tool('save_workbook', { path: 'fruit.xlsx', save_as: '../escape.xlsx' })).error);
    assert.ok(!fs.existsSync(path.join(root, '..', 'escape.xlsx')));
    assert.ok((await s.tool('save_workbook', { path: 'fruit.xlsx', save_as: 'fruit.csv' })).error);
    assert.equal((await s.call('tools/call', { name: 'no_such_tool', arguments: {} })).error.code, -32602);
  } finally { s.close(); }
});

test('a workbook saved with a password: refused until opened with it, then read and saved encrypted', async () => {
  const root = folder();
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync(new URL('./fixtures/basic.xlsx', import.meta.url))));
  wb.setPassword('Test-pass 1');
  fs.writeFileSync(path.join(root, 'locked.xlsx'), wb.toBytes());
  const s = server(root);
  try {
    await s.call('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } });
    const refused = await s.tool('list_sheets', { path: 'locked.xlsx' });
    assert.ok(refused.error && /open_workbook/.test(refused.text));
    assert.ok((await s.tool('open_workbook', { path: 'locked.xlsx', password: 'wrong' })).error);
    assert.equal((await s.tool('open_workbook', { path: 'locked.xlsx', password: 'Test-pass 1' })).text, 'Opened locked.xlsx.');
    assert.ok(JSON.parse((await s.tool('list_sheets', { path: 'locked.xlsx' })).text).length >= 3);
    await s.tool('save_workbook', { path: 'locked.xlsx', save_as: 'locked copy.xlsx' });
    const saved = new Uint8Array(fs.readFileSync(path.join(root, 'locked copy.xlsx')));
    assert.throws(() => Workbook.fromBytes(saved));
    assert.ok(Workbook.fromBytes(saved, { password: 'Test-pass 1' }).sheets.length >= 3);
  } finally { s.close(); }
});
