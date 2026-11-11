#!/usr/bin/env node
// An MCP server (Model Context Protocol, over stdin/stdout) that lets an AI assistant read and edit .xlsx and .xlsm
// workbooks in one folder, through the agent interface (src/agent.js, docs/AGENT.md): sheets, names and tables, any
// range as Markdown or values, search, notes; cells entered as a person types them, blocks of values, rows, columns,
// sheets, PivotTables; and a review of everything it changed, cell by cell, with any of it put back, before it saves.
// Excel has no tracked changes for cells, so the review is the workbook beside the one opened (agentreview.js). It reads
// and writes only inside the folder it's given (--root, else the current folder), never overwrites a file unless asked,
// and makes no network requests.
//   usage: open-xlsx-editor-mcp [--root <folder>]
import fs from 'node:fs';
import path from 'node:path';
import readline from 'node:readline';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../src/ooxml/xml.js';
import { openWorkbook } from '../src/agent.js';

setXmlImpl({ DOMParser, XMLSerializer });
const arg = (name, fallback) => { const i = process.argv.indexOf(name); return i > 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback; };
const ROOT = fs.realpathSync(path.resolve(arg('--root', process.cwd())));
const VERSION = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;
const open = new Map(); // full path -> AgentWorkbook

/** A path inside the folder (relative to it, or absolute within it); anything outside is refused. */
function inside(p, { mustExist = true } = {}) {
  if (typeof p !== 'string' || !p) throw new Error('A file path is needed.');
  const full = path.resolve(ROOT, p);
  const real = mustExist ? fs.realpathSync(full) : path.join(fs.realpathSync(path.dirname(full)), path.basename(full));
  if (real !== ROOT && !real.startsWith(ROOT + path.sep)) throw new Error(`${p} is outside the folder this server works in.`);
  if (!/\.xls[xm]$/i.test(real)) throw new Error(`${p} isn't an .xlsx or .xlsm file.`);
  return real;
}
function book(p, password = null) {
  const full = inside(p);
  if (!open.has(full)) {
    try {
      open.set(full, openWorkbook(new Uint8Array(fs.readFileSync(full)), { password }));
    } catch (e) {
      if (/password/i.test(e.message) && !password) throw new Error(`${p} is protected with a password: open it with open_workbook and its password.`);
      throw e;
    }
  }
  return open.get(full);
}
const json = (v) => JSON.stringify(v, null, 1);
const S = (description, extra = {}) => ({ type: 'string', description, ...extra });
const B = (description) => ({ type: 'boolean', description });
const N = (description) => ({ type: 'number', description });
const PATH = S('The workbook, relative to the folder the server works in.');
const RANGE = S('A cell or range as Excel\'s Name Box takes it: "B2", "B2:D9", "A:A", "Data!B2:C5", a defined name, a table ("Sales"), or a table reference ("Sales[Qty]", "Sales[#All]").');
const SHEET = S('The sheet a range without a sheet name is on (default: the active sheet).');
const obj = (properties, required = ['path']) => ({ type: 'object', properties: { path: PATH, ...properties }, required });

const TOOLS = [
  { name: 'list_workbooks', description: 'The .xlsx and .xlsm files in the folder the server works in (and its subfolders).', inputSchema: { type: 'object', properties: {} },
    run: () => {
      const out = [];
      const walk = (d) => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { if (e.name.startsWith('.') || e.name === 'node_modules') continue; const f = path.join(d, e.name); if (e.isDirectory()) { if (out.length < 500) walk(f); } else if (/\.xls[xm]$/i.test(e.name) && !e.name.startsWith('~$')) out.push(path.relative(ROOT, f)); } };
      walk(ROOT);
      return out.join('\n') || 'No workbooks here.';
    } },
  { name: 'open_workbook', description: 'Opens a workbook saved with a password (others open by themselves at their first use).', inputSchema: obj({ password: S('The workbook\'s password.') }, ['path', 'password']),
    run: (a) => { open.delete(inside(a.path)); book(a.path, a.password); return `Opened ${a.path}.`; } },
  { name: 'list_sheets', description: 'The sheets in tab order: name, kind, hidden, used range, tables.', inputSchema: obj({}),
    run: (a) => json(book(a.path).sheets()) },
  { name: 'list_names', description: 'The defined names and tables, with what each refers to and each table\'s columns.', inputSchema: obj({}),
    run: (a) => json(book(a.path).names()) },
  { name: 'read_range', description: 'A range as a Markdown table: column letters across, row numbers down, each cell as it shows (or its formula, with formulas). With no range, the used range of the sheet.', inputSchema: obj({ range: RANGE, sheet: SHEET, formulas: B('Show formulas instead of their results.') }),
    run: (a) => book(a.path).markdown(a.range ?? null, { sheet: a.sheet ?? null, formulas: !!a.formulas }) || 'Empty.' },
  { name: 'read_cells', description: 'A range\'s cells as JSON: each cell\'s ref, value (number, text, true/false, error text or null), text as shown, and formula.', inputSchema: obj({ range: RANGE, sheet: SHEET }, ['path', 'range']),
    run: (a) => json(book(a.path).read(a.range, { sheet: a.sheet ?? null })) },
  { name: 'find', description: 'Cells whose formula (or shown text, with look_in "values") matches, with Excel\'s wildcards * and ?.', inputSchema: obj({ query: S('The text to find.'), look_in: S('"formulas" (default) or "values".'), within: S('"workbook" (default) or "sheet".'), sheet: SHEET, match_case: B('Match capitals.'), whole: B('Match entire cell contents.') }, ['path', 'query']),
    run: (a) => json(book(a.path).find(a.query, { lookIn: a.look_in || 'formulas', within: a.within || 'workbook', sheet: a.sheet ?? null, matchCase: !!a.match_case, whole: !!a.whole })) },
  { name: 'list_notes', description: 'A sheet\'s notes and threaded comments, with replies.', inputSchema: obj({ sheet: SHEET }),
    run: (a) => json(book(a.path).notes({ sheet: a.sheet ?? null })) },
  { name: 'write_cells', description: 'Enters text into cells as a person types it: {"B2": "12", "C2": "=B2*2", "D2": "50%", "Data!A9": "Kiwi"}. Numbers, dates, percentages and TRUE are read as Excel reads them; "=" starts a formula; "\'" starts text. A cell\'s data validation refuses what it wouldn\'t take, unless force.', inputSchema: obj({ cells: { type: 'object', description: 'Cell -> text, as typed.' }, sheet: SHEET, force: B('Enter values data validation would refuse (only when the user asks).') }, ['path', 'cells']),
    run: (a) => json(book(a.path).write(a.cells || {}, { sheet: a.sheet ?? null, force: !!a.force })) },
  { name: 'set_values', description: 'A block of values from a top-left cell, row by row, taken as they are (text is never read as a formula or date; null empties a cell).', inputSchema: obj({ top_left: S('The block\'s first cell, e.g. "A1" or "Data!B2".'), rows: { type: 'array', items: { type: 'array' }, description: 'Rows of values: numbers, text, true/false or null.' }, sheet: SHEET }, ['path', 'top_left', 'rows']),
    run: (a) => `Wrote ${book(a.path).setValues(a.top_left, a.rows, { sheet: a.sheet ?? null })}.` },
  ...[['insert_rows', 'insertRows', 'Inserts rows before a row (its number), as Excel\'s Insert does: formulas, names, tables, charts and the rest follow.', 'row'], ['delete_rows', 'deleteRows', 'Deletes rows from a row (its number), as Excel\'s Delete does.', 'row'],
    ['insert_columns', 'insertColumns', 'Inserts columns before a column (its letter or number), as Excel\'s Insert does.', 'column'], ['delete_columns', 'deleteColumns', 'Deletes columns from a column (its letter or number), as Excel\'s Delete does.', 'column']]
    .map(([name, method, description, what]) => ({ name, description, inputSchema: obj({ at: S(what === 'row' ? 'The row number, e.g. "5".' : 'The column, e.g. "C" or "3".'), count: N('How many (default 1).'), sheet: SHEET }, ['path', 'at']),
      run: (a) => { book(a.path)[method](a.at, a.count ?? 1, { sheet: a.sheet ?? null }); return 'Done.'; } })),
  { name: 'add_sheet', description: 'Adds a worksheet (at the end, or before another). Returns its name.', inputSchema: obj({ name: S('Its name (default: Excel\'s next SheetN).'), before: S('The sheet to put it before.') }),
    run: (a) => `Added ${book(a.path).addSheet(a.name ?? null, { before: a.before ?? null })}.` },
  { name: 'rename_sheet', description: 'Renames a sheet; formulas, names and charts follow.', inputSchema: obj({ sheet: S('The sheet.'), name: S('Its new name.') }, ['path', 'sheet', 'name']),
    run: (a) => { book(a.path).renameSheet(a.sheet, a.name); return 'Renamed.'; } },
  { name: 'delete_sheet', description: 'Deletes a sheet.', inputSchema: obj({ sheet: S('The sheet.') }, ['path', 'sheet']),
    run: (a) => { book(a.path).deleteSheet(a.sheet); return 'Deleted.'; } },
  { name: 'create_pivot', description: 'Makes a PivotTable as Excel\'s Insert > PivotTable does, from a range with headings or a table: fields to rows and columns, values summarized by sum, count, average, max or min. At a cell, or on a new sheet at A3.', inputSchema: obj({
    source: S('The data: a range with its headings in the first row ("Data!A1:E19"), or a table\'s name.'), rows: { type: 'array', items: { type: 'string' }, description: 'Field names for rows, in order.' },
    columns: { type: 'array', items: { type: 'string' }, description: 'Field names for columns (at most one).' }, values: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' }, fn: { type: 'string' } } }, description: 'Values: [{"field": "Units", "fn": "sum"}].' },
    at: S('Where it goes, e.g. "Report!A3" (default: a new sheet).') }, ['path', 'source', 'values']),
    run: (a) => json(book(a.path).pivot(a.source, { rows: a.rows || [], columns: a.columns || [], values: a.values || [] }, { at: a.at ?? null })) },
  { name: 'refresh_pivots', description: 'Refreshes every PivotTable from its source, as Data > Refresh All does.', inputSchema: obj({}),
    run: (a) => { const names = book(a.path).refreshPivots(); return names.length ? `Refreshed ${names.join(', ')}.` : 'No PivotTables.'; } },
  { name: 'define_name', description: 'Defines a name, as Excel\'s Define Name does: refers_to as typed ("=Data!$B$2:$B$9", "=0.07"), workbook-wide or on one sheet.', inputSchema: obj({ name: S('The name, e.g. "TaxRate".'), refers_to: S('What it refers to, e.g. "=Data!$B$2".'), sheet: S('A sheet to make it local to (default: the workbook).'), comment: S('A comment.') }, ['path', 'name', 'refers_to']),
    run: (a) => { book(a.path).defineName(a.name, a.refers_to, { sheet: a.sheet ?? null, comment: a.comment || '' }); return `Defined ${a.name}.`; } },
  { name: 'edit_name', description: 'Changes a defined name: a new name (formulas using it follow), what it refers to, its comment.', inputSchema: obj({ name: S('The name.'), new_name: S('Its new name.'), refers_to: S('What it now refers to.'), comment: S('Its comment.'), sheet: S('The sheet it is local to, if it is.') }, ['path', 'name']),
    run: (a) => { book(a.path).editName(a.name, { name: a.new_name, refersTo: a.refers_to, comment: a.comment }, { sheet: a.sheet ?? null }); return 'Changed.'; } },
  { name: 'delete_name', description: 'Deletes a defined name (formulas using it then show #NAME?).', inputSchema: obj({ name: S('The name.'), sheet: S('The sheet it is local to, if it is.') }, ['path', 'name']),
    run: (a) => { book(a.path).deleteName(a.name, { sheet: a.sheet ?? null }); return 'Deleted.'; } },
  { name: 'calculate', description: 'Works out every formula again (Ctrl+Alt+F9).', inputSchema: obj({}),
    run: (a) => { book(a.path).calculate(); return 'Calculated.'; } },
  { name: 'review_changes', description: 'Everything changed since the workbook was opened: sheets added, deleted or renamed, rows and columns, PivotTables, and each cell before and after (and the results worked out again because of them). Show this to the user before saving.', inputSchema: obj({ format: S('"markdown" (default) or "json".') }),
    run: (a) => { const b = book(a.path); return a.format === 'json' ? json(b.changes()) : b.changes({ markdown: true }); } },
  { name: 'revert_changes', description: 'Puts cells back as they were when the workbook was opened (formats too), or every change with all.', inputSchema: obj({ cells: { type: 'array', items: { type: 'string' }, description: 'Cells or ranges to put back, e.g. ["Data!B2", "C3:C9"].' }, all: B('Take back every change since the workbook was opened.') }),
    run: (a) => { const b = book(a.path); if (a.all) { b.revertAll(); return 'Every change taken back.'; } return json(b.revert(a.cells || [])); } },
  { name: 'save_workbook', description: 'Saves the edits: to save_as (a new file in the folder), or over the original only with overwrite. Everything not changed is kept as it was.', inputSchema: obj({ save_as: S('A new file name for the edited copy, e.g. "budget (edited).xlsx".'), overwrite: B('Save over the original (only when the user asked for that).') }),
    run: (a) => {
      const b = book(a.path);
      let target;
      if (a.save_as) {
        target = inside(a.save_as, { mustExist: false });
        if (fs.existsSync(target) && !a.overwrite) throw new Error(`${a.save_as} already exists; give another name, or overwrite.`);
        if (path.extname(target).toLowerCase() !== path.extname(inside(a.path)).toLowerCase()) throw new Error(`Save as ${path.extname(inside(a.path))}, the kind of file it is.`);
      } else if (a.overwrite) target = inside(a.path);
      else throw new Error('Give save_as (a new file name), or overwrite: true to save over the original.');
      fs.writeFileSync(target, b.save());
      return `Saved ${path.relative(ROOT, target)}.`;
    } },
];

// --- JSON-RPC over stdin/stdout, one message per line (MCP's stdio transport) ---------------------------------------
const send = (msg) => process.stdout.write(JSON.stringify({ jsonrpc: '2.0', ...msg }) + '\n');
async function handle(msg) {
  const { id, method, params } = msg;
  if (id == null) return; // (notifications: initialized, cancelled... nothing to answer)
  try {
    if (method === 'initialize') {
      return send({ id, result: { protocolVersion: params?.protocolVersion || '2025-06-18', capabilities: { tools: {} }, serverInfo: { name: 'open-xlsx-editor', version: VERSION },
        instructions: 'Edits Excel workbooks in one folder. Read with list_sheets and read_range; write with write_cells (as a person types). Excel has no tracked changes for cells: before saving, show the user review_changes, and take back what they reject with revert_changes. Save to a new file with save_workbook save_as unless the user asks to overwrite.' } });
    }
    if (method === 'ping') return send({ id, result: {} });
    if (method === 'tools/list') return send({ id, result: { tools: TOOLS.map(({ run, ...t }) => t) } });
    if (method === 'tools/call') {
      const tool = TOOLS.find((t) => t.name === params?.name);
      if (!tool) return send({ id, error: { code: -32602, message: `Unknown tool: ${params?.name}` } });
      try {
        const text = await tool.run(params.arguments || {});
        return send({ id, result: { content: [{ type: 'text', text: String(text) }] } });
      } catch (e) {
        return send({ id, result: { content: [{ type: 'text', text: e.message }], isError: true } });
      }
    }
    send({ id, error: { code: -32601, message: `Method not found: ${method}` } });
  } catch (e) {
    send({ id, error: { code: -32603, message: e.message } });
  }
}
readline.createInterface({ input: process.stdin }).on('line', (line) => {
  if (!line.trim()) return;
  let msg;
  try { msg = JSON.parse(line); } catch { return send({ id: null, error: { code: -32700, message: 'Parse error' } }); }
  handle(msg);
});
