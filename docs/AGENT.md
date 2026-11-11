# The agent interface

`src/agent.js` lets an AI agent, or any program, read and edit a workbook in plain terms. It works without a browser
(`openWorkbook(bytes)`) or on the workbook open in an editor (`api.agent()`). Types are in `src/agent.d.ts`. It is the
sister of the Word editor's `agent.js`. Excel has no tracked changes for cells, so edits are made directly. In an
editor, each edit is one step of the undo history and is drawn as it is made.

```js
import { openWorkbook } from 'open-xlsx-editor';
const book = openWorkbook(bytes);        // in Node, set up a DOMParser and XMLSerializer first, as test/setup.js does
book.sheets();                            // [{ name: 'Data', kind: 'worksheet', hidden: false, used: 'A1:J9', tables: ['Fruit'] }, ...]
book.markdown('Fruit[#All]');             // the table as Markdown: column letters across, row numbers down
book.write({ 'Data!B2': '7', 'Data!F2': '=B2*2' });   // { written: ['Data!B2', 'Data!F2'], refused: [] }
const bytesOut = book.save();
```

## Ranges

Ranges are whatever the Name Box takes (`goto.js`):

- a cell or range: `B2`, `B2:D9`
- whole columns or rows: `A:A`, `3:5`, cut to the used range when read
- any of those on another sheet: `Data!B2`, `'My Data'!A1`
- a defined name: `Rate`
- a table, `Fruit` for its data rows, or a table reference: `Fruit[Qty]`, `Fruit[#All]`

A range without a sheet is on the `sheet` option (a name or an index), or else on the active sheet. Ranges that come
back are sheet-qualified, with the name quoted when it needs it.

## Reading

| Call | Gives |
|---|---|
| `sheets()` | Each sheet: name, kind, hidden, used range, its tables. |
| `names()` | The defined names and tables Go To lists, with what each refers to (and a table's columns). |
| `read(range)` | `{ range, rows }`: each cell's `ref`, `value` (number, string, boolean, error text or null), `text` (as shown, its number format applied, at any width) and `formula` (with "=", or null). |
| `markdown(range, { formulas })` | The same as a Markdown table, or with formulas shown instead of results. |
| `find(query, { lookIn, within, matchCase, whole })` | Matching cells, with Excel's wildcards, as Find does (`findops.js`). |
| `notes()` | A sheet's notes and threaded comments with their replies. |

## Editing

| Call | Does |
|---|---|
| `write({ ref: text })` | Enters each text as a person types it (`entry.js`): numbers, dates, percentages, currency and TRUE read as Excel reads them; `=` starts a formula; `'` starts text. A cell's data validation refuses what it wouldn't take, unless `force`. Refusals come back in `refused` with Excel's words; the other cells are still written. |
| `setValues(topLeft, rows)` | A block of values taken as they are: a string is never read as a formula or a date. |
| `insertRows` / `deleteRows` / `insertColumns` / `deleteColumns` | As Excel's Insert and Delete (`rowcol.js`); refusals throw with the editor's words. |
| `addSheet` / `renameSheet` / `deleteSheet` | As the tab menu does them; refusals throw with Excel's words. |
| `pivot(source, { rows, columns, values }, { at })` | A PivotTable as Excel's Insert > PivotTable makes it (`pivotnew.js`, docs/PIVOT.md): field names to rows and columns, values summarized by sum, count, average, max or min; at a cell, or on a new sheet at A3. Returns its name, sheet and range. |
| `refreshPivots()` | Every PivotTable refreshed from its source, as Data > Refresh All. |
| `calculate()` | Works out every formula again (Ctrl+Alt+F9). |
| `save()` | The bytes; parts that weren't changed come back as they were. |

## Reviewing an agent's edits

Excel has no tracked changes for cells, so a person can't accept or reject them in Excel as in Word. The review is the
workbook now beside the workbook as it was before the agent's edits (`src/agentreview.js`): for `openWorkbook`, the
file as opened; for `api.agent()`, the editor's workbook when the agent was made.

| Call | Gives or does |
|---|---|
| `changes()` | `{ sheets, structure, cells, more }`: sheets added, deleted or renamed (matched by their ids, so a rename is one change); rows and columns inserted or deleted and PivotTables made, from the agent's own log; each cell whose content changed (`edited`) or whose formula's result changed because of it (`result`, only where the file came with a result), with its value and formula before and after. Over 500 cells, the rest are counted in `more`. |
| `changes({ markdown: true })` | The same as a summary to show a person: the cells in a Before and After table, the results worked out again on one line. |
| `revert(refs)` | Cells and ranges put back as they were, formats included, as one edit. Refused on a sheet whose rows, columns or PivotTables the agent changed, where cells no longer line up. |
| `revertAll()` | Every edit taken back: the workbook as it was opened (not on `api.agent()`, where Undo does it). |

## The MCP server

`open-xlsx-editor-mcp` (`bin/oxe-mcp.mjs`) makes these calls tools for an AI assistant that speaks the Model Context
Protocol, over stdin and stdout, as open-docx-editor's server does for Word:

```json
{ "mcpServers": { "excel": { "command": "npx", "args": ["-y", "-p", "open-xlsx-editor", "open-xlsx-editor-mcp", "--root", "C:/Reports"] } } }
```

- **Tools:** `list_workbooks`, `open_workbook` (a password), `list_sheets`, `list_names`, `read_range` (Markdown),
  `read_cells` (JSON), `find`, `list_notes`, `write_cells`, `set_values`, `insert_rows`, `delete_rows`,
  `insert_columns`, `delete_columns`, `add_sheet`, `rename_sheet`, `delete_sheet`, `create_pivot`, `refresh_pivots`,
  `calculate`, `review_changes`, `revert_changes` and `save_workbook`.
- **The review:** the server's instructions ask the assistant to show `review_changes` before saving and to take back
  what the person rejects with `revert_changes`.
- **Safety:** it reads and writes only .xlsx and .xlsm files inside the folder it's given (`--root`, else the current
  folder). It saves to a new file unless asked to overwrite, never over an existing file otherwise, and keeps a
  workbook's extension. It makes no network requests. A workbook saved with a password stays encrypted.

Tests: `test/mcp.test.js` drives the server as an assistant does: reading, writing, the review, a cell taken back and
everything taken back, a PivotTable, saving a copy, the folder's limits, and a workbook with a password.

## Measured (10 November 2026)

The agent wrote twelve entries into a new sheet of `basic.xlsx`: `12`, `50%`, `1/2/2024`, `$1,234.50`, `TRUE`, `'007`,
`=SUM(1,2)`, `1e3`, `3:30 PM`, `(5)`, `=A1*2` and `1 1/2`. It saved the workbook, and a hidden Excel opened it
(read-only, its column widened to 30) and read each cell's `Value2`, `Text` and `Formula`. All 12 were the same as the
agent's `read()`.

At Excel's default width, `$1,234.50 ` shows as `########`. `read()` gives text at any width, so that cell differed
until the column was widened.

Tests: `test/agent.test.js` covers the standalone agent, and `test/editing-data.mjs` covers `api.agent()` (an entry drawn,
Ctrl+Z taking it back, an added sheet given a tab, its review and a cell put back). PivotTables made through the agent
are `pivotnew.js`'s, measured against Excel in docs/PIVOT.md.
