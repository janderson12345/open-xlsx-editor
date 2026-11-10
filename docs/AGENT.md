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
| `calculate()` | Works out every formula again (Ctrl+Alt+F9). |
| `save()` | The bytes; parts that weren't changed come back as they were. |

## Measured (10 November 2026)

The agent wrote twelve entries into a new sheet of `basic.xlsx`: `12`, `50%`, `1/2/2024`, `$1,234.50`, `TRUE`, `'007`,
`=SUM(1,2)`, `1e3`, `3:30 PM`, `(5)`, `=A1*2` and `1 1/2`. It saved the workbook, and a hidden Excel opened it
(read-only, its column widened to 30) and read each cell's `Value2`, `Text` and `Formula`. All 12 were the same as the
agent's `read()`.

At Excel's default width, `$1,234.50 ` shows as `########`. `read()` gives text at any width, so that cell differed
until the column was widened.

Tests: `test/agent.test.js` covers the standalone agent, and `test/editing.mjs` covers `api.agent()` (an entry drawn,
Ctrl+Z taking it back, an added sheet given a tab).
