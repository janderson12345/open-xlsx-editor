# Queries and connections

Data > Queries & Connections as far as a page can do it: the workbook's Power Query queries and connections listed,
their M shown and changed, and Refresh writing rows a host page gives into a query's table as Excel's Refresh writes
them. A page can't reach a database, a file share or most web servers, and can't run Power Query's M engine, so the
rows come from the host (the `onRefreshQuery` option). Code: `src/ooxml/queries.js` (the parts),
`src/editor/queriespane.js` (the pane, Refresh, the Advanced Editor).

## Where Excel keeps them

- `xl/connections.xml`: one `connection` each. A Power Query query loaded somewhere has one named "Query - Name"
  over the Mashup OLE DB provider (`Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Name`), its
  command `SELECT * FROM [Name]`; one loaded to the data model is type 100 with an `x15:connection`, beside the model's
  own ThisWorkbookDataModel.
- A query loaded to a table: the table (`tableType="queryTable"`) has a queryTable part naming the connection by id;
  its `queryTableFields` map the query's columns to the table's (`queryTableFieldId`), and the sheet has a hidden
  name ExternalData_N over the query's columns.
- The M: `customXml/itemN.xml` holds a `DataMashup` element (UTF-16 with a BOM, as Excel writes it) whose text is the
  base64 of [MS-QDEFF]'s binary: a version, then four pieces each after its length: a package (a zip whose
  `Formulas/Section1.m` is the section document, `shared Name = ...;` for each query), permissions, metadata (each
  query's load settings: FillEnabled, FillToDataModelEnabled, FillCount, FillTarget...) and permission bindings
  (a DPAPI-protected checksum of the rest).

## In the editor

- **Data > Queries & Connections** opens the pane, docked under the ribbon as the Watch Window is. Its Queries tab
  lists each query with what Excel's pane says under it ("4 rows loaded." from the table's rows, or the metadata's
  FillCount for one in the data model; "Connection only.") and where it loads (a table, with its sheet and range;
  the data model; a connection only), with Refresh, Edit (View when the workbook is read-only) and Go to (the
  table's first cell). Its Connections tab lists every connection: name, description, type, the query it reads and
  where it loads. The pane's Refresh All, and Data > Refresh All (Ctrl+Alt+F5), refresh every query and connection
  loaded to a table, then the PivotTables; Refresh (Alt+F5, and the cell menu's Refresh on a query's table) the one
  at the active cell.
- **Refresh** asks the host for each query's rows: `onRefreshQuery({ name, kind, connection, m, target })` (README,
  `src/index.d.ts`), `kind` 'query' for a Power Query query or 'connection' for another connection loaded to a table,
  `connection` its name, description, type, connection string and command, `m` the query's M, `target` the table's
  sheet, range and name. It returns (or resolves to) rows: arrays with the column names first, objects whose keys are
  the columns, or `{ columns, rows }`; a value is a number, text, true/false, a Date (written as its serial number) or
  null (an empty cell). Every query refreshed together is one step of Undo. On a protected sheet it's refused, as a
  PivotTable's refresh is.
- Without the option, Excel for the web's words: "Unable to refresh data for a data connection in the workbook. The
  following connections failed to refresh:" and the connections' names. A host that throws or rejects gets the same
  words with its message after them, and nothing changes.
- **Edit** opens the Advanced Editor over the query's M; Done writes it back (below). A formula that would break the
  section document (a string, comment or bracket left open, or a `;` that would end the query early) is refused with
  why, and nothing changes.

## Refresh as Excel writes it

Measured over COM (`tools/excel-compare/excel-queries-refresh.ps1`, `test/queries-refresh-excel.json`): the fixture's
query given a new result, its table refreshed with `QueryTable.Refresh`, the file Excel saved compared. Here the
same rows are written and the sheet's values and formulas, the table's columns, the queryTable's fields and the
ExternalData name compared with Excel's (`test/queries.test.js`), seven cases of seven:

- Rows: the table's data rows become as many as the result's, inserted or deleted as cells of the table's columns
  (Insert with Shift cells down), so what is below the table in those columns moves and nothing beside it does. A
  result with no rows leaves one empty data row.
- Columns: the query's columns become the result's, inserted or deleted as cells of the table's rows (Shift cells
  right), so what is right of the table moves. A column kept by name keeps its tableColumn (id, formats) and its
  queryTableField; a new one takes the lowest id no column of the table has (with `uniqueName` the same) and the
  queryTable's `nextId`; one gone leaves both parts; they follow the result's order.
- Columns of the user's (a column added to a query's table, `dataBound="0"`, `unboundColumnsRight`/`Left`) stay at
  their side; a calculated one's formula goes into the new rows.
- ExternalData_N covers the query's own columns (not the user's), the header row and the data rows.

## Changing a query's M

Section1.m is written again with the query's new text in place of the old; its package's other parts, the
permissions, the metadata and the permission bindings stay byte for byte, as does the part's text around the base64
and its UTF-16. Measured (`tools/excel-compare/queries-writeback.mjs` then `excel-queries-check.ps1`,
`test/queries-writeback-excel.json`): four files with M changed here (more rows, steps renamed with a comment and a
new step, two queries at once, text with accents and doubled quotes), each opened in Excel without repair,
`Workbook.Queries(...).Formula` the text written here exactly, every connection refreshed without an error (the
table's rows the new M's), and the M kept after Excel saved it. The bindings no longer match the package after an
edit; Excel then takes the permissions as unset (as [MS-QDEFF] says it does), which for these files changed nothing.

## Kept byte for byte

Everything not changed: a workbook with queries opened and saved untouched has every part as it was
(`test/queries.test.js`), the data model (`xl/model/item.data`) included. A refresh rewrites the sheet, the table, the
queryTable part and the workbook's names; an M edit, the DataMashup part.

## What isn't possible here

- **Running M.** The editor doesn't evaluate Power Query's M: not even the simple steps (Csv.Document, Web.Contents
  through the host, Table.PromoteHeaders, Table.TransformColumnTypes, Table.SelectRows). Doing it means M's lexical
  grammar and expressions (let, each, records, lists, types, `#"..."` names, operators and their null rules) and each
  function's own behaviour, including type conversion by culture; a partial evaluator would give rows that differ
  from Excel's without saying so. The host gets the M and decides how to answer.
- **Reaching the data.** A page can't open a database connection (OLE DB, ODBC), read a file path, or fetch most web
  sources (CORS, credentials). That is what the host option is for.
- **The data model.** `xl/model/item.data` is an Analysis Services backup that can't be read or written here: queries
  loaded to the data model aren't refreshed, and PivotTables or CUBE functions over the model aren't worked out.
- **What a refresh doesn't change here that Excel does:** the query's metadata in the DataMashup (FillCount,
  FillLastUpdated, column types and names, the error code), the number formats and dxfs Excel gives typed columns,
  connection settings acted on (refresh on open, background refresh, refresh every N minutes), and the Refresh
  dialog's progress. Connection-only queries aren't refreshed (there is nowhere to write them).
- **Other changes to queries:** a new query, a query deleted, renamed or duplicated, Load To... (changing where it
  loads), query properties and Query Dependencies aren't here; nor is the Power Query Editor's step list.
- **Query tables without a table**, from older Get External Data (a text, web or database query on a sheet), are
  listed with their connections but not written by Refresh.
- Excel for the web's refusal is from its own words for a connection it can't refresh; COM can't show it, so it isn't
  measured.

## Files

- `tools/excel-compare/excel-queries.ps1`: makes `test/queries-excel.xlsx` (three queries: Sales to a table, Big a
  connection only, Model to the data model; docProps' names put back as the editor's) and `test/queries-excel.json`.
- `tools/excel-compare/excel-queries-refresh.ps1`: `test/queries-refresh-excel.json`, Excel's refresh of changed
  results (grow, shrink, columns, wider, narrower, a column of the user's, the same).
- `tools/excel-compare/queries-writeback.mjs`, `queries-writeback-cases.json`, `excel-queries-check.ps1`:
  `test/queries-writeback-excel.json`, the editor's M read, refreshed and kept by Excel.
- `test/queries.test.js` (unit), `test/queries.mjs` (browser: the pane, Refresh All through the host, Ctrl+Z,
  Alt+F5, the cell menu, a failing host, the Advanced Editor, an editor without the option).
