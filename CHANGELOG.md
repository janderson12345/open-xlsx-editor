# Changelog

## 0.5.0

The first release. An editor for Excel workbooks that runs in the browser, matched to Excel by measuring Excel's own
output. `docs/FEATURES.md` has the details, ribbon tab by ribbon tab.

- **Opening and saving.** It opens .xlsx, .xlsm, .xltx, .xltm and .csv files, including workbooks saved with a
  password. Saving rewrites only the parts that were changed; every other part goes back byte for byte, macros
  included. A sheet can also be saved as CSV.
- **Looks.** The grid is laid out as Excel lays it out: column widths, row heights, fonts, fills, borders, number
  formats, conditional formatting (data bars, colour scales, icon sets), charts, pictures, notes and comments.
  Sparklines are kept in the file, but not drawn yet. Each was checked against Excel's PDFs.
- **Printing.**
  - Printing and `exportPdf` follow Excel's page setup: margins, scaling, fit to page, print titles, page breaks, and
    headers and footers with pictures.
- **Calculation.**
  - 500 of Excel's 519 functions, with dynamic arrays, LET, LAMBDA, table references and 3D references.
  - 1914 of 1914 test formulas give Excel's results.
  - Evaluate Formula, tracing precedents and dependents, and error checking.
- **Editing.**
  - Excel's ribbon, classic or simplified, its keys and its dialog boxes.
  - Format Cells, Find and Replace (formats included) and Go To Special.
  - Fill: the fill handle, series, Flash Fill and custom lists.
  - Paste Special, Sort, Filter, Remove Duplicates and Text to Columns.
  - Data Validation and Conditional Formatting rules.
  - Rows, columns and sheets inserted, deleted, hidden and moved, with references following.
  - Defined names, tables, PivotTables (made and refreshed), hyperlinks, notes and comments.
  - Sheet and workbook protection.
  - AutoCorrect.
- **Working with others.** Co-editing over Yjs (`open-xlsx-editor/collab`). An agent interface and an MCP server
  (`open-xlsx-editor-mcp`) let AI assistants read and edit workbooks, with every change reviewed before saving.
- **Hosting.** One script and one stylesheet, run under a strict Content-Security-Policy, with nothing fetched from
  another host. There are light and dark themes, and keyboard and screen reader support checked with axe.
