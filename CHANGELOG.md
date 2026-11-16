# Changelog

## 0.7.1

- The README shows the latest release, the npm version, CI and the licence as badges, on GitHub and on npm.

## 0.7.0

- **Tables.**
  - Drawn in their table style on screen, in print and in PDF: Excel's 60 built-in styles or the workbook's own.
  - Insert > Table, Ctrl+T and Format as Table, with the Table Design tab: name, resize, convert to range, header and
    total rows, banding, first and last column, filter buttons and the style gallery.
  - A table's own buttons filter and sort it. Excel reads back every table the editor makes (33 of 33).
- **Cell Styles:** Excel's 47 built-in cell styles in the Home tab's gallery.
- **Outlines.**
  - The outline bar beside the headings, sized as Excel's, with its level and +/− buttons.
  - Group and Ungroup (Shift+Alt+Right and Left), Show and Hide Detail, Auto Outline, Clear Outline, the automatic
    styles and Ctrl+8.
  - Data > Subtotal and Remove All, as Excel inserts and formats them.
- **ACCRINT** matches Excel in all 600 measured cases. FORECAST.ETS and its family stay missing: what was found of
  Excel's method is in docs/CALC.md.
- **Security.** A review of this release found and fixed:
  - crafted workbooks that hung the page: a negative table-style stripe size, a whole-sheet table, a huge outline
    level or column span;
  - XML that a crafted theme font name could add to a saved file.
  - Table looks are now worked out a screen at a time.
  - MAX, MIN and the other places that took the largest of a list no longer fail past 125,000 items.
  - ACCRINT gives #NUM! past 9999-12-31, as Excel does.
- **Maintainability:** one shared XML escape, the outline's cell styles declared as data, every module listed in
  docs/ARCHITECTURE.md, and more robust browser tests.

## 0.6.0

- **Charts.**
  - **Select Data:** change a chart's data range, Switch Row/Column, and add, edit, remove and reorder series (34 of
    34 changes written as Excel writes them).
  - **Format Chart (Ctrl+1):** axis titles; fills, lines and fonts of every part; combo charts and the secondary axis
    (33 of 33 as Excel writes them). Charts with their own formatting are drawn in it.
- **Sparklines** are drawn on screen, in print and in PDF export, and can be inserted and edited on the Sparkline tab,
  with Excel's 36 styles.
- **Draw Borders and More Colors:**
  - Draw Border, Draw Border Grid and Erase Border, with Line Color and Line Style;
  - a custom colour for fill, font and line, kept in Recent Colors.
- **Function Arguments:**
  - Excel's box after Insert Function, with a description of every function and argument;
  - references picked on the sheet, each argument's kind shown, and text quoted as Excel quotes it;
  - AutoSum > More Functions leads into it.
- **The Name Box** shows a selected name, and lists the names to go to.
- **Rich text** is shown formatted while a cell is edited.
- **Saved sorts:** the last sort is saved with the sheet as Excel saves it, in a table's part for a table. It follows
  inserted and deleted rows and columns.
- **Fonts:** metric twins for all 151 open-metric-fonts families where the computer lacks the font. PDF export embeds
  the served twins.
- **Fixed:**
  - a rich text run's colour, which was saved as `[object Object]`;
  - VLOOKUP with an empty table argument, which now gives #N/A as Excel does.
- **Housekeeping:**
  - the largest source files split into smaller modules;
  - the workflows' actions pinned to their release commits, with Dependabot keeping them current;
  - a favicon for the demo.

## 0.5.2

- **Fixed: right click on a chart.** On Windows the chart's menu (Edit Chart..., Delete) was replaced by the cell
  menu (Cut, Copy, Paste...) when the button was let go. The chart's menu now stays open.

## 0.5.1

The first release published from GitHub, with npm provenance. It has the security fixes below; 0.5.0, published by
hand from an earlier commit before them, should not be used.

An editor for Excel workbooks that runs in the browser, matched to Excel by measuring Excel's own
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
- **Security.** No CodeQL code scanning or Dependabot alerts are open. Typed entries and formulas that could make
  the parser slow are read in linear time, and rich text attributes are escaped (`test/scanning.test.js`).
