# Changelog

## Unreleased

- **Fix:** rows of wrapped text break their lines where Excel does. Lines are measured in the whole-pixel character
  widths Windows gives Excel, with Excel's own margin (7px for a short line in Calibri 11, a pixel more for every
  28px of text past 84), a hyphen is a place to break, and the spaces or hyphen a line ends with don't count toward
  its width. Under 8pt text is measured in Small Fonts, as Excel draws it; rich text is measured run by run, each
  line as tall as its tallest run. Measured in Excel itself at every column width. Against Excel on 146 customer
  workbooks, rows that differ went from 98 to 59 of 9,429 (rows a line or more out: 48 to 3), with none made
  wrong. Security review still to do before release.
- **Fix:** wrapped text in a hidden column no longer makes its row hundreds of points tall. It was counted at no width,
  a line for each character; Excel gives it one line. Five rows of customer workbooks at 390pt are Excel's 30 to 72pt
  again. Security review still to do before release.

## 0.9.0

- **Fix:** screen readers give each cell its own column's heading. NVDA in Edge and Chrome read column B's cells under
  column C's heading, because a screen reader matches a cell to its heading by its place in the row. The heading row
  now starts with the corner as each row starts with its number, and every column in view has a cell in each row.
- **Fix:** row heights Excel works out match it again on wrapped text. A formula's text result (such as
  =CONCATENATE) and a word wider than its column now count, lines are counted at Excel's text widths, a merged title's
  font counts toward its row, and a thick bottom border adds Excel's 0.75pt. A crosswalk workbook went from 41 to 60 of
  60 rows as Excel's.
- **Faster opening:** only sheets with a formula saved without its result are parsed on opening (found by scanning
  their bytes); the rest wait until shown. Large workbooks whose other sheets are big open in seconds instead of up to
  a minute.
- **Fast sheet reading:** a sheet's cells are read straight from its bytes instead of a parsed XML tree, and rows of
  formatted empty cells are kept as bands. Workbooks whose sheet is formatted down to the last row (millions of empty
  styled cells) open in 3 to 5 seconds instead of 30 to 57, and in Node without running out of memory. A sheet's XML
  is parsed only when it's first edited.
- **Excel's own row heights and column widths** for the 25 fonts workbooks use most, measured at every half point
  from 6 to 36pt, instead of a formula that was a pixel off for most fonts below 11pt. Fixes Lato and Aptos Narrow 12
  sheets whose columns were 11 to 12% off.
- **Fix:** a row of long wrapped text stops at 409.5pt, Excel's tallest row, instead of growing to thousands of
  points.
- **A sheet's standard row height is Excel's:** the stored default only when the file sets it, else the tallest font
  in the row's formatting, as Excel works it out. On 146 customer workbooks, rows that differ from Excel went from 995
  to 102 and workbooks with every row exact from 68 to 95.
- **Clearer refusals:** a workbook protected by rights management (a sensitivity label) says only Excel, signed in
  with access, can open it; a damaged file is no longer called too large.
- **Security:** a row height pass over a sheet whose rows of formatted empty cells alternate styles column by
  column now looks each column's style up by binary search: a crafted sheet of such rows took billions of steps.

## 0.8.3

- **open-metric-fonts 1.1.0.** `OMF_VERSION` is 1.1.0 and the README's setup steps serve that release, whose fonts
  are redrawn without the overlapping and badly spaced letters of 1.0.1. Letter widths and line metrics are unchanged,
  so cells, rows and page breaks are the same; a site moves by serving the 1.1.0 bundle in the folder it names.
- A weekly workflow opens an issue when open-metric-fonts has a newer release.
- **Power Query** is recorded as unsupported, by decision, as VBA is: M isn't run in the editor and the data model
  isn't refreshed. Queries are shown and their M changed here, and refreshed through the host (docs/QUERIES.md).

## 0.8.2

- **Fix:** filter buttons, note marks and note boxes, and a list rule's button and input message are drawn at their
  cells again. They were a row heading's width too far left, so a workbook with wide columns had its filter buttons
  over the headings' text.
- **Notes.** Show/Hide Note and Show All Notes are saved in the file, as Excel saves them, and Review > Notes > Convert
  to Comments turns notes into threaded comments.
- **Security:** a workbook whose parts say they unpack to more than 1 GB is refused before anything is unpacked.

## 0.8.1

- **Pictures, shapes and icons are drawn** on the grid, in print and in PDF, at Excel's place and size. They can be
  selected, moved, sized and deleted as charts can, and Go To Special > Objects selects them all.
- **Rich text.** Format Cells (Ctrl+1) for the selected characters while a cell is edited, and copy and paste keep a
  cell's runs.
- **Cell Styles:** Duplicate, Delete and Merge Styles.
- **Several selected ranges.** Fill, AutoSum, row and column sizes, hiding, Print Selection and the other commands
  that take a selection now act on each range as Excel does, or refuse with Excel's message.
- **Fill > Across Worksheets**, with sheets grouped on the tabs.
- **Hyperlinks:** Browse for File, and Create New Document through the new `onCreateDocument` option.
- **Queries & Connections.** A pane listing the workbook's connections and Power Query queries; Refresh and Refresh
  All through the new `onRefreshQuery` option; a query's M text shown and edited.
- **Display.**
  - East Asian built-in date formats are shown as Excel shows them, not as General.
  - Column widths follow the Normal style font's own digit width, not always Calibri's.
  - Gridlines are hidden under text that runs on into empty cells.
- **Fix:** a merged range is drawn with only its outline's borders. The covered cells' own borders no longer show
  inside it.
- **Macros.** VBA is not supported, by decision: macros are never run, shown or edited. An .xlsm keeps its VBA
  project byte for byte, so Excel still has it.

## 0.8.0

- **Co-editing piece by piece.** Formats, row and column settings, inserted and deleted rows and columns, the sheet
  list and each sheet's cells are shared as separate changes, not as the whole workbook.
- **Locale.** The `locale` option and `setLocale()`: numbers, dates and formulas shown and typed with the locale's
  separators, and the function tip uses the locale's list separator.
- **Text files.**
  - The Text Import box: opening a .txt file, and Data > From Text/CSV at the active cell. Encodings, delimited or
    fixed-width fields, column types and the start row are all handled.
  - The Save As box (F12): .xlsx, .xlsm, .csv with a separator and encoding, .txt and .pdf. Text files are written
    byte for byte as Excel writes them.
- **Clipboard and links.**
  - Pasting keeps other programs' formatting, read from the clipboard's HTML as Excel's Paste reads it.
  - Paste Link between open workbooks, Edit Links, and names that refer to other workbooks.
- **Formatting.**
  - Part of a cell's text in its own font, size or colour, from the Font group while editing.
  - New Cell Style and Modify, in Excel's Style box.
  - @mentions in threaded comments.
  - Find by format, with Choose Format From Cell.
  - Rotated, stacked and shrunk text printed as Excel prints it.
- **Inserting and deleting rows, columns and cells.** These now move what Excel moves:
  - sparklines and data validation rules;
  - Excel 2010 conditional formats and validation rules;
  - pictures and charts that don't move with cells;
  - slicers and timelines, which no longer block the insert or delete;
  - Circle Invalid Data's circles.
- **Other changes.**
  - Notes and threaded comments go with their cells in a sort.
  - SUBTOTAL leaves out rows a filter hides.
  - AutoSum works in filtered lists, in tables and on several areas.
  - Freeze Panes works on a scrolled or split window.
  - Validation rules that read another sheet now work.
- **Formulas.**
  - Regular expressions are matched as PCRE2, with a limit on how much work one search may take.
  - Text functions count an emoji as one character in workbooks made by Excel.
  - DATE and DAYS360 treat February 1900 as Excel does, and ACCRINT and ARRAYTOTEXT give Excel's answers in their
    edge cases.
- **Fixes.**
  - Fonts written with `<u val="none"/>` are no longer drawn underlined.
  - Rich text keeps font settings the editor doesn't model.
- **Security.** A review of this release found and fixed:
  - a relationship id from a file used as a regular-expression pattern;
  - validation rules from a file that could write unexpected XML;
  - clipboard CSS that took far too long to read;
  - regular-expression replaces whose total work had no limit.

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
