# Architecture

A browser viewer and editor for .xlsx files in plain JavaScript. Four layers that don't know about each other's
insides:

- `src/ooxml/` reads and writes the file. No DOM rendering; runs in Node too (unit tests use @xmldom/xmldom).
- `src/calc/` works out formulas: parsing, functions, recalculation (docs/CALC.md). No DOM; runs in Node.
- `src/editor/` is the view: the ribbon, the grid, selection, editing.
- `src/engine/` holds measured data: `metrics.js`, the vertical metrics of installed fonts (`tools/fontmetrics.mjs`).

## Module map

| Module | What it does |
|---|---|
| `src/index.js` | `mountXlsxEditor(host, { bytes, fileName, theme })`: opens a workbook and shows the sheet it opens on, with the sheet tabs; `showSheet` changes sheet; `insertSheet`, `renameSheet`, `moveSheet`, `hideSheet`, `setTabColor`, `deleteSheet`; the tab menu's actions; `save()` gives the bytes back. Types in `src/index.d.ts`. |
| `src/ooxml/package.js` | The OPC package (the zip): every part kept byte for byte, relationships, content types; only changed parts rewritten. From open-docx-editor. |
| `src/ooxml/xml.js` | Parsing and serialising XML (injectable DOMParser for Node), namespace-aware lookups defaulting to SpreadsheetML, `_xHHHH_` decoding. From open-docx-editor. |
| `src/ooxml/units.js` | The one home for unit conversions: EMU, points, pixels, Excel's column width unit. From open-docx-editor, extended. |
| `src/ooxml/refs.js` | A1 references and ranges; rows and columns are 0-based in code. |
| `src/ooxml/workbook.js` | `Workbook`: sheets in order (`canShow`, `tabColor` read from the part's start), shared strings, defined names, each sheet's model read on first use, `setCell`, `setFormula` and `apply` (then recalculation), `calculate`, `toBytes` (only dirty sheets serialised, with their XML declaration; calcChain.xml dropped when formulas change; `fullCalcOnLoad` when a formula couldn't be worked out). |
| `src/ooxml/sheetops.js` | Sheets added, renamed, moved, hidden, coloured and deleted as Excel does it, with Excel's refusals; references rewritten (formulas, names, charts, hyperlinks, pivot sources) or made #REF!, sheet-local names and the active tab kept with their sheets, a deleted sheet's own parts removed (docs/SHEETS.md). |
| `src/ooxml/stylesedit.js` | Cell formats changed as Excel writes them: a copy of the cell's format, font, fill or border with the change, an equal one reused, a new one at the end; Excel's Percent, Comma and Currency cell styles; a styles part made when there is none (docs/FORMATS.md). |
| `src/ooxml/formatops.js` | What each Home button does to a selection, as Excel does it (toggles from the active cell, font sizes, indent, border presets and the neighbour's facing side, decimals): the cells' new formats, for the editor to apply and undo. |
| `src/ooxml/mergeops.js` | Merge & Center, Merge Across, Merge Cells and Unmerge as Excel does them: the values, formats and merged ranges that change, for the editor to apply and undo as one step. |
| `src/ooxml/clearops.js` | Home > Clear (All, Contents, Formats) as Excel does it, merged ranges unmerged, part of one refused. |
| `src/ooxml/clearparts.js` | What Clear takes besides cells: notes, threaded comments and hyperlinks, as Excel does (docs/FORMATS.md). |
| `src/ooxml/stats.js` | The status bar's Average, Count and Sum of the selection, by Excel's rules. |
| `src/ooxml/fillops.js` | Fill Down, Right, Up and Left, and the fill handle (AutoFill) as Excel does them: series, lists, dates, formulas moved, formats repeated (docs/FILL.md). |
| `src/ooxml/charts.js` | A sheet's charts read from its drawing and chart parts, their data from the cells, and Excel's automatic value axis; drawn by `src/editor/chartview.js` (docs/CHARTS.md). |
| `src/ooxml/tables.js` | The workbook's tables, and what a table reference (Fruit[Qty], [#Totals], [@Qty]) stands for. |
| `src/ooxml/cfb.js` | Compound files (MS-CFB): every stream read by its path; written as version 3, small streams in the mini stream. |
| `src/ooxml/encryption.js` | Workbooks saved with a password (MS-OFFCRYPTO agile): decrypted and encrypted with @noble/hashes and @noble/ciphers; PasswordError; docs/ENCRYPTION.md. |
| `src/ooxml/printarea.js` | A sheet's print area and print titles, from the workbook's defined names (docs/PRINT.md). |
| `src/ooxml/headerfooter.js` | Headers and footers: Excel's codes read into sections and runs, fields filled in per page (docs/PRINT.md). |
| `src/pdf.js` | PDF export from the print layout (pdf-lib, fontkit): fills, borders and cell text at the layout's places, clipped to their cells; docs/PDF.md. |
| `src/collab.js` | Co-editing over Yjs: `connectWorkbook(api, ydoc)`; cell contents shared cell by cell from `api.onEdit`, made with the `remoteContents` command; any other change shared as the workbook's bytes; docs/COLLAB.md. |
| `src/collab-peers.js` | Others' cursors in a co-edited workbook: this editor's selection put in the provider's awareness (throttled, from `api.onSelect`), the others' read back for the sheet shown and given to the grid; each person's colour and the flag's text colour by contrast. |
| `bin/oxe-mcp.mjs` | The MCP server: the agent interface as tools for an AI assistant, over stdin and stdout, within one folder; docs/AGENT.md. |
| `src/agentreview.js` | An agent's edits reviewed: the workbook now beside the one opened, cell by cell; docs/AGENT.md. |
| `src/agent.js` | The agent interface: a workbook read and edited in plain terms, without a browser (`openWorkbook`) or on the editor's workbook (`api.agent()`, each edit one undo step); docs/AGENT.md. Types in `src/agent.d.ts`. |
| `src/ooxml/anchors.js` | What sits on a sheet at a place of its own, moved with inserted and deleted rows and columns: notes and threaded comments, the anchors of pictures, charts and shapes, tables, pivot tables and their caches' sources (docs/ROWCOL.md). |
| `src/ooxml/pivotnew.js` | A new PivotTable (Insert > PivotTable), its parts written as Excel writes a new one, then laid out by pivot.js's refresh; docs/PIVOT.md. |
| `src/ooxml/pivotfields.js` | A PivotTable's fields moved after it is made, as Excel's field list moves them, then laid out by pivot.js's refresh; docs/PIVOT.md. |
| `src/ooxml/pivot.js` | PivotTables: read from their parts, refreshed from their source as Excel's Refresh does (cache, records, items, compact layout); docs/PIVOT.md. |
| `src/ooxml/cfedit.js` | Conditional formatting rules added as Excel's menu adds them, with its preset formats as dxfs; Clear Rules (docs/CONDFORMAT.md). |
| `src/ooxml/dvedit.js` | Data validation rules written as Excel's Data Validation box writes them; ranges split around a new rule; a rule in the box's terms (docs/VALIDATION.md). |
| `src/ooxml/noteedit.js` | Notes written and deleted as Excel's New Note writes them: the comments part, the VML box, the parts and relationships made when the sheet has none (docs/NOTES.md). |
| `src/editor/reviewtab.js` | The Review tab: Notes. |
| `src/editor/datadialogs.js` | Excel's boxes for data: Sort, a filter's list, a list's choices, Data Validation, the Conditional Formatting menu's rule boxes (with `dialogs.js`'s modal and buttons). |
| `src/editor/geometry.js` | A sheet's rows and columns in pixels, on screen or printed (`sheetGeometry`), and the grid's layout constants. |
| `src/editor/cells.js` | A cell's element as the grid and print draw it (text as shown, aligned, wrapped, running on, indented), merged ranges, numbers too wide shown as #s, AutoFit's width. |
| `src/editor/gridlayers.js` | The grid's layers for the part in view: column headings, frozen panes, gridlines, rows of cells. |
| `src/editor/dom.js` | Small DOM helpers: CSS custom properties in pixels, a div with a class. |
| `src/editor/acts.js` | The ribbon's actions (the act dispatcher) and the boxes and menus they open: Insert and Delete, Paste Special, Sort, the validation question, a list's choices, a filter's list, the cell's and headings' menus, Row Height and Column Width; given `index.js`'s state and helpers as one object. |
| `src/editor/datatab.js` | The Data tab: Refresh All and Refresh; Sort & Filter. |
| `src/ooxml/special.js` | Go To Special: the cells of a kind Excel's box finds (constants, formulas, blanks, the current region, differences, precedents and dependents, rules), as ranges (docs/GOTOSPECIAL.md). |
| `src/ooxml/names.js` | Defined names made, changed (formulas following) and deleted as Excel's Name Manager does, with Excel's rules for a name; docs/NAMES.md. |
| `src/ooxml/goto.js` | What the Name Box goes to: a cell, range, whole lines, another sheet's, a defined name or a table (by name or table reference); the names Go To lists; Excel's words otherwise. |
| `src/ooxml/autosumops.js` | AutoSum and its Average, Count Numbers, Max and Min as Excel does them: where the totals go, the run of numbers each takes, subtotals, the proposal left to edit on one cell (docs/AUTOSUM.md). |
| `src/ooxml/validation.js` | Data validation judged as Excel judges it: blanks, types, Excel's comparisons, text length, typed and cell lists, custom formulas; a list's choices (docs/VALIDATION.md). |
| `src/ooxml/condformat.js` | Conditional formatting worked out as Excel draws it: rules in priority order with Stop If True, Excel's comparisons, range rules, colour scales; each cell's fill and font (docs/CONDFORMAT.md). |
| `src/ooxml/notes.js` | A sheet's notes and threaded comments read as Excel 365 writes them (the placeholder note of a thread left out); shown by `src/editor/notemarks.js` (docs/NOTES.md). |
| `src/ooxml/filterops.js` | AutoFilter as Excel does it: its range, the rows a filter hides by shown values, a column's list, the autoFilter element and the _xlnm._FilterDatabase name as Excel writes them (docs/FILTER.md). |
| `src/ooxml/sortops.js` | Sort as Excel does it: its order of values and of text, the range around a cell, the header guess, formulas moved as a copy moves them, borders left, Excel's refusals (docs/SORT.md). |
| `src/ooxml/clipops.js` | Copy, cut and paste as Excel does them: a snapshot of cells, pasted with formulas moved and repeated over the selection, or moved with every formula that reads them; Paste Special (contents, formats, operations, Skip blanks, Transpose) and Paste Link; text from other programs read as typed (docs/CLIPBOARD.md). |
| `src/ooxml/rowcol.js` | Whole rows and columns, or cells with the rest of their columns or rows shifting, inserted and deleted as Excel does it: the sheet's cells, widths, merges and ranges moved, new lines formatted from the one before, every formula of the workbook following, Excel's refusals (docs/ROWCOL.md). |
| `src/ooxml/sizes.js` | Column widths, row heights and hidden lines written as Excel writes them: widths in whole pixels over the digit width, col elements split where a range cuts them, AutoFit Row Height as the fixed height taken away (docs/SIZES.md). |
| `src/ooxml/findops.js` | Find and Replace as Excel does them: the formula bar's text or the shown text searched, Excel's wildcards, the order from the cell after the active one, replacements entered as typed (docs/FIND.md). |
| `src/ooxml/view.js` | The sheet view as Excel writes it: frozen panes (a pane element with a selection for each pane) and the zoom (docs/VIEW.md). |
| `src/ooxml/sheet.js` | A worksheet part to a sparse model: cells by row then column, row heights, column widths, merges, defaults, array formula ranges. Implied `r` positions followed; a shared formula's later cells get their own formula text. |
| `src/ooxml/input.js` | A typed entry as Excel reads it (numbers with separators, currency, percentages, exponents, fractions; dates and times; 'text; =formulas), the format it gives the cell, and a cell's text as the formula bar shows it (docs/ENTRY.md). |
| `src/ooxml/entry.js` | What typing text into a cell changes (its content, and its format when Excel would give one), for typing, pasting text and Replace. |
| `src/ooxml/sheetedit.js` | Changing cells in the worksheet's XML in place: values, formulas, formula results (the formula kept), shared formulas written out cell by cell. Implied positions pinned first, rows indexed, row and cell inserted in order, `spans` hint dropped when wrong, `dimension` grown. |
| `src/calc/tokens.js` | Formulas as tokens; moving a formula (`shiftFormula`), a sheet's references renamed or made #REF! (`renameSheetRefs`, `sheetPrefix`), references following cells cut and pasted (`moveRefs`) and rows or columns inserted or deleted (`shiftRefs`). |
| `src/calc/fileform.js` | The typed form of a formula to the file's (`normalizeFormula`: `_xlfn.` prefixes, Excel 365's dynamic-array and trim-reference forms) and back (`displayFormula`). |
| `src/calc/parse.js` | Tokens to a tree by Excel's precedence. |
| `src/calc/values.js` | Values (numbers, text, booleans, errors, blanks, `Area`, `Arr`) and Excel's coercions, comparison, numbers as text, the near-zero rule. |
| `src/calc/evaluate.js` | Operators, references, names and function calls; implicit intersection outside array formulas, element by element inside. |
| `src/calc/legacyat.js` | The @s Excel 365 shows in a formula of the old kind, where it cuts a range to one value: found by working it out at its cell (docs/DYNAMIC.md). |
| `src/calc/iterate.js` | Circular references: left out with iterative calculation off, gone round with it on, as Excel does (docs/CALC.md). |
| `src/calc/threed.js` | 3D references (Jan:Apr!A1) and the intersection operator, and which functions take a 3D reference (docs/CALC.md). |
| `src/calc/functions.js` | The worksheet functions, as Excel defines them: maths, statistics, logic, information, and the helpers the others share. |
| `src/calc/functions-text.js` | Text, and dates and times. |
| `src/calc/functions-lookup.js` | Lookup and reference, and the conditional sums and counts. |
| `src/calc/functions-dynamic.js` | Excel 365's dynamic array functions: SEQUENCE, SORT, SORTBY, UNIQUE, FILTER, TRANSPOSE, @ and D1#; docs/DYNAMIC.md. |
| `src/calc/functions-arrays.js`, `functions-finance.js`, `functions-stats.js` | Excel 365's array shaping and LAMBDA helpers; loans, depreciation, rates and working days; statistics, distributions, numerals and a cell's formula and sheet (docs/CALC.md). |
| `src/calc/functions-db.js`, `functions-bonds.js`, `criteria.js` | The database functions and their criteria (shared with Advanced Filter); coupons, bond prices and yields, and the discounted securities (docs/CALC.md). |
| `src/calc/functions-oddbonds.js` | ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD: securities with an odd first or last period (docs/CALC.md). |
| `src/calc/functions-groupby.js` | GROUPBY, PERCENTOF, RANDARRAY and TRIMRANGE, as Excel 365 works them out, and what GROUPBY and PIVOTBY share (docs/DYNAMIC.md). |
| `src/calc/functions-pivotby.js` | PIVOTBY: several row and column fields, subtotals, sorting, relative_to (docs/DYNAMIC.md). |
| `src/calc/functions-bessel.js` | BESSELJ, BESSELY, BESSELI, BESSELK by the approximations Excel uses (docs/CALC.md). |
| `src/calc/functions-amort.js` | AMORLINC and AMORDEGRC, the French depreciation, as Excel works them out (docs/CALC.md). |
| `src/calc/functions-regex.js`, `regex.js` | REGEXTEST, REGEXEXTRACT, REGEXREPLACE: Excel's PCRE2 patterns and replacement text turned into JavaScript's (docs/CALC.md). |
| `src/calc/functions-convert.js` | CONVERT, with Excel's units, factors and prefixes (docs/CALC.md). |
| `src/calc/functions-complex.js`, `trig.js` | The complex-number functions (COMPLEX, IMSUM and the rest); sine and cosine reduced by the x87's π, as Excel's (docs/CALC.md). |
| `src/calc/functions-dist.js`, `functions-agg.js`, `functions-eng.js` | The t, chi-squared, F, gamma, beta and other distributions and their tests; AGGREGATE, matrices and LINEST and its kin; bits, bases, hyperbolics and more money (docs/CALC.md). |
| `src/calc/spill.js` | Spilling: a dynamic formula's array placed from its cell, #SPILL! when blocked, cells left emptied, readers worked out again. |
| `src/ooxml/spillxml.js`, `src/ooxml/dynamicmeta.js` | Dynamic array formulas in the file as Excel 365 writes them: the anchor's `t="array"`, `ref` and `cm`, and `xl/metadata.xml`'s XLDAPR entry. |
| `src/ooxml/watches.js`, `src/editor/watchwindow.js` | The Watch Window: watched cells, saved in cellWatches as Excel saves them (docs/CALC.md). |
| `src/calc/stepper.js`, `src/editor/evalbox.js` | Evaluate Formula: a formula worked out a step at a time, the next part underlined (docs/CALC.md). |
| `src/ooxml/errorfix.js`, `src/editor/errormenu.js`, `errorbox.js` | The error button's menu and the Error Checking box: Ignore Error (ignoredErrors as Excel writes them) and the fixes (docs/CALC.md). |
| `src/ooxml/errorcheck.js` | Excel's error indicators: which cells break its default error-checking rules (docs/CALC.md). |
| `src/ooxml/trace.js`, `src/editor/tracing.js`, `arrows.js` | Trace Precedents and Trace Dependents: what a formula reads and what reads a cell, a level at a time, drawn as Excel's tracer arrows (docs/CALC.md). |
| `src/ooxml/calcmode.js` | Calculation Options: Automatic or Manual (changes waiting for F9), Calculate Now and Calculate Sheet, iterative calculation (`iteration`, `setIteration`), saved in calcPr (docs/CALC.md). |
| `src/calc/engine.js` | `Calc`: what each formula reads, what a change reaches, the order to work them out in, results written back. |
| `src/ooxml/strings.js` | Shared string table and inline strings; phonetic runs skipped. |
| `src/ooxml/styles.js` | Fonts, fills, borders and cell formats (cellXfs) from styles.xml; `fontFamily` resolves scheme fonts to the theme's. |
| `src/ooxml/theme.js` | The theme's fonts and colours (Excel 365's default theme without a theme part); `resolveColor`, the one colour resolver (theme swap, integer-HLS tints, indexed palette). |
| `src/ooxml/numfmt.js` | What a cell shows (`cellDisplay`): the General format, and format codes through format.js. |
| `src/ooxml/format.js` | Excel's number format codes: sections, digits, scaling, scientific, fractions, dates and times, text; built-in formats; serial dates (1900 and 1904). |
| `src/editor/grid.js` | The virtualised grid: only rows and columns in view (plus overscan) in the DOM; sticky headings; text running on over empty cells; merged ranges; gridlines; frozen panes (their own layers, sticking in place) and zoom. An ARIA grid (rows, row and column headers, gridcells with indexes). |
| `src/editor/print.js` | The printed sheet: used range split into pages (columns across the printable width, rows down its height, the sheet's page order), each at its paper size with the sheet's margins; `@page` size set. Built by `preparePrint()` and on `beforeprint`. |
| `src/editor/rowheights.js` | The row heights Excel works out itself (tallest font, wrapped lines) and each row's shared baseline. |
| `src/editor/selection.js` | The selection (active cell and range, and with Ctrl the other ranges) and its overlay, with the fill handle and the outline of a fill being dragged; view state only. |
| `src/editor/peers.js` | Others' cursors in a co-edited workbook drawn in the grid's overlay, under this editor's selection: the active cell outlined in the person's colour, their ranges shaded, their name on a flag. |
| `src/editor/navigate.js` | Moving as Excel does: a step (over merged ranges), Ctrl+arrow to the edges of blocks, Ctrl+End. |
| `src/editor/findbar.js` | The Find and Replace bar under the ribbon (Ctrl+F, Ctrl+H), from open-docx-editor's, with Excel's options. |
| `src/editor/ribbonkit.js` | The ribbon's building blocks: buttons, toggles, split buttons, groups, the one drop-down menu, select fields, line icons. From open-docx-editor. |
| `src/editor/ribbon.js` | The ribbon: tabs and panels, groups that don't fit behind More, the one-line simplified ribbon, F6 and Esc, `features`. From open-docx-editor's toolbar.js, with each tab in a file of its own. |
| `src/editor/hometab.js` | The Home tab: Undo, Font, Alignment, Number, Cells, Editing; its controls set from the active cell. |
| `src/editor/viewtab.js` | The View tab: Zoom (Zoom..., 100%, Zoom to Selection) and Window (Freeze Panes, Top Row, First Column, Unfreeze). |
| `src/editor/commands-sheets.js`, `commands-data.js`, `commands-parts.js` | The rest of the named commands, as factories given `commands.js`'s shared state and helpers as one object: the sheets' (show, insert, rename, move, hide, colour, delete, through sheetOp); data (copy, cut, paste, fill, sort, AutoSum, the Name Box, the filter); and parts beyond the cells (PivotTables, notes and rules, another editor's contents, the agent). |
| `src/editor/commands.js` | Every action as a named command (select, move, enter, clear, undo, redo, showSheet, nextSheet, insertSheet, renameSheet, moveSheet, hideSheet, tabColor, deleteSheet, and the formatting ones: format, toggleFont, stepFont, toggleWrap, stepIndent, align, borders, stepDecimal, clearFormats), with the workbook's undo history of cell contents and formats (values and formulas; undoing a change on another sheet shows it; cleared by a rename or delete) and each sheet's selection and scroll position, both by sheetId. |
| `src/editor/celleditor.js` | Editing in a cell: a text box in the cell's font; enter and edit modes, Enter, Tab, Esc, Alt+Enter. |
| `src/editor/autocorrect.js`, `autocorrect-list.js` | AutoCorrect as a cell or the formula bar is typed in (Excel's options and entries of its list), and the spelling underlines. |
| `src/editor/formulabar.js` | The active cell's name and contents; Enter there enters into the cell. |
| `src/editor/interact.js` | Keys and mouse on the grid wired to the commands, the editor and the bar; screen-reader announcements (`live.js`). |
| `src/editor/pointing.js` | Pointing at cells while a formula is typed (Excel's Point mode): arrows, clicks and drags put references in where one can go, F4 cycles $, the formula's ranges outlined in colours, references on another sheet with its name. |
| `src/editor/tabs.js` | The sheet tabs: an ARIA tablist under the grid, sheets whose tabs aren't hidden, in their tab colours (text black or white, whichever reads better); chart sheets' tabs there but not shown yet; the New sheet button, renaming in place, dragging a tab to move it. |
| `src/editor/sheetmenu.js` | A menu of buttons (arrows, Home, End, Esc, a colour submenu) and the tab's entries: Insert, Delete, Rename, Move Left/Right, Tab Color (theme and standard colours), Hide, Unhide. |
| `src/editor/dialogs.js` | Modal message boxes (Excel's refusals, the question before a delete), the Unhide list and boxes of choices (Excel's Insert, Delete and Paste Special boxes), as `<dialog>` elements. |
| `src/editor/specialbox.js` | Excel's Go To Special box. |
| `src/editor/pivotbox.js` | Insert > PivotTable's boxes (the source and where it goes; the fields), and the act that runs them. |
| `src/editor/inserttab.js` | The Insert tab: Tables > PivotTable. |
| `src/editor/formulastab.js`, `src/editor/namesbox.js` | The Formulas tab: Defined Names, Excel's Name Manager and its New / Edit Name box. |
| `src/calc/catalog.js`, `src/editor/functionbox.js`, `src/editor/fntip.js` | The Function Library: each function's category and syntax, Excel's Insert Function box, and the syntax screen tip under a cell being edited. |
| `src/editor/commands-view.js` | The view commands: Freeze Panes, Show Formulas, the zoom, Zoom to Selection, View > Show (Gridlines, Formula Bar, Headings) and Page Layout > Sheet Options' Print boxes; kept in the file but not in the undo history. |
| `src/editor/pagelayouttab.js`, `src/ooxml/printopts.js` | The Page Layout tab (Sheet Options), and its Print boxes written as Excel writes them (`printOptions`). |
| `src/editor/commands-pivot.js` | The newPivot command: a PivotTable made on a new sheet or at a cell, styled and fitted, as one step. |
| `src/editor/axis.js` | Sizes and offsets of rows or columns; binary search from a scroll position to an index. |
| `src/editor/celltext.js` | The one place cell text is sized and placed: font size, line height, baseline shifts for top, middle and wrapped text, on screen and in print (custom properties). |
| `src/editor/fonts.js` | Font stacks with metric-compatible stand-ins; `cellTextMetrics`, the line heights and shifts that put the browser's baselines where Excel's are. |
| `src/editor/decor.js` | Fill rectangles and border edges for the cells in view (and on each printed page), with Excel's printed geometry and dash patterns. |
| `src/editor/cellstyles.js` | Each cell format as a CSS class (`.oxe-s3`) in a `<style>` scoped to one editor; the Normal font on the root. |
| `src/editor/editor.css` | Theme variables (light, dark, host-overridable) and the grid's layout. Geometry through custom properties; no `!important`. |

## Tools and tests

| File | What it does |
|---|---|
| `test/fixtures/make-fixtures.mjs` | Generates the test workbooks (deterministic zips). CI checks the committed files match. |
| `test/*.test.js` | Unit tests in Node (`npm test`): round trips byte for byte, edits, units, references, geometry, the checker, every formula of formulas.xlsx against Excel's results (`test/formulas-excel.json`). |
| `test/e2e.mjs`, `test/editing.mjs` (with `editing-data.mjs`, `editing-more.mjs` and their shared `editing-kit.mjs`), `test/a11y.mjs` | Browser suites (`npm run test:e2e`): each fixture opened, scrolled, saved; editing by keys and mouse, undo, read-only, the saved file read back; axe in light and dark at 1280 and 320px. All assert standards mode. |
| `tools/check-ooxml.mjs` | Structural check of .xlsx files (what makes Excel offer a repair). |
| `tools/excel-compare/` | Comparison with Excel's own PDFs and values: `excel-print.ps1` (Excel through COM, read-only), `excel-values.ps1` (Excel's results of a workbook's formulas, COM, read-only), `print-ours.mjs` (the editor in Edge), `cells.cjs` (cell text and positions), `shapes.cjs` (filled rectangles and lines), `compare.py`, `compare-shapes.py`, `diff.py`, `run.ps1`. See its README. |
| `tools/scan-bytes.mjs` | Control characters and byte order marks in source text (part of `npm run lint`). |

## How a file becomes a grid

1. `OpcPackage.fromBytes` unzips; the workbook is the target of the package's `officeDocument` relationship.
2. `Workbook` reads `xl/workbook.xml` (sheets, `date1904`) and the shared strings.
3. `wb.sheet(i)` parses that sheet's part once and keeps the XML document with the model.
4. `createGrid` turns the model into row and column axes (`sheetGeometry`), then draws the window the scroll
   position shows; scrolling redraws on the next animation frame when the window changes.
5. An edit changes the cell in the sheet's XML and its model, then `Calc` works out again the formulas it reaches.
6. `save()` serialises only the sheets that changed; every other part goes out as it came in (calcChain.xml apart, which
   is dropped when formulas change).
