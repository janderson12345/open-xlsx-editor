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
| `src/ooxml/stats.js` | The status bar's Average, Count and Sum of the selection, by Excel's rules. |
| `src/ooxml/fillops.js` | Fill Down, Right, Up and Left, and the fill handle (AutoFill) as Excel does them: series, lists, dates, formulas moved, formats repeated (docs/FILL.md). |
| `src/ooxml/charts.js` | A sheet's charts read from its drawing and chart parts, their data from the cells, and Excel's automatic value axis; drawn by `src/editor/chartview.js` (docs/CHARTS.md). |
| `src/ooxml/tables.js` | The workbook's tables, and what a table reference (Fruit[Qty], [#Totals], [@Qty]) stands for. |
| `src/ooxml/goto.js` | What the Name Box goes to: a cell, range, whole lines, another sheet's, or a defined name; Excel's words otherwise. |
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
| `src/calc/tokens.js` | Formulas as tokens; moving a formula (`shiftFormula`), a sheet's references renamed or made #REF! (`renameSheetRefs`, `sheetPrefix`), references following cells cut and pasted (`moveRefs`) and rows or columns inserted or deleted (`shiftRefs`), the typed form to the file's (`normalizeFormula`, `_xlfn.` prefixes) and back (`displayFormula`). |
| `src/calc/parse.js` | Tokens to a tree by Excel's precedence. |
| `src/calc/values.js` | Values (numbers, text, booleans, errors, blanks, `Area`, `Arr`) and Excel's coercions, comparison, numbers as text, the near-zero rule. |
| `src/calc/evaluate.js` | Operators, references, names and function calls; implicit intersection outside array formulas, element by element inside. |
| `src/calc/functions.js` | The worksheet functions, as Excel defines them. |
| `src/calc/engine.js` | `Calc`: what each formula reads, what a change reaches, the order to work them out in, results written back. |
| `src/ooxml/strings.js` | Shared string table and inline strings; phonetic runs skipped. |
| `src/ooxml/styles.js` | Fonts, fills, borders and cell formats (cellXfs) from styles.xml; `fontFamily` resolves scheme fonts to the theme's. |
| `src/ooxml/theme.js` | The theme's fonts and colours (Excel 365's default theme without a theme part); `resolveColor`, the one colour resolver (theme swap, integer-HLS tints, indexed palette). |
| `src/ooxml/numfmt.js` | What a cell shows (`cellDisplay`): the General format, and format codes through format.js. |
| `src/ooxml/format.js` | Excel's number format codes: sections, digits, scaling, scientific, fractions, dates and times, text; built-in formats; serial dates (1900 and 1904). |
| `src/editor/grid.js` | The virtualised grid: only rows and columns in view (plus overscan) in the DOM; sticky headings; text running on over empty cells; merged ranges; gridlines; frozen panes (their own layers, sticking in place) and zoom. An ARIA grid (rows, row and column headers, gridcells with indexes). |
| `src/editor/print.js` | The printed sheet: used range split into pages (columns across the printable width, rows down its height, the sheet's page order), each at its paper size with the sheet's margins; `@page` size set. Built by `preparePrint()` and on `beforeprint`. |
| `src/editor/rowheights.js` | The row heights Excel works out itself (tallest font, wrapped lines) and each row's shared baseline. |
| `src/editor/selection.js` | The selection (active cell and range) and its overlay, with the fill handle and the outline of a fill being dragged; view state only. |
| `src/editor/navigate.js` | Moving as Excel does: a step (over merged ranges), Ctrl+arrow to the edges of blocks, Ctrl+End. |
| `src/editor/findbar.js` | The Find and Replace bar under the ribbon (Ctrl+F, Ctrl+H), from open-docx-editor's, with Excel's options. |
| `src/editor/ribbonkit.js` | The ribbon's building blocks: buttons, toggles, split buttons, groups, the one drop-down menu, select fields, line icons. From open-docx-editor. |
| `src/editor/ribbon.js` | The ribbon: tabs and panels, groups that don't fit behind More, the one-line simplified ribbon, F6 and Esc, `features`. From open-docx-editor's toolbar.js, with each tab in a file of its own. |
| `src/editor/hometab.js` | The Home tab: Undo, Font, Alignment, Number, Cells, Editing; its controls set from the active cell. |
| `src/editor/viewtab.js` | The View tab: Zoom (Zoom..., 100%, Zoom to Selection) and Window (Freeze Panes, Top Row, First Column, Unfreeze). |
| `src/editor/commands.js` | Every action as a named command (select, move, enter, clear, undo, redo, showSheet, nextSheet, insertSheet, renameSheet, moveSheet, hideSheet, tabColor, deleteSheet, and the formatting ones: format, toggleFont, stepFont, toggleWrap, stepIndent, align, borders, stepDecimal, clearFormats), with the workbook's undo history of cell contents and formats (values and formulas; undoing a change on another sheet shows it; cleared by a rename or delete) and each sheet's selection and scroll position, both by sheetId. |
| `src/editor/celleditor.js` | Editing in a cell: a text box in the cell's font; enter and edit modes, Enter, Tab, Esc, Alt+Enter. |
| `src/editor/formulabar.js` | The active cell's name and contents; Enter there enters into the cell. |
| `src/editor/interact.js` | Keys and mouse on the grid wired to the commands, the editor and the bar; screen-reader announcements (`live.js`). |
| `src/editor/pointing.js` | Pointing at cells while a formula is typed (Excel's Point mode): arrows, clicks and drags put references in where one can go, F4 cycles $, the formula's ranges outlined in colours, references on another sheet with its name. |
| `src/editor/tabs.js` | The sheet tabs: an ARIA tablist under the grid, sheets whose tabs aren't hidden, in their tab colours (text black or white, whichever reads better); chart sheets' tabs there but not shown yet; the New sheet button, renaming in place, dragging a tab to move it. |
| `src/editor/sheetmenu.js` | A menu of buttons (arrows, Home, End, Esc, a colour submenu) and the tab's entries: Insert, Delete, Rename, Move Left/Right, Tab Color (theme and standard colours), Hide, Unhide. |
| `src/editor/dialogs.js` | Modal message boxes (Excel's refusals, the question before a delete), the Unhide list and boxes of choices (Excel's Insert, Delete and Paste Special boxes), as `<dialog>` elements. |
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
| `test/e2e.mjs`, `test/editing.mjs`, `test/a11y.mjs` | Browser suites (`npm run test:e2e`): each fixture opened, scrolled, saved; editing by keys and mouse, undo, read-only, the saved file read back; axe in light and dark at 1280 and 320px. All assert standards mode. |
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
