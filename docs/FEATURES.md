# Features: what Excel does and what this editor does

The editor's features listed against Excel 365's, by its ribbon tabs and groups. The tab and group keys are the names
the `features` option will take (as in open-docx-editor: a tab by name, a group by its label in lower case with
dashes), so this list is also the ribbon's plan. Update a row in the commit that changes it.

| Mark | Meaning |
|---|---|
| Done | Works as in Excel, measured against Excel where there is something to measure. |
| Partly | Some of it works; the note says what is missing. |
| Shown | Read from the file and shown, but there is no way to change it yet. |
| Kept | Not shown, but saved back as it came in (every part we don't change goes out byte for byte, and an edited sheet is changed in place). |
| Missing | Not there. |

Nothing is lost on save: what this editor doesn't understand is Kept, never dropped. The exception is calcChain.xml,
which is dropped when formulas change (Excel rebuilds it).

## Options (planned, as in open-docx-editor)

| Option | Status | Notes |
|---|---|---|
| `readOnly` | Done | `setReadOnly()` at any time. No ribbon, nothing editable; the tab menu shows only what doesn't change the workbook. |
| `toolbar: false` | Done | Hides the ribbon but keeps editing. |
| `features: { tab or group: false }` | Done | Hides a tab or group; `setFeatures()` at any time. A formatting group turned off also turns off its keys (Ctrl+B with `font: false`). |

## Outside the ribbon

| Feature | Status | Notes |
|---|---|---|
| Open .xlsx | Done | Any workbook; sheets parsed on first view. |
| Open .xlsm, .xltx, .xlsb, .xls, .csv | Partly | .xlsm opens, and an edit, a new sheet and a save keep its VBA project, custom ribbon (customUI) and macro-enabled content type byte for byte (test/macros.test.js); macros don't run and can't be edited. .xltx, .xlsb, .xls and .csv missing. |
| Workbooks saved with a password | Done | Opened with Excel's Password box (or `{ password }`), saved encrypted again with it, or with a new one (`setPassword`); Excel 2010 and later's agile encryption; Excel opens what's saved here (docs/ENCRYPTION.md). Excel 2007's older encryption refused. |
| Co-editing | Partly | Over Yjs, with any provider (src/collab.js, docs/COLLAB.md): cell contents live, cell by cell; formats, rows, columns and sheets as the whole workbook, which the others reopen. Others' cursors drawn on the grid in each person's colour with their name (through the provider's awareness); undo is each editor's own and starts again when another's whole-workbook change arrives. |
| Agent interface | Done | Read sheets, names, tables, ranges (values, shown text, formulas, Markdown), find and notes; enter text as typed, write blocks, insert and delete rows and columns, add, rename and delete sheets, make and refresh PivotTables; a review of every change (cells before and after, sheets, rows and columns) with any of it taken back; an MCP server (`open-xlsx-editor-mcp`) for AI assistants; without a browser or on the open editor, each edit one undo step (docs/AGENT.md). 12 of 12 typed entries read the same in Excel. |
| Save .xlsx | Done | Only changed sheets rewritten; opened in Excel with no repair (`tools/check-ooxml.mjs`). |
| Formula results on open when the file has none | Done | Formulas saved without a result (as some programs save them) are worked out on opening, in the editor and the agent, as Excel does; the results saved with the others are kept, and a workbook saved unchanged still comes back byte for byte (`calculateMissing`, docs/CALC.md; `calculateOnOpen: false` turns it off). |
| Virtualised grid | Done | 5,000 rows by 20 columns scroll smoothly. |
| Fonts, fills, borders, alignment, number formats | Done | Shown as measured against Excel's PDFs (docs/LAYOUT.md); changed from the Home tab (below). |
| Rich text in a cell (runs in several fonts) | Partly | The text is shown in the cell's own font; the runs' formatting isn't. |
| Merged cells | Done | Navigation steps over them; merged and unmerged from the Home tab. |
| Hidden rows and columns | Done | Hidden and shown from the headings' menu, Format, Ctrl+9 and Ctrl+0 (docs/SIZES.md). |
| Row heights from fonts and wrapping | Done | As Excel works them out (docs/LAYOUT.md). |
| Gridlines on or off | Shown | From the sheet's view settings. |
| Selecting: keys, mouse, Shift, Ctrl+arrow, Ctrl+End | Done | |
| Selecting whole rows and columns | Done | Headings, Shift+Space, Ctrl+Space, Ctrl+A, the corner. |
| Selecting several ranges (Ctrl+click) | Partly | Ctrl+click or drag adds a range; formatting, Clear and Data Validation go to each; Copy and Cut refuse, as Excel does when they don't line up (docs/GOTOSPECIAL.md). Other commands act on the last range. |
| Typing, F2, Enter/Tab/Esc, Alt+Enter, Delete | Done | |
| Entries read as Excel reads them | Done | Dates, times, currency, thousands, percentages, exponents, fractions, with the format Excel gives the cell (docs/ENTRY.md). United States settings only. |
| Undo and redo | Done | Cell contents and formats across sheets, and sheets inserted, renamed, moved, hidden, coloured and deleted (each one step, the sheet shown as it was). Excel can't undo deleting a sheet (it warns that the sheet goes for good); here Ctrl+Z brings it back, a difference kept on purpose. |
| Formula bar and name box | Partly | The Name Box shows the active cell; a cell, range, whole columns or rows, another sheet's, a defined name, a table's name (its data rows) or a table reference ("Fruit[Qty]") typed there and Enter goes to it (src/ooxml/goto.js, Excel's documented behaviour: COM doesn't reach the Name Box). A new name typed there names the selection (docs/NAMES.md); showing a selected name missing. |
| Formulas typed, Point mode, F4 | Done | docs/CALC.md. |
| Recalculation | Done | 541 of 541 test formulas give Excel's results; 189 functions; table references (Fruit[Qty], [#Totals], [@Qty]). |
| Sheet tabs: show, insert, rename, move, colour, hide, unhide, delete | Done | docs/SHEETS.md. |
| Chart sheets | Kept | Tab shown, sheet not opened. |
| Copy, cut, paste (and paste from Excel or the web) | Done | Ctrl+C, X, V and the Clipboard group; text from other programs read as if typed (docs/CLIPBOARD.md). Their formats are not read. |
| Fill handle, Ctrl+D, Ctrl+R | Done | Series, lists, dates and formulas as Excel fills them; Ctrl-drag the other way (docs/FILL.md). Auto Fill Options and double-click not yet. |
| Right-click menu on cells | Done | Cut, Copy, Paste, Paste Special..., Insert..., Delete..., Clear Contents (docs/CLIPBOARD.md). |
| Formatting whole rows and columns | Partly | A selection past the used range is cut to it; Excel gives the row or column itself the format. |
| Find and Replace (Ctrl+F, Ctrl+H) | Done | Excel's options, Find All as a count, Replace All in one undo step (docs/FIND.md). |
| Status bar (Sum, Average, Count of the selection) | Done | Beside the sheet tabs; Count alone for text or an error; hidden rows and columns left out; the first number's format (`src/ooxml/stats.js`, by Excel's documented rules: COM doesn't reach the status bar). Choosing Min, Max and Numerical Count missing; `statusBar: false` hides it. |
| Zoom | Done | View tab, Ctrl+wheel; saved (docs/VIEW.md). |
| Print | Done | Pages as Excel lays them out, measured against its PDFs; print area, titles, breaks, scale, fit to page, headers and footers as Excel applies them (docs/PRINT.md). |
| PDF export | Done | `exportPdf` (src/pdf.js, docs/PDF.md): the sheet shown or the whole workbook, from the print layout, as real text in the installed fonts; 102,079 of 102,097 cells within 1px of Excel's PDFs. Headers, footers, charts and pictures not drawn, as in print. |
| Accessibility | Done | ARIA grid and tablist, announcements, axe clean in light and dark at 1280 and 320px. |
| Light and dark | Done | The chrome follows the theme; the sheet stays white, as in Excel. |

## Home (`home`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Clipboard (`clipboard`) | Cut, Copy, Paste | Done | As Excel moves formulas, and the formulas that read cut cells. |
| | Paste Special | Partly | Excel's box (all, formulas, values, formats, all except borders, column widths, the number-format pairs; the four operations, Skip blanks, Transpose, Paste Link) and the Paste menu; for cells copied here (docs/CLIPBOARD.md). Not yet: comments, validation, source theme, merging conditional formats. |
| | Format Painter | Missing | |
| Undo (`undo`) | Undo, Redo | Done | |
| Font (`font`) | Font, size, grow and shrink | Done | Grow and Shrink step through Excel's sizes. The font list is a fixed one with the theme's fonts first. |
| | Bold, italic, underline, strikethrough | Done | Ctrl+B, I, U, 5 (and 2, 3, 4). Double underline is shown, not offered. |
| | Borders | Done | Excel's 13 presets. Line colour and style, and Draw Borders, missing. |
| | Fill colour, font colour | Done | Excel's palette: the theme's colours and tints, the standard ones. More Colors missing. |
| | Pattern and gradient fills | Partly | Solid and pattern fills shown; gradients not. |
| Alignment (`alignment`) | Horizontal and vertical | Done | Left, center, right; top, middle, bottom. Justify, distributed, fill and center across selection are shown, not offered. |
| | Wrap text | Done | |
| | Indent | Done | |
| | Orientation (rotated text) | Missing | |
| | Merge & Center, Merge Across, Merge Cells, Unmerge | Done | Excel's question before values are discarded. |
| | Shrink to fit | Missing | |
| Number (`number`) | Number format list, currency, %, comma, decimals | Done | Every format code Excel has is shown; the list offers Excel's eleven. $, % and comma apply Excel's cell styles. The Format Cells dialog missing. |
| Styles (`styles`) | Conditional Formatting | Partly | Highlight Cells Rules, Top/Bottom Rules, two Color Scales and Clear Rules added from the menu as Excel writes them (14 of 14 drawn the same by Excel; docs/CONDFORMAT.md). Drawn as Excel draws it: cell values, formulas, text, top and bottom, average, duplicates, blanks, errors, colour scales (docs/CONDFORMAT.md). Data bars, icon sets and the menu missing; kept. |
| | Format as Table | Kept | |
| | Cell Styles | Kept | Cells show their style's formatting. |
| Cells (`cells`) | Insert and delete rows and columns | Partly | Formulas on every sheet, names, chart series, merges, widths and the sheet's ranges follow (docs/ROWCOL.md). Notes, pictures, charts and shapes move as Excel moves them (20 of 20 steps), tables widen, narrow, move or go (35 of 35), and pivot tables move, go or refuse as Excel does, their sources following (28 of 28; docs/ROWCOL.md). Refused for now on a sheet with slicers or timelines. |
| | Insert and delete cells (shifting the rest) | Done | Excel's Insert and Delete boxes, Insert Cells... and Delete Cells...; formulas, names, formats and merges as Excel changes them (docs/ROWCOL.md). |
| | Insert and delete sheets | Done | Cells > Insert, Delete; also on the tab menu. |
| | Row height, column width (and AutoFit) | Done | Dragging a heading's edge, double-click and Format's AutoFit, the Row Height and Column Width boxes; stored as Excel stores them. AutoFit's width is within a few pixels of Excel's (docs/SIZES.md). |
| | Hide and unhide rows, columns | Done | |
| | Rename, hide and unhide sheet | Done | Cells > Format; moving and tab colour on the tab menu. |
| Editing (`editing`) | AutoSum | Done | Sum, Average, Count Numbers, Max, Min and Alt+= as Excel places and fills them (docs/AUTOSUM.md). More Functions... missing. |
| | Fill (down, right, series) | Partly | Down, Right, Up, Left (docs/FILL.md). Series..., Justify, Across Worksheets missing. |
| | Clear (all, formats, contents) | Partly | Clear All, Clear Formats, Clear Contents as Excel (docs/FORMATS.md). Clear Comments and Notes, Clear Hyperlinks missing. |
| | Sort & Filter | Done | Sort A to Z, Sort Z to A and Custom Sort... (Excel's Sort box: levels, headers, case, left to right) as Excel sorts (docs/SORT.md); Filter (Ctrl+Shift+L), its buttons and lists, Clear and Reapply as Excel filters (docs/FILTER.md). Text, number, date and colour filters missing (kept). |
| | Find & Select | Partly | Find and Replace; Go To (F5 or Ctrl+G: the defined names and tables listed, or a reference typed; src/editor/dialogs.js goToBox). Go To Special (Excel's box and Find & Select's shortcuts; every choice but Objects; docs/GOTOSPECIAL.md). |

## Insert (`insert`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Tables (`tables`) | Table, PivotTable | Partly | Pivot sources follow a sheet rename. A PivotTable refreshed from its source as Excel refreshes it (Alt+F5, Data > Refresh All; docs/PIVOT.md): 8 of 8 cases the same cells and parts as Excel's refresh, and Excel opens the result. Insert > PivotTable makes one as Excel does (8 of 8 cases Excel's own parts and cells; docs/PIVOT.md), its fields chosen in a box. Moving fields afterwards, filters, grouping and calculated fields, missing. |
| Illustrations (`illustrations`) | Pictures, shapes, icons | Kept | Not drawn on the grid. |
| Charts (`charts`) | Charts | Partly | Column, bar, line, area, pie and doughnut drawn, with Excel's automatic axis (34 of 34) and live data; given to screen readers as figures with tables (docs/CHARTS.md). Editing and other kinds missing; kept. Series follow a sheet rename. |
| Sparklines (`sparklines`) | Line, column, win/loss | Kept | |
| Links (`links`) | Hyperlinks | Kept | Not clickable yet; follow a sheet rename. |
| Comments (`comments`) | Comments and notes | Partly | Shown: Excel's marks, the note or thread beside the cell on hover, read to screen readers; notes written, edited and deleted as Excel writes them (Shift+F2, the cell's menu, Review > Notes; docs/NOTES.md). Threads shown and deleted, not edited. |
| Text (`text`) | Text box, header and footer, symbol | Kept | Headers and footers printed and exported as Excel prints them (docs/PRINT.md); not editable yet. |

## Page Layout (`layout`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Themes (`themes`) | Theme colours and fonts | Shown | Excel 365's default theme when the file has none. |
| Page Setup (`page-setup`) | Margins, orientation, paper size, page order | Shown | Used when printing. |
| | Print area, print titles, page breaks | Shown | Used when printing and exporting, as Excel does (docs/PRINT.md); can't be set here yet. |
| Scale to Fit (`scale-to-fit`) | Width, height, scale | Shown | Used when printing and exporting, rows and columns in whole printer dots as Excel prints them (docs/PRINT.md); can't be set here yet. |
| Sheet Options (`sheet-options`) | Gridlines and headings, view and print | Partly | Gridlines on screen follow the file; nothing can be changed. |

## Formulas (`formulas`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Function Library (`function-library`) | Insert Function, the function lists | Missing | 189 functions work when typed (docs/CALC.md). |
| | Dynamic arrays (spilling), LAMBDA, LET | Partly | Formulas spill as Excel 365's do, #SPILL! when blocked; SEQUENCE, SORT, SORTBY, UNIQUE, FILTER, TRANSPOSE, LET, LAMBDA, `D1#` and `@`; 20 of 20 cases Excel's, and Excel reads the files the same (docs/DYNAMIC.md). The newer array functions (TAKE, VSTACK, MAP...) missing. |
| Defined Names (`defined-names`) | Name Manager, define, use in formula | Done | Name Manager (Ctrl+F3) and Define Name as Excel's, the Name Box naming the selection; Excel's rules for a name, 34 of 34 verdicts, and its saved order (docs/NAMES.md). Create from Selection and Use in Formula missing. |
| Formula Auditing (`formula-auditing`) | Trace precedents and dependents, show formulas, evaluate | Missing | |
| Calculation (`calculation`) | Automatic, manual, Calculate Now | Partly | Always automatic. |

## Data (`data`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Get & Transform (`get-and-transform`) | Queries and connections | Kept | |
| Queries & Connections (`queries-and-connections`) | Refresh All | Partly | Refreshes the workbook's PivotTables (Ctrl+Alt+F5); Refresh, the one at the active cell (Alt+F5). Data connections aren't refreshed. |
| Sort & Filter (`sort-and-filter`) | Sort, filter, advanced | Partly | Sort A to Z, Z to A, Sort..., Filter, Clear and Reapply on the Data tab, as in Home > Editing (docs/SORT.md, docs/FILTER.md). Advanced Filter missing. |
| Data Tools (`data-tools`) | Text to Columns, Remove Duplicates | Missing | |
| | Data Validation (lists, rules) | Done | Typed values checked as Excel checks them, its Stop, Warning and Information boxes; a list's dropdown; the input message; Excel's Data Validation box to add, change and clear rules, written as Excel writes them (15 of 15; docs/VALIDATION.md). Circle Invalid Data missing. |
| Outline (`outline`) | Group, ungroup, subtotal | Kept | Grouped rows and columns shown expanded or hidden as saved, no outline buttons. |

## Review (`review`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Proofing (`proofing`) | Spelling | Missing | |
| Accessibility (`accessibility`) | Check Accessibility | Missing | |
| Comments (`comments`) | New, show, delete comments; notes | Partly | Notes: New, Edit, Delete, Show/Hide (docs/NOTES.md); threaded comments shown and deleted, not written. |
| Protect (`protect`) | Protect sheet, protect workbook | Kept | Not enforced: a protected sheet can be edited. |

## View (`view`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Workbook Views (`workbook-views`) | Normal, Page Break Preview, Page Layout | Partly | Normal only. |
| Show (`show`) | Gridlines, formula bar, headings | Partly | Gridlines follow the file; none can be switched. |
| Zoom (`zoom`) | Zoom, 100%, zoom to selection | Done | 10% to 400%, as Excel allows; saved with the sheet. |
| Window (`window`) | Freeze Panes | Done | Freeze Panes, Top Row, First Column, Unfreeze, written as Excel writes them; frozen panes in a file shown frozen (docs/VIEW.md). |
| | Split, new window | Missing | |

## File (`file`)

| Feature | Status | Notes |
|---|---|---|
| Open, Save | Done | Through the host: `open(bytes)`, `save()`. |
| Print | Done | |
| Properties (title, author) | Kept | |
