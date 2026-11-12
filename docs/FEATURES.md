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
| Gridlines on or off | Done | From the sheet's view settings; switched from View > Show or Page Layout > Sheet Options and saved as Excel saves them (docs/VIEW.md). |
| Selecting: keys, mouse, Shift, Ctrl+arrow, Ctrl+End | Done | |
| Selecting whole rows and columns | Done | Headings, Shift+Space, Ctrl+Space, Ctrl+A, the corner. |
| Selecting several ranges (Ctrl+click) | Partly | Ctrl+click or drag adds a range; formatting, Clear and Data Validation go to each; Copy and Cut refuse, as Excel does when they don't line up (docs/GOTOSPECIAL.md). Other commands act on the last range. |
| Typing, F2, Enter/Tab/Esc, Alt+Enter, Delete | Done | |
| Entries read as Excel reads them | Done | Dates, times, currency, thousands, percentages, exponents, fractions, with the format Excel gives the cell (docs/ENTRY.md). United States settings only. |
| Undo and redo | Done | Cell contents and formats across sheets, and sheets inserted, renamed, moved, hidden, coloured and deleted (each one step, the sheet shown as it was). Excel can't undo deleting a sheet (it warns that the sheet goes for good); here Ctrl+Z brings it back, a difference kept on purpose. |
| Formula bar and name box | Partly | The Name Box shows the active cell; a cell, range, whole columns or rows, another sheet's, a defined name, a table's name (its data rows) or a table reference ("Fruit[Qty]") typed there and Enter goes to it (src/ooxml/goto.js, Excel's documented behaviour: COM doesn't reach the Name Box). A new name typed there names the selection (docs/NAMES.md); showing a selected name missing. |
| Formulas typed, Point mode, F4 | Done | docs/CALC.md. |
| Recalculation | Done | 1914 of 1914 test formulas give Excel's results; 500 of Excel's 519 functions (the 19 missing are listed under Functions below); table references (Fruit[Qty], [#Totals], [@Qty]); 3D references (Jan:Apr!A1), following sheets renamed, deleted and moved as Excel's do, intersection (A1:C3 B2:D4) and the range operator (A1:INDEX(A:A,3)). |
| Sheet tabs: show, insert, rename, move, colour, hide, unhide, delete | Done | docs/SHEETS.md. |
| Chart sheets | Kept | Tab shown, sheet not opened. |
| Copy, cut, paste (and paste from Excel or the web) | Done | Ctrl+C, X, V and the Clipboard group; text from other programs read as if typed (docs/CLIPBOARD.md). Their formats are not read. |
| Fill handle, Ctrl+D, Ctrl+R | Done | Series, lists, dates and formulas as Excel fills them; Ctrl-drag the other way (docs/FILL.md). Auto Fill Options and double-click not yet. |
| Right-click menu on cells | Done | Cut, Copy, Paste, Paste Special..., Insert..., Delete..., Clear Contents, Format Cells..., and on a PivotTable, a note or comment, or a hyperlink their own entries (docs/CLIPBOARD.md, docs/LINKS.md). From the keyboard, Shift+F10 or the menu key opens it beside the active cell, as in Excel. |
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
| | Format Painter | Done | Copies the selection's formats and gives them to the next cells selected (click or drag, or Enter), as Paste Special > Formats does; pressed twice, it stays on until pressed again or Esc. |
| Undo (`undo`) | Undo, Redo | Done | |
| Font (`font`) | Font, size, grow and shrink | Done | Grow and Shrink step through Excel's sizes. The font list is a fixed one with the theme's fonts first. |
| | Bold, italic, underline, strikethrough | Done | Ctrl+B, I, U, 5 (and 2, 3, 4). Double underline is shown, not offered. |
| | Borders | Done | Excel's 13 presets; any line style and colour, inside lines and diagonals in Format Cells. Draw Borders missing. |
| | Fill colour, font colour | Done | Excel's palette: the theme's colours and tints, the standard ones. More Colors missing. |
| | Pattern and gradient fills | Done | Format Cells > Fill: pattern style and colour, two-colour gradients (Fill Effects), written as Excel writes them; gradients drawn, patterns drawn as their blend of colours. |
| Alignment (`alignment`) | Horizontal and vertical | Done | Left, center, right; top, middle, bottom on the ribbon; justify, distributed, fill and center across selection in Format Cells. |
| | Wrap text | Done | |
| | Indent | Done | |
| | Orientation (rotated text) | Done | Format Cells > Alignment: -90 to 90 degrees or stacked, as Excel stores them; drawn turned inside the cell, rows grown as Excel grows them (stacked exactly, angles within 2pt). |
| | Merge & Center, Merge Across, Merge Cells, Unmerge | Done | Excel's question before values are discarded. |
| | Shrink to fit | Done | Format Cells > Alignment; the text drawn small enough to fit (Excel's own steps of size not measured). |
| Number (`number`) | Number format list, currency, %, comma, decimals | Done | Every format code Excel has is shown; the list offers Excel's eleven. $, % and comma apply Excel's cell styles. Format Cells (Ctrl+1, also on the cell menu and Cells > Format) has all six tabs: Number, Alignment, Font, Border, Fill, Protection. |
| Styles (`styles`) | Conditional Formatting | Done | The menu: Highlight Cells Rules, Top/Bottom Rules, Data Bars, Icon Sets, two Color Scales, New Rule (six types), Clear Rules, Manage Rules (delete, move, Stop If True, applies to), all written as Excel writes them (14 of 14 menu rules, 33 of 33 bars and icon sets, 5 of 5 manager changes; docs/CONDFORMAT.md). Drawn as Excel draws it: values, formulas, text, top and bottom, average, duplicates, blanks, errors, colour scales; data bars and icons by Excel's documented rules. Not yet: Edit Rule, date rules. |
| | Format as Table | Kept | |
| | Cell Styles | Kept | Cells show their style's formatting. |
| Cells (`cells`) | Insert and delete rows and columns | Partly | Formulas on every sheet, names, chart series, merges, widths and the sheet's ranges follow (docs/ROWCOL.md). Notes, pictures, charts and shapes move as Excel moves them (20 of 20 steps), tables widen, narrow, move or go (35 of 35), and pivot tables move, go or refuse as Excel does, their sources following (28 of 28; docs/ROWCOL.md). Refused for now on a sheet with slicers or timelines. |
| | Insert and delete cells (shifting the rest) | Done | Excel's Insert and Delete boxes, Insert Cells... and Delete Cells...; formulas, names, formats and merges as Excel changes them (docs/ROWCOL.md). |
| | Insert and delete sheets | Done | Cells > Insert, Delete; also on the tab menu. |
| | Row height, column width (and AutoFit) | Done | Dragging a heading's edge, double-click and Format's AutoFit, the Row Height and Column Width boxes; stored as Excel stores them. AutoFit's width is within a few pixels of Excel's (docs/SIZES.md). |
| | Hide and unhide rows, columns | Done | |
| | Rename, hide and unhide sheet | Done | Cells > Format; moving, tab colour and Move or Copy (a copy as Excel copies it, docs/DATATOOLS.md) on the tab menu. |
| Editing (`editing`) | AutoSum | Done | Sum, Average, Count Numbers, Max, Min and Alt+= as Excel places and fills them (docs/AUTOSUM.md). More Functions... missing. |
| | Fill (down, right, series) | Partly | Down, Right, Up, Left (docs/FILL.md); Series... and Flash Fill (Ctrl+E) as Excel (docs/DATATOOLS.md). Justify, Across Worksheets missing. |
| | Clear (all, formats, contents) | Done | Clear All, Clear Formats, Clear Contents, Clear Comments and Notes, Clear Hyperlinks and Remove Hyperlinks as Excel, notes, comments and hyperlinks taken by Clear All and Clear Contents as Excel takes them (docs/FORMATS.md). |
| | Sort & Filter | Done | Sort A to Z, Sort Z to A and Custom Sort... (Excel's Sort box: levels, headers, case, left to right) as Excel sorts (docs/SORT.md); Filter (Ctrl+Shift+L), its buttons and lists, Clear and Reapply as Excel filters (docs/FILTER.md). Text, Number and Date Filters, Top 10, averages and Filter by Color as Excel (31 of 31). |
| | Find & Select | Partly | Find and Replace; Go To (F5 or Ctrl+G: the defined names and tables listed, or a reference typed; src/editor/dialogs.js goToBox). Go To Special (Excel's box and Find & Select's shortcuts; every choice but Objects; docs/GOTOSPECIAL.md). |

## Insert (`insert`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Tables (`tables`) | Table, PivotTable | Partly | Pivot sources follow a sheet rename. A PivotTable refreshed from its source as Excel refreshes it (Alt+F5, Data > Refresh All; docs/PIVOT.md): 8 of 8 cases the same cells and parts as Excel's refresh, and Excel opens the result. Insert > PivotTable makes one as Excel does (8 of 8 cases Excel's own parts and cells; docs/PIVOT.md), its fields chosen in a box; Show Field List moves them afterwards (10 of 10 moves Excel's own parts and cells). Filters, grouping and calculated fields missing. |
| Illustrations (`illustrations`) | Pictures, shapes, icons | Kept | Not drawn on the grid. |
| Charts (`charts`) | Charts | Partly | Insert > Charts: 14 kinds written as Excel writes them (14 of 14, and read back by Excel), the series taken as Excel takes them (10 of 10). Picked, moved, sized, deleted; Edit Chart: kind, title, legend, data labels (15 of 15 edits as Excel writes them). Column, bar, line, area, pie, doughnut and scatter drawn with Excel's automatic axis (34 of 34) and live data; figures with tables for screen readers (docs/CHARTS.md). Not yet: the data range, axis titles, secondary axes, combo and the 2016 kinds. |
| Sparklines (`sparklines`) | Line, column, win/loss | Kept | |
| Links (`links`) | Hyperlinks | Partly | Followed as Excel follows them: a click opens the address or goes to the place in the workbook (a held click selects), Excel's ScreenTip and hand, Open Hyperlink and Remove Hyperlink on the cell's menu; written as Excel writes them, with its Hyperlink style; they follow a sheet rename (docs/LINKS.md). Insert > Link (Ctrl+K) and Edit Hyperlink in Excel's box: a web address, a place in the workbook, an e-mail address, the text to display and a ScreenTip. A HYPERLINK formula's cell is followed too, and gets the Hyperlink style when typed, as in Excel. Not yet: Create New Document and browsing for a file. |
| Comments (`comments`) | Comments and notes | Partly | Shown: Excel's marks, the note or thread beside the cell on hover, read to screen readers; notes written, edited and deleted as Excel writes them (Shift+F2, the cell's menu, Review > Notes; docs/NOTES.md). Threads written as Excel writes them: New Comment (Ctrl+Alt+M), reply, edit, delete a reply or the thread (docs/NOTES.md). |
| Text (`text`) | Text box, header and footer, symbol | Kept | Headers and footers printed and exported as Excel prints them (docs/PRINT.md); not editable yet. |

## Page Layout (`page-layout`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Themes (`themes`) | Theme colours and fonts | Shown | Excel 365's default theme when the file has none. |
| Page Setup (`page-setup`) | Margins, orientation, paper size, page order | Shown | Used when printing. |
| | Print area, print titles, page breaks | Shown | Used when printing and exporting, as Excel does (docs/PRINT.md); can't be set here yet. |
| Scale to Fit (`scale-to-fit`) | Width, height, scale | Shown | Used when printing and exporting, rows and columns in whole printer dots as Excel prints them (docs/PRINT.md); can't be set here yet. |
| Sheet Options (`sheet-options`) | Gridlines and headings, view and print | Done | Gridlines and Headings, View and Print, saved as Excel saves them (Excel reads them back the same; docs/VIEW.md); Print used when printing and exporting (docs/PRINT.md). |

## Formulas (`formulas`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Function Library (`function-library`) | Insert Function, the function lists | Partly | Insert Function (Shift+F3, fx): search, Most Recently Used, All and Excel's categories, each function's syntax; Recently Used and each category's list on the Formulas tab, More Functions for the rest; the choice typed into the cell, or at the caret of a cell being edited; the syntax screen tip, the argument typed in bold (docs/CALC.md). Excel's descriptions of each function and its Function Arguments box are missing. 500 of Excel's 519 functions work when typed (Functions, below). |
| | Dynamic arrays (spilling), LAMBDA, LET | Partly | Formulas spill as Excel 365's do, #SPILL! when blocked; SEQUENCE, SORT, SORTBY, UNIQUE, FILTER, TRANSPOSE, LET, LAMBDA, `D1#` and `@`; 20 of 20 cases Excel's, and Excel reads the files the same (docs/DYNAMIC.md). TAKE, VSTACK, MAP and the other newer array functions too; a LAMBDA kept in a defined name, called as a function (=Twice(21)); GROUPBY, PIVOTBY (any number of row and column fields, subtotals, relative_to), PERCENTOF, RANDARRAY and functions named as LAMBDAs (=MAP(A1:A3,SQRT)). TRIMRANGE and trim references (A1:.E8), 27 cases as Excel's. |
| Defined Names (`defined-names`) | Name Manager, define, use in formula | Done | Name Manager (Ctrl+F3) and Define Name as Excel's, the Name Box naming the selection; Excel's rules for a name, 34 of 34 verdicts, and its saved order (docs/NAMES.md). Create from Selection and Use in Formula missing. |
| Formula Auditing (`formula-auditing`) | Trace precedents and dependents, show formulas, evaluate | Partly | Trace Precedents and Trace Dependents a level at a time, Excel's arrows (a range boxed, red from an error, another sheet's from its icon), Remove Arrows; Show Formulas (Ctrl+`): what each cell holds, columns twice as wide, saved as Excel saves it (docs/CALC.md). Excel's error indicators (green triangles) for its six default rules, ignoredErrors kept (`errorChecking: false` turns them off); the error button and its menu (Convert to Number, Copy Formula from Above, Lock Cell, Ignore Error written as Excel writes it); the Error Checking box, through the sheet's errors from the active cell; Evaluate Formula a step at a time (Step In and Step Out not yet). the Watch Window (Add Watch, Delete Watch, saved as Excel saves it). Error Checking Options missing; arrows stay after edits until removed. |
| Calculation (`calculation`) | Automatic, manual, Calculate Now | Partly | Calculation Options (Automatic, Automatic Except for Data Tables, Manual), Calculate Now (F9), Calculate Sheet (Shift+F9), Ctrl+Alt+F9; in Manual a change waits and the status bar says Calculate, a formula entered is worked out, saving works it out first; saved as Excel saves it (docs/CALC.md). Inserted rows and moved sheets still work their formulas out in Manual. Circular references as Excel has them: with iterative calculation off left out and announced (one just entered 0); with it on (read from the file, kept, or `setIteration`) gone round to Maximum Iterations or Maximum Change, 41 cases as Excel's to the last bit (docs/CALC.md). Missing: an on-screen switch for it (Excel's File > Options > Formulas). |

## Functions

Excel 365 has 519 worksheet functions: the 520 on Microsoft's alphabetical list ("Excel functions (alphabetical)")
plus BETA.INV, which that list leaves out, less EUROCONVERT (from the Euro Currency Tools add-in) and JIS (Japanese
Excel only), which Excel 365 in English gives #NAME?. Each name was typed into Excel over COM (Formula2 "=NAME()"): a
function it has refuses the empty call, an unknown name gives #NAME?. 500 of the 519 work here (`src/calc`; DATEDIF
among them, which Excel too leaves out of Insert Function). The 19 missing:

| Kind | Missing | Count | Notes |
|---|---|---|---|
| Possible here | FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT | 4 | Excel's AAA exponential smoothing; its exact fitting is to be measured. |
| Need something outside the workbook | CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE | 7 | An OLAP data model or connection. |
| | RTD, WEBSERVICE, STOCKHISTORY, IMAGE, TRANSLATE, DETECTLANGUAGE | 6 | A server, a web address or Microsoft's online services. |
| | CALL, REGISTER.ID | 2 | Code in a Windows DLL; off by default in Excel too. |

A formula using one of these keeps the value Excel saved (it isn't worked out again here, even when what it reads
changes); one typed here shows #NAME? until Excel opens the file. Excel's saved results are kept for all 19.

## Data (`data`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Get & Transform (`get-and-transform`) | Queries and connections | Kept | |
| Queries & Connections (`queries-and-connections`) | Refresh All | Partly | Refreshes the workbook's PivotTables (Ctrl+Alt+F5); Refresh, the one at the active cell (Alt+F5). Data connections aren't refreshed. |
| Sort & Filter (`sort-and-filter`) | Sort, filter, advanced | Done | Sort A to Z, Z to A, Sort..., Filter, Clear and Reapply on the Data tab, as in Home > Editing (docs/SORT.md, docs/FILTER.md). Text, Number and Date Filters, Top 10, averages and Filter by Color (31 of 31 as Excel), and Advanced Filter in place or copied (7 of 7). |
| Data Tools (`data-tools`) | Text to Columns, Remove Duplicates | Done | Text to Columns (delimited), Flash Fill, Remove Duplicates as Excel (24 of 24 with Series; docs/DATATOOLS.md). Fixed width not yet. |
| | Data Validation (lists, rules) | Done | Typed values checked as Excel checks them, its Stop, Warning and Information boxes; a list's dropdown; the input message; Excel's Data Validation box to add, change and clear rules, written as Excel writes them (15 of 15; docs/VALIDATION.md). Circle Invalid Data missing. |
| Outline (`outline`) | Group, ungroup, subtotal | Kept | Grouped rows and columns shown expanded or hidden as saved, no outline buttons. |

## Review (`review`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Proofing (`proofing`) | Spelling | Partly | AutoCorrect as a cell is typed, with Excel's options and entries of its own replacement list (teh to the, (c) to ©, monday to Monday, THe to The); the browser's spelling underlines while text is typed (docs/PROOFING.md). The Spelling box (F7) is missing. |
| Accessibility (`accessibility`) | Check Accessibility | Missing | |
| Comments (`comments`) | New, show, delete comments; notes | Done | Notes: New, Edit, Delete, Show/Hide; threaded comments: New, Reply, Edit, Delete (docs/NOTES.md). Not yet: @mentions, resolving a thread. |
| Protect (`protect`) | Protect sheet, protect workbook | Partly | Protect Sheet and Unprotect Sheet in Excel's boxes, written as Excel writes them (its password hash to the byte; Excel opens the file protected the same way). Locked cells refused in Excel's words; formatting, sizes, inserting and deleting, sorting and filtering refused unless allowed (docs/PROTECT.md). Protect Workbook: its structure, as Excel writes it, sheets' changes refused. Hidden cells (Format Cells > Protection) show no formula while the sheet is protected, as in Excel. Allow Edit Ranges: Excel's box (New, Modify, Delete; a range's password), honoured as Excel does, a password range asking Unlock Range. Notes, charts, conditional formats, validation, hyperlinks and PivotTables refused on a protected sheet as Excel refuses them. Not yet: the selection options enforced. |

## View (`view`)

| Group | Feature | Status | Notes |
|---|---|---|---|
| Workbook Views (`workbook-views`) | Normal, Page Break Preview, Page Layout | Partly | Normal only. |
| Show (`show`) | Gridlines, formula bar, headings | Done | Gridlines, Formula Bar and Headings switched on and off; gridlines and headings saved with the sheet as Excel saves them, the formula bar the editor's own, as Excel's is (docs/VIEW.md). Excel's Ruler only works in Page Layout view, which isn't here. |
| Zoom (`zoom`) | Zoom, 100%, zoom to selection | Done | 10% to 400%, as Excel allows; saved with the sheet. |
| Window (`window`) | Freeze Panes | Done | Freeze Panes, Top Row, First Column, Unfreeze, written as Excel writes them; frozen panes in a file shown frozen (docs/VIEW.md). |
| | Split, new window | Missing | |

## File (`file`)

| Feature | Status | Notes |
|---|---|---|
| Open, Save | Done | Through the host: `open(bytes)`, `save()`. |
| Print | Done | |
| Properties (title, author) | Kept | |
