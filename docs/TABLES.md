# Tables

How tables (Excel's ListObjects) are read (`src/ooxml/tables.js`: the range, header and total rows, columns, each
column's total, the style and Table Style Options, the filter buttons) and drawn in their table style
(`src/ooxml/tablestyle.js` works out each cell's look; `src/editor/decor.js` and `cells.js` draw it on screen and in the
printed pages, and `src/pdf.js` exports what print draws). Tables follow inserted and deleted rows and columns
(`src/ooxml/anchors.js`, docs/ROWCOL.md), work in formulas (structured references, `tables.js tableArea`) and keep
their saved sort (`src/ooxml/sortstate.js`, docs/SORT.md).

## Table styles

Excel has 60 built-in table styles (TableStyleLight1 to 21, Medium1 to 28, Dark1 to 11) that a workbook names
(`<tableStyleInfo name="TableStyleMedium2">`) but doesn't hold. They were measured by having Excel duplicate each one
(`TableStyle.Duplicate` over COM, `tools/excel-compare/excel-tablestyles.ps1`): it saves the copy as a style of the
workbook's own, each element a differential format (dxf) in `styles.xml`. Those dxfs are
`test/tablestyles-excel.json`; `tools/excel-compare/tablestyles-data.mjs` writes them into `src/engine/tablestyles.js`
(180 distinct dxfs). Every built-in stripe is one row or column (`TableStyleElement.StripeSize`). A workbook's own
styles (`<tableStyles><tableStyle>`, each `tableStyleElement` a `dxfId` and a stripe `size`) are read from its styles
part and win over a built-in style of the same name.

How a style is laid on a table, measured against what Excel shows (`Range.DisplayFormat` of every cell, over COM:
`tools/excel-compare/excel-tablelook.ps1`):

- The elements are laid on in this order, each over the ones before: the whole table, the column stripes, the row
  stripes, the last column, the first column, the header row, the total row, then the first and last header cell and
  the first and last total cell. The stripes cover the data rows only and start from the first data row and column;
  the first and last column run from the header to the total row. Each element counts only when its option is on
  (Header Row, Total Row, Banded Rows, Banded Columns, First Column, Last Column).
- An element's `left`, `right`, `top` and `bottom` border sides are its region's outline, its `vertical` and
  `horizontal` sides the lines inside it; a stripe's region is each band. One line per edge: the last element to set
  it.
- The cell's own format is over the table's: a fill (any pattern, white too), a border side, and a font colour, bold or
  italic that differs from its cell style's. An rgb black on a white header shows black; bold turned off on a header
  cell (the Normal style's) stays bold.

Checked cell by cell (`test/tablestyle.test.js`): all 60 styles with every option on and a total row
(`test/tablelook-all-excel.json`), all 60 with only the header and total rows (`tablelook-plain-excel.json`), and
Medium2, Light1, Light8, Light9, Light15 and Dark1 with the options in turn, no style, and cells of their own format
(`tablelook-excel.json`): fill, font colour, bold, italic and the four edges the same as Excel's in every cell.

Tints: a file's tint is Excel's to four places (0.79998168889431442 is lighter 80%). Rounded so, `applyTint`
(src/ooxml/theme.js) gives 220 of 220 colours as Excel does over COM (22 colours, ten tints each:
`tools/excel-compare/excel-tints.ps1`, `test/tints-excel.json`); without the rounding, 0.79998 came out a step darker
in some table colours (TableStyleLight7's stripe D9F1CF for Excel's DAF2D0).

## Making and changing tables: what Excel writes

`src/ooxml/tableedit.js` (and `tableconvert.js` for Convert to Range) write what Excel saves; measured over COM
(`tools/excel-compare/excel-tables.ps1` on `table-cases.json`: 33 cases on made-up lists, each saved, the table parts,
cells, relationships and styles kept in `test/tablemake-excel.json`). `test/tableedit.test.js` replays each case and
compares: 33 of 33 the same table parts (uids aside), cells and formulas, and after Convert to Range the same formats.
Excel reads back the editor's file of each case (`tools/excel-compare/excel-tables-check.ps1` on `test/out/tablemake`,
`test/tablemake-readback-excel.json`): 33 of 33 open, every table's name, range, style, options, filter, columns and
totals as the editor wrote them, and a cell given Good by the editor in Good.

- **Insert > Table, Ctrl+T, Format as Table** (`createTable`, ListObjects.Add): a part `xl/tables/tableN.xml` with
  its relationship from the sheet, a content type override and the sheet's `<tableParts>`. `<table id name
  displayName ref totalsRowShown="0">`, an `<autoFilter>` over the whole range, a `<tableColumn>` per column and
  `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1"
  showColumnStripes="0"/>` (the workbook's `defaultTableStyle`). The name is the first free Table1, Table2...
- With headers, the first row's texts name the columns, written back as text: a number as its digits (2024), a formula
  as its result; a blank is Column1, Column2... (the first free); a name repeated, capitals aside, gets 2, 3 (Qty,
  Qty2). Without headers (or when Excel guesses there are none, a first row of numbers), a header row of Column1... is
  put above the range, the cells there moved down across the table's width. A header alone gets an empty data row
  below, the cells under it moved down, and `insertRow="1" insertRowShift="1"`.
- **Total Row** (`setTotalsRow`, ShowTotals): the range a row longer, `totalsRowCount="1"`, the autoFilter staying
  over the header and data; the first time, "Total" in the first column (`totalsRowLabel="Total"`) and in the last
  `SUBTOTAL(109,Table1[Price])` (`totalsRowFunction="sum"`) for numbers, else `SUBTOTAL(103,...)` (count). A filled
  row below moves down. Off: the row emptied and the range shorter, each column's total kept in the part, written
  back when it comes on again.
- **The total row's drop-down** (`setTotalFunction`, TotalsCalculation): Average 101, Count 103 (COUNTA), Count
  Numbers 102, Max 104, Min 105, Sum 109, StdDev 107, Var 110 in `SUBTOTAL(n,Table1[Col])`; None empties the cell; a
  function in the first column takes the place of its label. A formula typed in a total cell is
  `totalsRowFunction="custom"` with `<totalsRowFormula>`; text typed there its `totalsRowLabel`.
- **Header Row** off: the header cells emptied, `headerRowCount="0"` (before `totalsRowShown`), no autoFilter; on, the
  names written back above the data.
- **Filter Button** off: a `<filterColumn colId hiddenButton="1"/>` for each column; **Data > Filter** in a table takes
  the autoFilter away altogether.
- **Banded Rows, Banded Columns, First Column, Last Column**: `showRowStripes`, `showColumnStripes`,
  `showFirstColumn`, `showLastColumn`; the style **None** is a `tableStyleInfo` without a name.
- **Table Name**: a defined name's rules (`names.js validName`), not a name a table or defined name has; every formula
  naming the table renamed (`SUM(Table1[Qty])` is `SUM(Sales[Qty])`). Typing a header cell renames its column (a blank
  ColumnN). Typing just below a table without a total row makes it a row longer (AutoExpansion).
- **Resize Table**: the header must stay in its row and the new range overlap the old; columns kept by place, new
  ones named from their header cells; a total row is taken out (the cells under it moving up) and made again at the
  new last row.
- **Convert to Range** (Unlist): the table's look kept as each cell's own format, in the style's theme colours and
  tints (each line once: a cell's top and left, the right only in the last column, the bottom only in the last row);
  references to the table made absolute ranges with the sheet's name (`SUM(Sheet1!$D$3:$D$7)`), the total row's too;
  the part, its relationship and content type gone.

## In the editor

- **Insert > Table**, **Ctrl+T** and **Ctrl+L** open Create Table: "Where is the data for your table?" filled with the
  selection, or with one cell the data round it (its current region), as `=$B$2:$E$7`, and "My table has headers"
  ticked when the first row is text over a row that isn't (a guess of ours: Excel's own guess is measured only for a
  first row of numbers, which it takes for data). The table is made in the workbook's default style.
- **Home > Format as Table** opens the style gallery; outside a table, Create Table, then the table in the style
  picked; inside one, the style given to it.
- **Table Design** (shown while the active cell is in a table): Table Name (Enter or leaving the box; refused in
  Excel's words), Resize Table (its box, with Excel's note and refusal), Convert to Range (after Excel's question),
  Header Row, Total Row, Banded Rows, First Column, Last Column, Banded Columns, Filter Button, and the gallery with
  Clear. Each is one step of the history.
- A total cell shows a drop-down button, as a validation list does (Alt+Down too): None, Average, Count, Count
  Numbers, Max, Min, Sum, StdDev, Var.
- A table's filter buttons sort and filter it as the sheet's AutoFilter does (docs/FILTER.md); the filter is kept in
  the table's part, its sort too (docs/SORT.md).
- Typing a header renames its column, a total cell its label or custom total, a cell just below a table makes it a
  row longer: one step of the history with the typing.
- Tested in the browser: `test/table-tab.mjs`.

## Against Excel's PDF

`test/fixtures/tables.xlsx` (`test/fixtures/tables.mjs`) has a table in each built-in style with the options in turn,
one in a style of its own (every element, stripes two rows wide, a first header cell) and one with cells of their own
format, and rows of accent colours lightened 40, 60 and 80%. Excel opens it without repair. Excel's PDF against ours
(`tools/excel-compare/run.ps1`, then `compare-cover.py`, which compares the colour showing at points 3px apart inside
each of Excel's filled shapes, since Excel joins neighbouring cells of one colour into one rectangle):

- cell text: 1292 of 1292 cells on the same page, 1288 within 1px across and down;
- colours: 157678 of 161045 points (97.9%) the same colour, none off by a step. Open: the total row's double top line
  (our print draws a CSS double line, which the PDF's shapes don't show as two rectangles) and a few lines at page
  breaks.
