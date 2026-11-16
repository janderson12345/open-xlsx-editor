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
