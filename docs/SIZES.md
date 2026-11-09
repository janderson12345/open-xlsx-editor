# Column widths, row heights, hiding

How columns are resized and rows given a height, hidden and shown again (`src/ooxml/sizes.js`), and what Excel was
seen to do. Each rule was measured over COM (`ColumnWidth`, `RowHeight`, `Hidden`, `AutoFit` on a new workbook in a
hidden instance) and in the file Excel saved. `npm test` (`test/sizes.test.js`) expects what Excel wrote; a workbook
resized here reads back in Excel with the widths, heights and hidden lines Excel itself had set.

## In the editor

- Drag a heading's edge (the cursor changes within 4 pixels of it): a guide shows where it will be, with Excel's tip,
  "Width: 20.00 (145 pixels)" or "Height: 30.00 (40 pixels)". When the line is one of several selected whole lines,
  they all take the size. Dragged to nothing, the line is hidden.
- Double-click a heading's edge: AutoFit.
- Home > Cells > Format: Row Height..., AutoFit Row Height, Column Width..., AutoFit Column Width, Hide and Unhide
  Rows and Columns; and the headings' right-click menu (Row Height... or Column Width..., Hide, Unhide).
- Ctrl+9 hides rows, Ctrl+0 columns; Ctrl+Shift+9 and Ctrl+Shift+0 show them again.
- Each is one step of the undo history.

## Excel's rules

| Rule | Seen |
|---|---|
| A width typed in characters becomes whole pixels: the characters times the maximum digit width (7 for Aptos Narrow and Calibri 11), rounded, plus 5. It is stored as those pixels over the digit width, cut to 1/256, with `customWidth`. | 20: `20.7109375`. 12.5: 93 pixels, `13.28515625`, shown as 12.57. 5: `5.7109375`. 255: `255.7109375`. |
| A width can be 0 to 255 characters, a height 0 to 409 points. | 255.01 and 409.5 refused. |
| A width or height of 0 hides the line. A hidden column is written `width="0" hidden="1"`; a hidden row `hidden="1"`. | Column H hidden: `<col min="8" max="8" width="0" hidden="1" customWidth="1"/>`. |
| Setting part of a range of columns splits its `col` element. | D:F at 5, then E at 9: three elements. |
| A row height is stored in points with `customHeight`. | 30: `ht="30" customHeight="1"`. |
| AutoFit Row Height takes the fixed height away; Excel fits the row to its text when it opens the file. | Row 2 at 30pt, AutoFit: 15, its `ht` gone. |
| AutoFit Column Width stores the pixels found with `bestFit`; an empty column is left as it is. | "Hello world": 77 pixels, `width="11"`. "W": 19, `2.7109375`. |

## AutoFit's width is approximate

Excel measures with Aptos Narrow (or the workbook's font) through Windows, rounding each character its own way;
browsers don't have Aptos Narrow, and the editor draws it with a stand-in. So AutoFit here is the cell's text measured
in the font the grid draws it in, plus 9 pixels: on eight measured strings (plain, bold, 24pt, numbers) within 4
pixels of Excel's width, and never so narrow that the text is cut.

## Not yet

Default Width..., a new row height for the whole sheet (selecting all and setting a height changes the used rows
only), AutoFit of merged or wrapped cells as Excel does it, and the unhide cursor at a hidden line's edge.
