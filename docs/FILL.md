# Fill and the fill handle

How cells are filled (`src/ooxml/fillops.js`), and what Excel was seen to do. Measured over COM: 42 cases of
`Range.FillDown`, `FillRight`, `FillUp`, `FillLeft` and `Range.AutoFill` (the fill handle) on a new workbook in a
hidden instance (`tools/excel-compare/excel-steps.ps1` on `fill-cases.json`, from `fill-cases.mjs`, giving
`test/fill-excel.json`). `npm test` (`test/fill.test.js`) fills the same cells and expects Excel's in every cell
(formula and value, number format, bold, fill): 42 of 42. Filled by the editor, saved and opened in Excel
(`excel-readback.ps1`), 42 of 42 read back as Excel made them.

## In the editor

- The fill handle: the small square on the selection's bottom-right corner (a crosshair over it). Dragged down, up,
  right or left (the way the pointer is furthest out of the selection), the cells it will fill are outlined; let go,
  they are filled and selected. Ctrl held while letting go fills the other way: a single cell counts on, more cells
  are copied as they are (Excel's Ctrl-drag, from its documentation; COM doesn't reach the drag).
- Ctrl+D (Fill Down) and Ctrl+R (Fill Right), and Home > Editing > Fill: Down, Right, Up, Left.
- Each is one step of the undo history. Not on a read-only workbook; `features: { editing: false }` turns them off.

## Excel's rules

| Rule | Seen |
|---|---|
| Fill Down, Right, Up, Left copy the selection's first row (column; last for Up and Left) over the rest: values, formulas moved as a copy moves them, formats. One row selected (one column), Ctrl+D copies the row above (Ctrl+R the column to the left). | A1 `=B1+1` bold, A1:A3 down: `=B2+1`, `=B3+1`, bold. A2:B2 down under 7, `=A1+1`: 7, `=A2+1`. |
| The handle on one number copies it; with Fill Series (or Ctrl), it counts on by one. | 1: 1, 1, 1. Series: 2, 3, 4. |
| Two numbers go on by their difference; three or more by the straight line through them (least squares), shown to 15 digits. | 1, 3: 5, 7, 9. 5, 3: 1, −1. 0.1, 0.2: 0.3 ... 0.6. 2, 4, 7: 9.33333333333333, 11.8333333333333 ... |
| Text ending (or starting) with a number counts on, by the step of two such cells; a number in the middle is copied with the text. | `Item 1`: `Item 2`. `Item 1`, `Item 3`: `Item 5`. `1 apple`: `2 apple`. `a1b`: `a1b`. |
| Day and month names, short and long, go round their lists in the step given and in the same capitals; quarters go round four. | `Mon`: `Tue` ... `Sun`, `Mon`. `January`, `March`: `May`. `JAN`: `FEB`. `Q1`, `Qtr1`: up to 4, then 1. |
| A date goes on by a day, a time by an hour. Two dates a whole number of months apart (the day kept, or the month's last day) go on by months; otherwise by their days. | 1/30/2024: 1/31, 2/1. 1/1 and 1/8: weekly. 1/31 and 2/29/2024: 3/31, 4/30, 5/31. 9:00: 10:00. |
| Formulas are moved as a copy moves them. | `=B1*2`: `=B2*2`. `=$B$1+B1`: `=$B$1+B2`. |
| Cells of different kinds repeat as a block, each run of like cells going on in its own way; plain text is copied. | `a`, `1`: `a`, `2`, `a`, `3`. `x`, `y`: `x`, `y`, `x`. Across, `Jan`, `10`: `Feb`, `11`, `Mar`, `12`. |
| Each column (row) of the selection is filled on its own; filling upwards or leftwards runs the series backwards. | `1`, `3` beside `Item 7`, `Item 8`: 5, 7, 9 beside `Item 9` ... 5, 6 filled up: 1 to 4. |
| The copied cells' formats repeat over the new cells. Copy Cells repeats the values too; Formatting Only the formats; Without Formatting the series alone. | 1 bold, 2: bold every other cell. |

## Not yet

The Auto Fill Options button after a fill, a double click on the handle (filling down beside the next column's
data), Flash Fill, Fill > Series... (its box: linear, growth, date unit, stop value), Fill > Justify and Across
Worksheets, custom lists, and a drag back into the selection (Excel clears those cells).
