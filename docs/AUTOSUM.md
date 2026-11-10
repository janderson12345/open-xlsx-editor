# AutoSum

How AutoSum and its Average, Count Numbers, Max and Min work (`src/ooxml/autosumops.js`), and what Excel was seen to
do. Measured over COM: 22 cases, the ribbon's own command run with `CommandBars.ExecuteMso` (AutoSum, AutoSumAverage,
AutoSumCount, AutoSumMax, AutoSumMin) on a selection in a new workbook of a visible Excel
(`tools/excel-compare/excel-autosum.ps1` on `autosum-cases.json`, from `autosum-cases.mjs`, giving
`test/autosum-excel.json`; a hidden instance doesn't answer the command). `npm test` (`test/autosum.test.js`) does the
same here and expects Excel's formula, value and number format in every cell, and its selection after: 22 of 22.
Saved here and opened in Excel (`excel-readback.ps1`), 22 of 22 read back with Excel's formulas and values; one
spells the test's own starting format `$#,##0.00` as `\$#,##0.00`, in the cells it started in too.

Only selections of several cells were measured: on one cell Excel proposes the formula and leaves the cell being
edited, waiting for Enter, and answers nothing over COM until a key is pressed. That was seen on screen once, over a
column of dates: `=AVERAGE()` left open in the cell.

## In the editor

- Home > Editing > AutoSum (∑), and its menu: Sum, Average, Count Numbers, Max, Min. Alt+= is AutoSum.
- Several cells: the formulas entered as one step of the undo history, the totals selected with what they sum.
- One cell, or nothing found to sum: the cell opens for editing with Excel's proposal (`=SUM(B2:B5)`, or `=SUM()`),
  for Enter to enter it, as Excel leaves it.
- Not on a read-only workbook; `features: { editing: false }` turns them off.

## Excel's rules

| Rule | Seen |
|---|---|
| Numbers selected, a blank row under them in the selection: the totals go there, each over its column's selected numbers. Without one, they go in the row below (the first one free), and the selection takes in that row. | B2:B5 or B2:B6: B6 `=SUM(B2:B5)`. B6 filled: B7 `=SUM(B2:B5)`, B2:B6 selected. |
| One row of numbers: the total to its right. | B8:E8: F8 `=SUM(B8:E8)`. |
| A blank row and a blank column in the selection: totals both ways, and the corner over the total row. | B2:D6: D2 `=SUM(B2:C2)` ... B6 `=SUM(B2:B5)`, D6 `=SUM(B6:C6)`. |
| Blank cells selected under numbers: each sums up the column, past blanks right above it, to the first blank or text after the numbers (or the top). Under no numbers, to the left along the row. | B6:C6 under a table: `=SUM(B2:B5)`. A blank at B4: `=SUM(B5)`. Text at B4: `=SUM(B5)`. A blank row then numbers: `=SUM(B2:B6)`. 2024 at the top: `=SUM(B1:B5)`. D2:D5 beside: `=SUM(B2:C2)` each. |
| The selection's first column decides the rows for every column, and the number format. | Text only in column B: column C's total also `=SUM(C4:C5)`. Currency in B, percentages in C: both totals `$#,##0.00`. |
| SUM subtotals in the column: only they are summed. | `=SUM(B7,B4)`. |
| Average, Count Numbers, Max and Min the same, with their own function. | `=AVERAGE(B2:B5)`, `=COUNT(B4:B5)`, `=MAX(...)`, `=MIN(...)`. |
| Dates aren't taken. | Over a column of dates: `=AVERAGE()`, left being edited. |

## Not yet

More Functions... (the Insert Function box), AutoSum in a filtered list (SUBTOTAL), the proposed range shown on the
sheet while the cell is being edited, a selection of several areas, and the cases not measured: several blank rows
selected, and blank cells with nothing above or to the left.
