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

- Home > Editing > AutoSum (∑), and its menu: Sum, Average, Count Numbers, Max, Min, and More Functions..., which
  opens Insert Function and then Function Arguments for the function picked, as Excel does (docs/CALC.md, The Function
  Library). Formulas > AutoSum has the same menu. Alt+= is AutoSum.
- Several cells: the formulas entered as one step of the undo history, the totals selected with what they sum.
- One cell, or nothing found to sum: the cell opens for editing with Excel's proposal (`=SUM(B2:B5)`, or `=SUM()`),
  for Enter to enter it, as Excel leaves it, the range it takes outlined on the sheet (the range finder). The column
  above is looked at first, then the row to the left.
- Several areas (Ctrl+click): each summed as alone, all in one step of the undo history.
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

## Filtered lists, tables, several areas

Measured the same way on 17 more cases, in an Excel started for the purpose with its window off the screen (minimized,
it refused the command): `tools/excel-compare/excel-autosum2.ps1` on `autosum2-cases.json` (from
`autosum2-cases.mjs`), giving `test/autosum2-excel.json`; `test/autosum2.test.js` expects Excel's formula, value and
format in every cell and its selection after: 17 of 17. Saved here and opened in Excel (`excel-readback.ps1`), 17 of
17 read back with Excel's formulas and values.

| Rule | Seen |
|---|---|
| Sum in a filtered list (the sheet's filter with a criterion on): `SUBTOTAL(9, ...)`, so that the rows filtered out aren't counted, for the totals under, beside and in the corner. | A list filtered to its apples: `=SUBTOTAL(9,B2:B6)`, 9 of 15; B2:D7 selected: D2 to D6 `=SUBTOTAL(9,B2:C2)`..., D7 `=SUBTOTAL(9,B7:C7)`. |
| Average, Count Numbers, Max and Min stay themselves in a filtered list. | `=AVERAGE(B2:B6)`, 3: the hidden rows counted. |
| A filter that hides nothing, or rows hidden by hand: SUM. | `=SUM(B2:B6)`, 15. |
| Blank cells under a table's column: `SUBTOTAL(109, ...)`, filtered or not, with plain references. A table's own numbers selected: SUM in the row under it. | `=SUBTOTAL(109,B2:B6)`: 9 filtered, 15 not. B2:C6 selected: B7 `=SUM(B2:B6)`. |
| Several areas: each summed as if selected alone, each grown by its total. | B2:B6 and E2:E4: E5 `=SUM(E2:E4)`, B7 `=SUM(B2:B6)`, B2:B7 and E2:E5 selected. |
| Text or a blank among the numbers selected: the whole column summed in the row below. | B4 text: `=SUM(B2:B6)`, 12. |
| A blank column beside a block of numbers: totals in it, row by row. | B2:D6: D2 `=SUM(B2:C2)` to D6. |
| Several blank rows selected under the numbers: Excel opens the first for editing with its proposal, as for one cell (and answers nothing more over COM until a key is pressed: that case stopped the run, and is left out of it). Done the same here. Blank rows beside numbers are filled at once. | B7:C9: B7 being edited. D2:D5 beside a table: filled. |

SUBTOTAL itself leaves out the rows a filter hides for every function number, as Excel's does (and rows hidden by hand
for 101 to 111): the engine's `rowFiltered` (src/calc/engine.js), which these totals' values (9 of 15) measured.

## Not yet

The cases Excel leaves being edited can't be measured over COM: one cell (here the column above is taken before the
row to the left, as Excel's proposal is described), the proposal Excel leaves in the first of several blank rows (seen,
not read), and blank cells with nothing above or to the left. A table's own numbers selected while the table is
filtered isn't measured (here SUM).
