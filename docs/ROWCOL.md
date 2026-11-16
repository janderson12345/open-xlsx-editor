# Inserting and deleting rows, columns and cells

How whole rows and columns are inserted and deleted (`src/ooxml/rowcol.js`, references moved by `shiftRefs` in
`src/calc/tokens.js`), and what Excel was seen to do. Each rule was measured by doing the same thing in Excel over COM
(`Rows(n).Insert`, `Range("2:4").Delete`, `Columns(n).Insert`... on a new workbook in a hidden instance; they give the
format from above or the left, as the ribbon's Insert Sheet Rows does). `npm test` (`test/rowcol.test.js`) repeats
each step and expects Excel's result. A workbook changed here, worked out by Excel (`excel-values.ps1`), gives the
same results as ours: 406 of 406.

## In the editor

- Click a row or column heading to select the whole line (Shift+click extends; the corner selects the sheet), or
  Shift+Space (rows) and Ctrl+Space (columns); Ctrl+A selects the sheet.
- Ctrl++ inserts and Ctrl+- deletes as many rows or columns as are selected; with cells selected, Excel's Insert or
  Delete box: Shift cells right (left), Shift cells down (up), Entire row, Entire column.
- The headings' right-click menu (Insert, Delete, Clear Contents), and Home > Cells > Insert (Insert Cells..., Insert
  Sheet Rows, Insert Sheet Columns) and Delete.
- One step of the undo history: the workbook goes back to how it was (`Workbook.snapshot` and `restore`).

## Excel's rules

| Rule | Seen |
|---|---|
| A reference at or below an inserted row moves down, $ or not. A range the insert falls inside grows; one ending just above it doesn't. | Row inserted at 3: `=A5` is `=A6`, `=$A$4` is `=$A$5`, `=SUM(A1:A5)` is `=SUM(A1:A6)`. At 6: `=SUM(A1:A5)` stays. At 1: `=SUM(A2:A6)`. |
| A deleted row takes away the references to it: a range shrinks, one whose rows are all deleted becomes `#REF!`, its sheet kept. | Row 3 deleted: `=SUM(A1:A5)` is `=SUM(A1:A4)`, `=SUM(A3:A3)` is `=SUM(#REF!)`; Rows 2 to 4: `=Sheet1!A4` is `=Sheet1!#REF!`. |
| Formulas on other sheets, defined names and chart series follow the same way. | `Four` (`Sheet1!$A$4`) is `$A$5` after an insert at 3, `#REF!` after rows 2 to 4 go. |
| Columns follow the same rules. | Column A deleted: `=SUM(A1:A5)` is `=SUM(#REF!)`. |
| New rows take the row above's height and format, and the cells of that row whose format isn't the row's own. New columns take the width and format of the column to their left, and its cells' formats. | Row 2 bold, 30pt, B2 yellow: the new row 3 is bold, 30pt, B3 yellow. Column B 20 wide: the new C is 20 wide. |
| New rows and columns take the outline level of the one before them too, shown and never collapsed, and the group they join opens (docs/OUTLINE.md). | Rows 2:4 grouped, a row inserted at 3: level 1; in a collapsed group, shown and the summary no longer collapsed. |
| Merged ranges move, grow and shrink as ranges do; one whose lines are all deleted goes. | E2:E4: E2:E5 after an insert at 3, E2:E3 after a delete of 3, gone after 2 to 4. |
| An insert that would push filled cells off the sheet is refused, with Excel's message; so is an insert or delete that would split an array formula. | "To prevent possible loss of data, Excel cannot shift nonblank cells off of the worksheet..." |

The sheet's other ranges move too: its used range, hyperlinks, conditional formats and data validations (their
formulas as well), the filter, ignored errors, protected ranges, and the view's selection.

## Inserting and deleting cells

The same change limited to a band: shifting cells down or up moves only the selection's columns, from its first row
on; right or left only its rows, from its first column on (`shiftCells` in `src/ooxml/rowcol.js`, references by
`shiftRefs` with a band). Measured over COM: 25 cases of `Range.Insert` and `Delete`, with and without a shift, on
one sheet of numbers with 24 formulas reading it, formulas on the next sheet and defined names
(`tools/excel-compare/excel-steps.ps1` on `shift-cases.json`, from `shift-cases.mjs`, giving
`test/shift-excel.json`). `test/shift.test.js` does each and expects Excel's cells, formats, merged ranges, formulas
and names. Saved by the editor and opened in Excel (`excel-readback.ps1`), 22 of the 23 that Excel does read back as
Excel made them; the other is the reference below.

| Rule | Seen |
|---|---|
| A reference within the shifted columns (rows) follows the rules for rows (columns) above; one reaching outside them stays as it is, whole rows and columns too. | B3:C4 inserted down: `=B5` is `=B7`, `=SUM(B1:C10)` is `=SUM(B1:C12)`, `=SUM(B2:C3)` grows to `=SUM(B2:C5)`; `=SUM(A1:D10)`, `=SUM(B3:D4)`, `=SUM(B:B)` stay. |
| A deleted cell's references go, as with rows; and a range the deleted cells cross, all of its rows (columns) deleted, loses the end they cover. In its middle they leave it. | B3:C4 deleted up: `=SUM(B3:C4)` is `=SUM(#REF!)`, `=SUM(B3:D4)` is `=SUM(D3:D4)`, `=SUM(A3:C3)` is `=SUM(A3:A3)`. Left: `=SUM(B2:C3)` is `=SUM(B2:C2)`. |
| New cells take the format of the cell above (shifting down) or to the left (right). | B2 yellow, C2 bold: the new B3:C4 yellow and bold. A3 blue: the new B3:C3 blue. |
| Left to choose, Excel shifts right (left) when the selection is taller than it is wide, down (up) otherwise; the Insert and Delete boxes start there. | B3:C5: right. B3:D4, B3:C4, B3: down. |
| A merged range within the band moves, grows and shrinks as a range does; one the shift would split is unmerged, Excel asking first: "This operation will cause some merged cells to unmerge. Do you wish to continue?" | B7:C7 is B9:C9 after B3:C4 down; A3:B3 grows to A3:D3 shifting right; A7:B7 and F3:F5 are unmerged. |
| An array formula the shift would split is refused ("You can't change part of an array."); one within the band moves. Pushing filled cells off the sheet is refused. | J4:J5 with B3:C4 shifted right: refused. |

The sheet's other ranges (hyperlinks, conditional formats, data validations, the filter, the view's selection) move
when they are within the band and stay otherwise; that isn't measured. The question's wording comes from Excel's
dialog, which COM doesn't reach.

## Notes, pictures, charts and shapes

Measured by `tools/excel-compare/excel-anchors.ps1` on `anchor-cases.mjs --no-table`. The sheet holds a stock list on
B2:D12, a column chart anchored F3 to the start of L15, a rectangle M5 to P9, and notes at A1, B6, D12 and F20. That
sheet is `test/fixtures/anchors.xlsx`, written by `make-fixtures.mjs`. Twenty steps were measured: whole rows and
columns inserted and deleted around, inside and across each object, and cells shifted down, up and right
(`src/ooxml/anchors.js`).

| What | As Excel does it |
|---|---|
| A note or threaded comment | Moves with its cell, also when cells shift, and is deleted with it. |
| A picture, chart or shape (a two-cell anchor) | Each corner after the change moves with it. An end corner on the boundary where lines are inserted stays, so the object doesn't grow. Lines deleted inside it shrink it. One whose whole extent is deleted is deleted, a chart with its part. |
| The same when cells shift | Not moved. |
| One that moves but doesn't size (`editAs="oneCell"`), or doesn't move (`absolute`) | Moved by its start, or left where it is. Not measured. |

Results: all 20 steps put the chart, the rectangle and the notes where Excel did (`test/anchors.test.js` against
`test/anchors-excel.json`). The editor's 20 results were saved (`anchors-ours.mjs`) and opened in Excel
(`excel-anchors-check.ps1`), and Excel read every one back the same.

## Tables

Measured on the same sheet with the table Stock on B2:D12 and `=SUM(Stock[Qty])+COUNTA(Stock[Price])` in A20
(`anchor-cases.mjs`, `test/fixtures/anchors-table.xlsx`). The same 20 steps were done, plus 15 at the table's edges,
its header row, all of its data and part of it (`src/ooxml/anchors.js`).

| What | As Excel does it |
|---|---|
| Lines inserted strictly inside the table (a row above its last row, a column between two of its columns) | It widens. New columns are named Column1, Column2 (the first names not taken), in its header too. |
| A row inserted just below it, a column just right of it | It stays as it was. |
| A row inserted above it, a column at its left edge | It moves. |
| Lines deleted inside it | It narrows. A formula's reference to a deleted column becomes `#REF!` (`=SUM(#REF!)+COUNTA(Stock[Price])`). |
| All of its data rows deleted (its header kept) | Excel keeps the first data row as it is and deletes the others, so the rows deleted are one fewer. |
| All of it deleted | It is deleted, and its references become `#REF!`. |
| Its header row deleted without the rest | Refused. COM says only that the Delete failed, so the editor's words are its own. |
| Cells shifted through part of it | Refused: "This won't work because it would move cells in a table on your worksheet." Across its whole width (or height) it widens, narrows or moves as for lines. |

Results: all 35 steps match Excel (`test/tables.test.js` against `test/tables-excel.json`). That covers the table's
range, its column names, where the formula went with its text and value, and the chart, rectangle and notes. Six of
the steps are refused, as Excel refuses them. The editor's 29 results, opened in Excel, read back 29 of 29 as Excel did
them.

One rule explains the chart Excel left one row tall when rows 3:14 were deleted: those rows held all the table's data,
so Excel kept row 3 and deleted only 4:14.

Not measured: a table's totals row (deleting it without the rest of the table is refused), and calculated columns
(Excel fills their formula into rows inserted inside the table; the editor doesn't yet).

## PivotTables

Measured by `tools/excel-compare/excel-pivot-lines.ps1` with a pivot table at Pivot!C5 (rows Region, Sum of Units)
of Src!A1:E19: 18 steps on the pivot's sheet and 10 on its source's (`src/ooxml/anchors.js`).

| What | As Excel does it |
|---|---|
| Lines before the pivot | It moves. |
| Lines after it | It stays. |
| All of it deleted | It is deleted. Its cache stays in the workbook. |
| Lines through it, cells shifted through part of its rows or columns | Refused: "We can't make this change for the selected cells because it will affect a PivotTable..." |
| Cells shifted across its whole width (or height) | It moves. |
| Lines on the source's sheet | The cache's source range follows, as a reference does: it grows, shrinks and moves. When all of it is deleted, Excel leaves it as it was. |

Results: 28 of 28 steps match Excel (`test/pivot-lines.test.js` against `test/pivot-lines-excel.json`): the
pivot's location, its source, and a value beside it. The editor's 22 results opened in Excel read back 22 of 22 as
Excel did them. Excel's own refresh of each moved pivot left it where it was.

One step Excel takes its own way: a cell inserted inside the pivot, shifting down, leaves Excel's pivot with an empty
cell in it. The editor refuses it, as it refuses any cells shifted through part of a pivot.

## Sparklines and data validation

Measured by `tools/excel-compare/excel-refmoves.ps1` on `refmove-cases.mjs`: a sheet with a group of line sparklines
O2:O5 over B2:M5, a column sparkline at O7, one at Q2 over another sheet's B2:M2, four validation rules on B10:E15 (a
list of $H$10:$H$12, whole numbers 1 to $I$10, a custom `D10>C10`, a list of `Data!$A$1:$A$3`, which Excel writes in
the Excel 2010 form, `x14:dataValidation` in the sheet's extLst), notes at B12 and F20 and a threaded comment
at C13. Forty-two steps (whole lines, cells shifting each way, lines on the other sheet, sorts) were
done in Excel and the workbook it saved compared (`src/ooxml/rulemoves.js`; `test/fixtures/refmoves.xlsx`).

| What | As Excel does it |
|---|---|
| A sparkline | Its cell moves with it, and its data as a reference does (the data on another sheet follows that sheet's lines). Deleted with its cell, or when all of its data is deleted (column B:M deleted: O2:O5 and O7 went though their cells stayed). |
| Sparklines and new lines | Lines and cells inserted take the sparklines of the line or cell before them, in the same group, their data moved as a copy moves it: column P inserted after O gives P2 over C2:N2; rows 3:4 inserted give O3 over B3:M3 and O4 over B4:M4, Q3 and Q4 over Data!B3:M3 and B4:M4. |
| A validation rule's ranges | Move, grow and shrink as ranges do, the Excel 2010 form's too. Lines and cells inserted take the rules of the line or cell before them: B10:B15 is B10:C15 after a column is inserted at C; all four end at row 16 after a row is inserted there; D10:D15 gets E12:E13 when E12:E13 is inserted shifting right. |
| A rule's formulas | Stay relative to its first cell: references move as references do; when the first cell is deleted they read from the first cell left (D10:D15's `D10>C10` stays `D10>C10` after rows 10:12 go, the rule now D10:D12). |
| Cells shifted through part of a rule | The part that moves is cut from it: a rule of its own when its formulas have relative references (D12, shifted left to C12: `C12>#REF!`), else it stays in the rule (E10:E11 E13:E15 D12). A part deleted is taken out when the formulas are relative (D10:D11 D13:D15), and left in when they are absolute (C10:C15 kept C12). |
| A sort | Sparklines, validation rules and conditional formats stay where they are, also when the sparklines' cells are sorted. |

Results: the sparklines and rules after 41 of the 42 steps match Excel's (`test/refmoves.test.js` against `test/refmoves-excel.json`); the other is
below. The editor's 43 results (the start and each step) opened in Excel without a repair and read back 42 of 43 as
Excel had made them (`excel-refmoves.ps1 -Books test/out/refmoves`), leaving out the threaded comment's placeholder
box, which Excel lays out again itself when it opens any workbook.

The one difference: with B12:C12 deleted shifting left, Excel took B12 out of B10:B15's list rule, whose formula is
absolute, though it left C12 in C10:C15's and, with C12:E12 deleted, E12 in E10:E15's. The editor leaves an absolute
rule whole each time. Excel also leaves rules overlapping there (C12 is in two), as the editor does.

## Not yet

A sheet with slicers or timelines is refused ("Rows and columns can't be inserted or deleted yet on a sheet with...",
or "Cells can't..."): their positions aren't moved yet. Also missing: page breaks and the Insert Options button (format
from below, clear formatting). Rules of the Excel 2010 form are moved
but not read: their cells aren't checked, shown a list or circled (E12 above). Conditional formats' ranges move as
before; whether new lines take them as they take validation rules isn't measured.

One reference Excel rewrites its own way: `=SUM(B3:B3)`, B3 being the hidden half of A3:B3, is `=SUM(A3:A3)` in Excel
after cells are inserted at B3 shifting right (the merged range widening); here it follows B3 to `=SUM(D3:D3)`.
