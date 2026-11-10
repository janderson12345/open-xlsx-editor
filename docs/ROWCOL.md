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

## Not yet

A sheet with pictures, charts, comments, tables or pivot tables is refused ("Rows and columns can't be inserted or
deleted yet on a sheet with...", or "Cells can't..."): their anchors and ranges aren't moved yet. Also missing: page
breaks, pivot caches' source ranges on other sheets, and the Insert Options button (format from below, clear
formatting). One reference Excel rewrites its own way: `=SUM(B3:B3)`, B3 being the hidden half of A3:B3, is
`=SUM(A3:A3)` in Excel after cells are inserted at B3 shifting right (the merged range widening); here it follows B3 to
`=SUM(D3:D3)`.
