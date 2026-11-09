# Inserting and deleting rows and columns

How whole rows and columns are inserted and deleted (`src/ooxml/rowcol.js`, references moved by `shiftRefs` in
`src/calc/tokens.js`), and what Excel was seen to do. Each rule was measured by doing the same thing in Excel over COM
(`Rows(n).Insert`, `Range("2:4").Delete`, `Columns(n).Insert`... on a new workbook in a hidden instance; they give the
format from above or the left, as the ribbon's Insert Sheet Rows does). `npm test` (`test/rowcol.test.js`) repeats
each step and expects Excel's result. A workbook changed here, worked out by Excel (`excel-values.ps1`), gives the
same results as ours: 406 of 406.

## In the editor

- Click a row or column heading to select the whole line (Shift+click extends; the corner selects the sheet), or
  Shift+Space (rows) and Ctrl+Space (columns); Ctrl+A selects the sheet.
- Ctrl++ inserts and Ctrl+- deletes as many rows or columns as are selected; with cells selected, Excel's question
  (entire row or entire column; shifting cells isn't there yet).
- The headings' right-click menu (Insert, Delete, Clear Contents), and Home > Cells > Insert and Delete.
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

## Not yet

A sheet with pictures, charts, comments, tables or pivot tables is refused ("Rows and columns can't be inserted or
deleted yet on a sheet with..."): their anchors and ranges aren't moved yet. Also missing: inserting or deleting cells
with the rest of the row or column shifting, page breaks, pivot caches' source ranges on other sheets, and the
Insert Options button (format from below, clear formatting).
