# Save fidelity

What survives when a workbook is opened in the editor, edited and saved. The editor keeps every part of the package
it doesn't change byte for byte, and inside the parts it does change it keeps the XML it doesn't model where it was
(`CLAUDE.md`, "Keep what isn't understood"). This benchmark measures that, and has Excel itself check the result.

## How it is measured

1. `tools/excel-compare/excel-make-books.ps1 <folder>` has a hidden Excel make four workbooks of what editors tend to
   lose: three kinds of chart, a pivot table with its cache, a table with a total row, a three-colour scale, data bars,
   an icon set, sparklines, list validation with an input message, notes, a threaded comment, hyperlinks (a web page
   and a place in the workbook), shapes and a text box, freeze panes, grouped rows, a protected sheet, print titles and
   a footer, and a defined name. They are written outside the repository (they hold Excel's own parts).
2. `node tools/excel-compare/fidelity.mjs <folder> <out>` opens each workbook in the editor, types into one empty cell
   (two rows under the first sheet's data), saves, and compares the two packages part by part.
3. `pwsh tools/excel-compare/excel-fidelity.ps1 <folder> <out>` has a hidden Excel open the original and the saved copy
   and compare, sheet by sheet, every cell of the original's used range (formula, value, number format; the typed cell
   aside) and the count of charts, pivot tables, tables, shapes, notes, threaded comments, hyperlinks, conditional
   formats, sparkline groups and validated cells, whether the sheet is protected, its print titles, and the workbook's
   names. A saved copy Excel won't open is reported as such.

The comparison was checked against a control: a saved copy with one cell changed (found: `Review!C5`) and one with a
chart taken out of its drawing (Excel wouldn't open it: found).

## Results (10 November 2026)

| Workbooks | Parts kept byte for byte | Cells Excel reads as before | Features Excel counts as before |
|---|---|---|---|
| 4 made by Excel (charts, pivot, formats, review) | 62 of 67 | 458 of 458 | all |
| 11 test fixtures (all but large.xlsx) | 94 of 105 | 5,789 of 5,789 | all |

Every part that changed is a sheet that was edited, except one: in the pivot workbook `xl/workbook.xml` gains
`fullCalcOnLoad="1"`, because its table's total row (`SUBTOTAL(109,SalesTable[Units])`) uses a table reference the
calculation engine doesn't work out yet, and the editor then asks Excel to calculate on opening rather than leave a
value it couldn't check.

The same folders can be pointed at any workbooks; the real ones stay outside the repository and are named by label.
