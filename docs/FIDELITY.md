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
| 4 made by Excel (charts, pivot, formats, review) | 63 of 67 | 458 of 458 | all |
| 11 test fixtures (all but large.xlsx) | 94 of 105 | 5,789 of 5,789 | all |

Every part that changed is a sheet that was edited. (Before table references were worked out, the pivot workbook's
`xl/workbook.xml` also changed: its table's total row, `SUBTOTAL(109,SalesTable[Units])`, couldn't be checked, and the
editor asked Excel to calculate on opening. That made 62 of 67.)

The same folders can be pointed at any workbooks; the real ones stay outside the repository and are named by label.

## Cached results other readers see (10 November 2026)

A program that reads a workbook without calculating it (openpyxl with `data_only=True`, pandas, most viewers) sees
each formula's cached result, the `<v>` the last program to save it wrote. The editor writes a result whenever it
works one out, so after an edit those readers see the new values, not the ones from before it.

Measured on `test/fixtures/formulas.xlsx` (541 formula cells, none with a cached result, as the fixture script writes
them), Data!B2 set to 7, saved, and read back with openpyxl 3.1.5:

| Saved after | Results written | openpyxl reads as the editor worked them out |
|---|---|---|
| The edit alone | the 64 that depend on Data!B2 | 64 of 64 |
| The edit, then Ctrl+Alt+F9 (`calculate()`) | all 541 | 539 of 541 |

The two that differ are results that are empty text (`T(1)`, `MID` past the end of its text): the editor writes
`<v></v>` with `t="str"`, and openpyxl reads that as None. The results nothing changed keep the cached
values the file came with.
