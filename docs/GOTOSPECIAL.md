# Go To Special, and more than one range selected

The cells of a kind found and selected (`src/ooxml/special.js`, the box in `src/editor/specialbox.js`), and a selection
of several ranges (`src/editor/selection.js`, the commands over them `src/ooxml/multiarea.js`). Measured over COM:
`test/fixtures/special.xlsx` (made by `test/fixtures/special.mjs`, made-up data) opened in a hidden Excel, 51 cases
asked as the box asks them (`Selection.SpecialCells`, `ActiveCell.CurrentRegion` and `CurrentArray`,
`RowDifferences`, `ColumnDifferences`, `Precedents` and `Dependents`). `npm test` (`test/special.test.js`) asks the
same and expects the cells Excel found.

The commands over several ranges: `tools/excel-compare/excel-multirange.ps1` on `multirange-cases.json` (79 cases,
answers in `test/multirange-excel.json`), each a new workbook with A1:F12 filled with row * 10 + column, two ranges
selected, the ribbon's command run (`CommandBars.ExecuteMso`; whether its button is enabled, `GetEnabledMso`) or the
object model's call where the ribbon's only opens a box. The boxes Excel puts up are read through UI Automation and
answered by `dialog-watch.ps1` (window messages to that Excel alone; no key or click goes through the desktop).
`test/multirange.test.js` does each case as the editor does and expects Excel's cells; `test/editing-multirange.mjs`
runs them in the browser.

## In the editor

- Home > Find & Select > Go To Special..., or the Go To box (F5, Ctrl+G) and its Special... button: Excel's box, its
  choices in two columns, each with its options. The last choice is kept for the next time.
- Find & Select's shortcuts as Excel's: Formulas, Notes, Conditional Formatting, Constants, Data Validation.
- Ctrl+click (or Ctrl and drag) adds a range to the selection; Shift and the arrows change the range added last. The
  ranges are shaded; there is no fill handle while there are several, as in Excel.
- Over several ranges, the commands do what Excel's do (the table below). Copy and Cut refuse with Excel's "This
  action won't work on multiple selections."

## Commands over several ranges

| Command | Excel, over two ranges | Seen |
|---|---|---|
| Formatting (font, fill, alignment, number format, cell styles), Borders, Clear, Data Validation | Each range. Borders outline each. | Percent Style over A2:B3 and D6:E7: all eight cells 0%. |
| Fill Down, Right (Ctrl+D, Ctrl+R) | Each column (row) filled over the cells selected in it, from its first one. | A2:B4 and A7:B9 down: A7:B9 from A2:B2. A2:B4 and B3:C5: B3:B5 from B2, C4:C5 from C3. |
| AutoSum | Each range summed as alone (docs/AUTOSUM.md). | A2:A4 and C2:C4: A5 and C5. |
| Insert, Delete Sheet Rows and Columns | The lines of every range; overlapping ones taken together. Inserting, each range's rows above its own first row. | Rows of A2:B3 and D6:E6: 2, 3 and 6. A2:B4 and D3:E5 insert one row above 2 and three above 3. |
| Insert, Delete Cells (shifted) | Each range's cells, as if done where it was; overlapping ranges refused: "Cannot use that command on overlapping selections." | A2:B3 and B6:C6 up: column B loses rows 2, 3 and 6. |
| Hide, Unhide, Row Height, Column Width, AutoFit | The lines of each range. | Ctrl+9 on A2:B3 and D6:E6: rows 2, 3, 6. |
| Merge & Center, Merge Cells, Merge Across, Unmerge | Each range, after Excel's "Merging cells only keeps the upper-left value and discards other values." Overlapping ranges: nothing. | A2:B3 and D6:E7: both merged; Across, each of their four rows. |
| Find, Replace | Only in the selection, every range of it (as with one range of more cells); Find Next moves the active cell there, the ranges kept. | "6" in A2:B3 and D6:E7: D6, E6. |
| Go To Special | Looks in every range. | Constants in A2:B3 and D6:E7 (B2 a formula, E7 empty): A2, A3:B3, D6:E6, D7. |
| Define Name, the Name Box | The name refers to every range, `=Sheet1!$A$2:$B$3,Sheet1!$D$6:$E$7`. | |
| Create from Selection | Names from each range. | Top row of A2:B3 and D6:E7: A3, B3, D7, E7. |
| Conditional Formatting | One rule over every range. | Applies to A2:B3,D6:E7. |
| Charts | Ranges over the same rows as one table of their columns; otherwise each range as alone, its series after the ones before. | A1:A5 and C1:C5: two series. A1:B3 and D6:E8: four, a column each. |
| Print Selection | Each range on pages of its own (`preparePrint({ selection: true })`). | Two pages. |
| The status bar's Sum, Count, Average | Every range; a cell in two of them counted twice. | A2:B3 and B3:C4: Sum 256, Count 8, Average 32 (Excel's SUM and COUNTA of the selection; COM doesn't reach the status bar). |
| Sort (A to Z, Z to A, Custom Sort), Filter, Text to Columns, Remove Duplicates, Fill Justify, Group, Subtotal | Refused, before any box: "This can't be done on a multiple range selection. Select a single range and try again." | |
| Insert Table, Format as Table, PivotTable | Greyed out. Here the buttons (and Ctrl+T) do nothing and say so. | |

## Excel's rules

| Rule | Seen |
|---|---|
| One cell selected: the used range is looked in (its first to its last cell with a value, a formula or a format). More selected: only those cells. | Constants from A1 find cells all over the sheet; in A1:D6, only there. |
| Visible cells only, Conditional formats and Data validation, from one cell, look at the whole sheet. | Visible cells only from A1: `A1:E3,G1:XFD3,A5:E1048576,G5:XFD1048576` (row 4 and column F hidden). |
| Constants and Formulas by type: numbers (dates among them), text, logicals, errors. A formula's type is its result's; the cells of an array formula are formulas. | `=""` is a text formula; G2:G4 `{=B2:B4*10}` are formulas, not constants. |
| Blanks: no value and no formula. A formula giving "" isn't blank; a formatted empty cell is. | H10 (bold, empty) is a blank. |
| A merged range is found whole when any cell of it is. | Constants find A9:B9 for the text in A9; Blanks in A8:C10 find all of it, A9 with them. |
| Current region: grown from the active cell while any cell next to it (diagonals too) has something; an empty cell alone is itself. | B3: A1:E6. M20, nothing near: M20. |
| Row differences: in each row, the cells unlike the one in the active cell's column. Column differences, by rows. | A12:C16 by column A: C12:C13, B14, C15, B16. |
| A formula is unlike any constant, even of the same value; two formulas are alike when the same relative to their cells (as R1C1 writes them). Text is compared without regard to case. | `=A14` against 5: unlike. `=A20*2` in A21 and `=B20*2` in B21: alike. `Abc` and `abc`: alike. |
| Precedents and Dependents: of the selected cells, on this sheet; Direct only, or All levels. An array formula is found whole. | Dependents of B2: C2, E2, G2:G4; all levels add E3. |
| Last cell: the used range's bottom-right cell. | K21. |
| Conditional formats and Data validation: All, or Same (the rules of the active cell). | Same as B3: B2:B6. |
| Nothing found: Excel's "No cells were found." | Data validation, Same as A1. |

## Where it differs

- Excel lists the ranges found in an order of its own (Constants, Numbers from A1 starts at D4), and selecting them
  makes the first range's top-left cell active, even in a hidden row. Here the ranges go top to bottom, left to right,
  and the first cell found that is in sight is made active. The cells selected are the same.
- Objects (selecting the sheet's pictures, charts and shapes) isn't offered: drawings can't be selected yet.
- Copy over several ranges that line up (the same rows or the same columns), which Excel allows, is refused.
- Merging over several ranges asks once; Excel asks for each range that would lose values.
- Insert Table, Format as Table and PivotTable stay enabled on the ribbon over several ranges (they do nothing),
  where Excel's are greyed out; Excel's Ctrl+T wasn't measured (no key is sent to it).
- Ungroup refuses as Group does; only Group was measured. Text to Columns' refusal was measured through COM's
  TextToColumns (its button has no idMso to run); its words are the same as the others'.
