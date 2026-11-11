# Go To Special, and more than one range selected

The cells of a kind found and selected (`src/ooxml/special.js`, the box in `src/editor/specialbox.js`), and a selection
of several ranges (`src/editor/selection.js`). Measured over COM: `test/fixtures/special.xlsx` (made by
`test/fixtures/special.mjs`, made-up data) opened in a hidden Excel, 51 cases asked as the box asks them
(`Selection.SpecialCells`, `ActiveCell.CurrentRegion` and `CurrentArray`, `RowDifferences`, `ColumnDifferences`,
`Precedents` and `Dependents`). `npm test` (`test/special.test.js`) asks the same and expects the cells Excel found.

## In the editor

- Home > Find & Select > Go To Special..., or the Go To box (F5, Ctrl+G) and its Special... button: Excel's box, its
  choices in two columns, each with its options. The last choice is kept for the next time.
- Find & Select's shortcuts as Excel's: Formulas, Notes, Conditional Formatting, Constants, Data Validation.
- Ctrl+click (or Ctrl and drag) adds a range to the selection; Shift and the arrows change the range added last. The
  ranges are shaded; there is no fill handle while there are several, as in Excel.
- Over several ranges, formatting (Home's font, fill, borders, alignment, number format) and Clear go to each range;
  Data Validation's box sets the rule on each. Borders outline each range. Copy and Cut refuse with Excel's "This
  action won't work on multiple selections." Other commands act on the range added last.

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
