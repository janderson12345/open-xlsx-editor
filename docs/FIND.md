# Find and Replace

How cells are searched and replaced (`src/ooxml/findops.js`, the bar in `src/editor/findbar.js`), and what Excel was
seen to do. Each rule was measured over COM (`Range.Find` and `FindNext`, `Cells.Replace`, on a new workbook in a
hidden instance). `npm test` (`test/find.test.js`) repeats each search and replacement and expects Excel's result.
After replacements across formulas.xlsx, Excel works out all 399 formulas as the editor does.

## In the editor

- Ctrl+F opens the bar under the ribbon (as open-docx-editor's), Ctrl+H with its Replace row; Home > Editing > Find &
  Select has both. Enter finds the next match, Shift+Enter the previous; Esc closes the bar and goes back to the sheet.
- Options as Excel's dialog has them: Match case, Match entire cell contents, Within (Sheet or Workbook), Search (By
  Rows or By Columns), Look in (Formulas or Values; Replace always looks in formulas, as Excel's does).
- Find All says how many cells match; Replace replaces the active cell when it matches and goes on to the next;
  Replace All replaces every match as one step of the undo history.
- The bar keeps focus while the matches are selected, as Excel's dialog does; its status line is read out by screen
  readers.

## Excel's rules

| Rule | Seen |
|---|---|
| A search starts at the cell after the active one and goes round to the start; the active cell is found last. | "apple" from A1: B1, C1, A1. From B1: C1, A1, B1. |
| Look in Formulas searches the text of the formula bar: a formula with its `=`, a percentage as a percentage, a date as a date. | `50%` finds 0.5 shown as 50%, `0.5` finds nothing; `1/1/2024` finds a date, `45292` doesn't; `=` finds the formulas. |
| Look in Values searches the text cells show. | `1` finds `=A2*2` showing 1 and the date 1/1/2024, not `=A1+A10`. `A2` finds nothing. |
| Case is ignored unless Match case. Match entire cell contents compares the whole text. | `Apple` with Match case: B1 only. `apple` whole: A1 only. |
| `*` stands for any characters, `?` for one, `~` makes the next character plain. | `a*b` whole: `a*b` and `axxb`; `a~*b`: `a*b` only; `a?xb`: `axxb`. |
| Replace changes every match within a cell (or the whole cell), and the result is entered as if typed: a number becomes a number, `=...` a formula, nothing empties the cell. | `apple` with `pear`: `pear pie`, `pinepear`. `1` with `9`: 12 becomes the number 92. `12` with `=1+1`: a formula. |
| In a formula, replacing is plain text: references aren't recognised. | `A1` with `B9`: `=A1+A10` becomes `=B9+B90`. |
| Replace doesn't look in values: a formula's result isn't replaced. | `apple` doesn't change `=A1&"x"`, which shows `applex`. |

The messages are Excel 365's: "We couldn't find what you were looking for. Select Options for more ways to search."
and "All done. We made 3 replacements." (Excel shows them in a message box; here they are in the bar's status line.
They come from Excel's dialog, which COM doesn't reach, so their wording isn't measured.)

## Not yet

Look in Notes and Comments, Format (finding and replacing formats), Go To and Go To Special, and the Find All list of
cells. (Replacing within a date now works as in Excel: docs/ENTRY.md.)

## Found on the way

Replacing 5 with 7 in formulas.xlsx left a column unsorted (0, 70, 60, 70, 80), where Excel's approximate `VLOOKUP` and
`MATCH` gave answers the engine didn't: Excel searches such lists its own way. The engine now does the same
(docs/CALC.md), and after those replacements Excel works out all 399 formulas as the editor does.
