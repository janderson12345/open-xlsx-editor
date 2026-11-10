# Data validation

How data validation works (`src/ooxml/validation.js`, the button and message in `src/editor/dvmarks.js`), and what
Excel was seen to do. Measured over COM: 15 cases, each a column of values under one rule (`Validation.Add` in a
hidden Excel), every cell judged by Excel (`Range.Validation.Value`), and the workbook saved so its `dataValidations`
XML could be kept (`tools/excel-compare/excel-dv.ps1` on `dv-cases.json`, from `dv-cases.mjs`, giving
`test/dv-excel.json`). `npm test` (`test/validation.test.js`) builds each sheet here with Excel's own XML for its rule
and expects Excel's verdict for every cell: 15 of 15.

## In the editor

- A value typed into a cell with a rule (or a formula, by its result) is judged as Excel judges it. Refused, Excel's
  box: Stop offers Retry (back into the cell's editor with the text) and Cancel; Warning asks "Continue?" with Yes
  (entered), No (back into the editor) and Cancel; Information offers OK (entered) and Cancel. The rule's own title
  and message are shown, or Excel's: "This value doesn't match the data validation restrictions defined for this
  cell." Pasting, filling and the API's setCell aren't checked, as Excel doesn't check them.
- A list rule's cell, when active, has a dropdown button beside it (unless the rule hides it); the button, or Alt+Down,
  shows the choices, and one picked is entered.
- A rule's input message shows under the active cell and is read with it by a screen reader.

## Excel's rules

| Rule | Seen |
|---|---|
| A blank passes when the rule ignores blanks (the default), else it is refused. An error never passes. | Whole number 1 to 10, blanks not ignored: the blank refused. |
| Whole number, decimal, date and time rules take numbers, a formula's result too; text is refused, even a number stored as text. A whole number has no fraction. | 1 to 10: `5.5`, `abc`, `'7` refused; `=2+3` passes. |
| They compare as Excel compares, TRUE above every number. | Greater than 4: TRUE passes; between 1 and 10: TRUE refused. |
| Numbers are compared to 15 significant digits. | 17:00 is saved as 0.708333333333333, and 17:00 passes "between 9:00 and 17:00". |
| Text length is of the shown text. | At most 3: `123` passes, `12.5` and `TRUE` are refused. |
| A list typed into the rule matches case, and spaces round the value don't count. | "a,b,c": `A` refused; `b ` and ` b` pass. Numbers by their text: `'1` and `2.0` pass "1,2,3". |
| A list from cells matches case aside, numbers by value. | `Y` passes x, y, z; `5.0` passes 5. |
| A custom rule is its formula, relative to the range's first cell, TRUE or a number other than 0 to pass. | `=A1>B1` row by row. |

The boxes' buttons and words follow Excel's interface, which COM doesn't reach.

## Not yet

The Data Validation box to add or change rules, Circle Invalid Data, a list's choices from a name defined as a formula
of another kind, and rules following their cells when rows are inserted or sorted (their ranges move with inserted
rows already).
