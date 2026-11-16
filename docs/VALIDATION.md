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

## Writing rules

Data > Data Validation... opens Excel's box for the selection, holding the active cell's rule. It has three tabs:

- **Settings:** Allow (any value, whole number, decimal, list, date, time, text length, custom), Data (between,
  not between, equal to, not equal to, greater than, less than, and or equal to), the values, which are labelled
  Minimum, Maximum, Value, Source or Formula as the choices make them, Ignore blank, and In-cell dropdown.
- **Input Message:** whether to show it, its title and its message.
- **Error Alert:** whether to show it, its style (Stop, Warning, Information), its title and its message.

OK gives the selection the rule; Clear All, then OK, takes its rules away. The selection's cells come out of any
rule they were in (that rule's ranges split around them), and the new rule is added after the others, as Excel does.
Each change is one step of the undo history (`src/ooxml/dvedit.js`).

A rule is written as Excel writes it:

- **Attributes** in its order, `operator` left out for between.
- **A typed list** in quotes (`"a,b,c"`).
- **A reference or formula** without its `=`.
- **A date or time** as its serial number, to 15 significant digits (17:00 is 0.708333333333333).
- **A list without its dropdown** as `showDropDown="1"`.

Measured: each of the 15 rules of `dv-cases.mjs`, written here, gives the `dataValidation` element Excel saved for it
(`test/dvedit.test.js` against `test/dv-excel.json`), less Excel's `xr:uid`, and each cell is judged as Excel judged
it. `dv-ours.mjs` writes five rules (one split by another, one with its messages and a Warning, a list without its
dropdown) and Excel read all 10 cells checked as written (`excel-dv-check.ps1`): types, comparisons, values, styles,
titles, messages and the dropdown.

A fix with this: rules covering several ranges (`sqref="A1:A3 A6:A10"`) were read as one. The reader split the
list on the letter s, not on spaces. Conditional formats had the same fault, also fixed.

## Circle Invalid Data (measured)

Data > Data Validation ▾ > Circle Invalid Data draws a red oval round each cell of the sheet shown whose value breaks
its rule; Clear Validation Circles takes them away (`validation.js invalidCells`, `editor/dvcircles.js`). Measured
with CircleInvalid in a hidden Excel, whose ovals are shapes that COM can read:

| Excel's rule | Excel's example |
|---|---|
| A cell is circled when its value breaks its rule, judged as typing is (above): formulas by their result, Information and Warning rules too. | 50, text, 2.5 and =40+2 under whole numbers 1 to 10; `pink` and `RED` under the typed list red,green,blue: 14 ovals in all. |
| A blank cell is circled when its rule doesn't ignore blanks. | C3:C6 blank under such a rule: four ovals. |
| At most 255. | 300 broken cells: 255 ovals. |
| Each oval is 3.75pt left of the cell and 1.5pt above it, 7.5pt wider and 3pt taller, in any row height or column width; cut off at the sheet's top and left (A1's at 0,0, that much narrower and shorter). | B2 (48, 15, 48 x 15pt): 44.25, 13.5, 55.5 x 18. A1: 0, 0, 51.75 x 16.5. |
| A merged area's oval is as wide as the area and as tall as its first row. | F6:G7, 96 x 30pt: 103.5 x 18. |
| Red (FF0000), 1.35pt, solid, no fill; not saved in the file. | The workbook saved with ovals shown has no drawing. |

The ovals stay with their sheet when another is shown. One goes when its cell is put right: that is Microsoft's
description of the command, not measured (a value changed over COM left Excel's oval where it was). Which 255 of
more are circled, and the message Excel may show then, are not measured either: here, rule by rule, row by row.

## Rules reading another sheet (Excel's 2010 form)

Measured by `tools/excel-compare/excel-dv-x14.ps1`: nine rules added over COM to B2:B5 of a new workbook with a sheet
Data and a name Choices, the workbook saved and its rules kept (`test/dv-x14-excel.json`).

| Excel's rule | Excel's example |
|---|---|
| A rule with a formula that reads another sheet is written in the sheet's extLst, not in `dataValidations`: an `x14:dataValidation` with the same attributes, its formulas as `x14:formula1` and `x14:formula2` holding an `xm:f`, its ranges as `xm:sqref`, in the ext `{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}`. | A list of `=Data!$A$1:$A$3`; whole numbers 1 to `=Data!$B$1` (its `1` in an `xm:f` too); greater than `=Data!$B$1`; a custom `=B2>Data!B1`; a text length equal to `=Data!$B$1`. |
| A rule of this sheet's references, or of a defined name (whatever it refers to), stays in the main form. | `=$H$1:$H$3`, `=Choices` (`Data!$A$1:$A$3`), `=B2>C2`. |
| Its ext comes after the conditional formats' and before the sparklines'. | A sheet with a sparkline: `{CCE6A557-...}`, then `{05C60535-...}`. |

The editor reads both forms (`sheet.js`, each rule's `ext` saying which), so a rule of the 2010 form judges typing,
shows its list (the other sheet's cells) and is circled by Circle Invalid Data. The Data Validation box, Clear All,
cut and paste take cells out of, move and add cells to rules of either form; a rule the box writes is put in the form
Excel would put it in (`dvedit.js`). The nine rules written here are Excel's XML, less its `xr:uid`
(`test/dvx14.test.js`); Excel opened the nine workbooks without a repair and read each rule back as given: type,
operator, formulas and B2:B5 (`excel-dv-x14.ps1 -Open test/out/dvx14`).

## Not yet

"Apply these changes to all other cells with the same settings", and a list's choices from a name defined as a
formula of another kind. Rules follow inserted and deleted rows, columns and cells as Excel moves them, split and
extends them, and stay where they are in a sort; Circle Invalid Data's ovals go after an insert or delete on their
sheet, as Excel's do (docs/ROWCOL.md).
