# Conditional formatting

How conditional formats are drawn (`src/ooxml/condformat.js`), and what Excel was seen to do. Measured over COM: 29
cases, each a sheet given rules with `FormatConditions` in a hidden Excel, then every cell's fill, font colour and
bold read as Excel draws them (`Range.DisplayFormat`), and the workbook saved so its `conditionalFormatting` XML and
`dxfs` could be kept (`tools/excel-compare/excel-cf.ps1` on `cf-cases.json`, from `cf-cases.mjs`, giving
`test/cf-excel.json`). `npm test` (`test/condformat.test.js`) builds each sheet here with Excel's own XML for its
rules and expects Excel's look in every cell: 29 of 29.

## In the editor

The rules in a file are drawn: a rule's fill over the cell's own, its font colour, bold, italic and strikethrough
over the cell's. They're worked out again after every change, so a value typed in takes its format at once. The
rules and their formats are saved as they were.

## Excel's rules

| Rule | Seen |
|---|---|
| Rules go in priority order (1 first). A rule that holds gives the cell what its format sets and an earlier rule hasn't; Stop If True ends the cell's rules there. Rules added over COM have Stop If True set. | >10 red then >6 green and bold: 12 red, not bold (stopped); 7 green and bold. |
| Cell value rules compare as Excel compares: text after every number, case aside; a blank as 0 against a number (as "" against text); errors never hold. | Greater than 10: 12, 20, 15 and the text `x`. Not between 5 and 12: 20, 1, 15, the blank and `x`. Not equal 7: not the error. |
| Formula rules are relative to the first cell of the range, and hold when the formula is TRUE or a number other than 0. | `=A1>B1` on A1:A7 compares each row; `=$B1="y"` colours whole rows. |
| Text rules (contains, does not contain, begins with, ends with) are the formulas Excel writes for them, so case aside. | `NOT(ISERROR(SEARCH("apple",A1)))`: apple, Pineapple, APPLE pie, crab apple. |
| Blank and error rules are judged by the cell: an error is not blank, text of spaces is. | No blanks: the error cell too. |
| Top and bottom take the numbers of the range; a percentage is of their count, rounded down, at least one. | Top 25% of 7 numbers: only 20. Bottom 2: 1 and 5. |
| Above and below average take the numbers' mean; "or equal" includes it. | 1, 2, 3, 2: equal or above average 2, 3, 2. |
| Duplicates and unique values compare case aside, leave blanks out, and count an error as a value. | apple, apple, APPLE: duplicates. Unique: the error too. |
| Colour scales place each number between their points (lowest, a percentile, highest) and move each channel from the lower point's colour by whole steps, the fraction cut off. | 1 to 9 on red, yellow, green: 2 is #FA9D75 (248 + 2.8 is 250). |

## Adding rules

Home > Styles > Conditional Formatting has the menu's rules, each added for the selection:

- **Highlight Cells Rules:** Greater Than, Less Than, Between, Equal To, Text that Contains, A Date Occurring,
  Duplicate Values (or Unique).
- **Top/Bottom Rules:** Top 10 Items, Top 10%, Bottom 10 Items, Bottom 10%, Above Average, Below Average.
- **Color Scales:** Excel's default three-colour and two-colour scales.
- **Clear Rules** from Selected Cells (their cells taken out of every rule's range) or from Entire Sheet. A rule left
  without its first cell has its formulas written for its new first cell, as Excel's (*measured*: =E1>F1 over E1:E4,
  E1:E2's rules deleted, is =E3>F3 over E3:E4); before, the formula kept its text and so meant other cells.

Each opens Excel's small box: "Format cells that are GREATER THAN:", the values, and "with" one of the menu's preset
formats (Light Red Fill with Dark Red Text, Yellow Fill with Dark Yellow Text, Green Fill with Dark Green Text, Light
Red Fill, Red Text). Each rule added is one step of the undo history (`src/ooxml/cfedit.js`).

A rule is written as Excel writes it:

- **Its attributes** in Excel's order.
- **Values:** a number as it is, text quoted, "=" a formula.
- **Text that Contains** as `NOT(ISERROR(SEARCH("an",B1)))` from the range's first cell.
- **A Date Occurring** as `type="timePeriod"` with its period (yesterday, today, tomorrow, last7Days, lastWeek,
  thisWeek, nextWeek, lastMonth, thisMonth, nextMonth) and Excel's formula for it from the range's first cell:
  `FLOOR(A1,1)=TODAY()-1` for Yesterday, `AND(MONTH(A1)=MONTH(EDATE(TODAY(),0-1)),YEAR(A1)=YEAR(EDATE(TODAY(),0-1)))`
  for Last month, weeks from Sunday by WEEKDAY (`cfedit.js` CF_PERIODS). It's drawn by that formula, so against
  today's date, as Excel draws it.
- **Its format** as a dxf: the preset's font colour, and its fill as the pattern's background.
- **Placement:** in the conditionalFormatting of the same range, or a new one in its place among the sheet's elements.

Measured: `excel-cf-write.ps1` had Excel add each of 14 rules (`cf-write-cases.mjs`) with the menu's presets, and kept
the XML it saved (`test/cf-write-excel.json`). The same rules written here give the same rules, ranges and formats
(`test/cfedit.test.js`). Opened in Excel beside Excel's own (`cf-write-ours.mjs`, `excel-cf-look.ps1`), all 14 are
drawn the same: every cell's displayed fill and font colour.

A Date Occurring, measured apart (`excel-cf-dates.ps1`, `test/cf-dates-excel.json`): Excel added each of the ten
periods to a column of the dates from 45 days before today to 45 after, saved them, and gave each cell's displayed fill.
Written here, the ten rules are Excel's, and with the clock at Excel's day (15 November 2026, a Sunday) they highlight
the same 91 by 10 cells (`test/cfdates.test.js`). Not measured: other weekdays and month ends (the formulas are
Excel's own, worked out by the calculation engine).

Two things the measurement could not settle, because Excel's COM is not its menu. COM's rules Stop If True, and COM
gives a later rule the lower priority. The editor follows the menu: a new rule doesn't stop, and it comes first
(priority 1, as the Rules Manager lists the newest first). The tests compare everything else and set those two aside.

## Data bars and icon sets (`src/ooxml/cfext.js`, `cfbars.js`, `cfbarsedit.js`)

Home > Conditional Formatting > Data Bars (gradient or solid, Excel's six colours) and Icon Sets (Excel's twenty), and
New Rule's "Format all cells based on their values".

Measured over COM (`tools/excel-compare/excel-cf-bars.ps1`: `AddDatabar` set as the ribbon gives it, and
`AddIconSetCondition` with each set, on a new workbook; `test/cfbars-excel.json`): 33 of 33 rules written here as Excel
writes them, the rule and its extLst part text for text, Excel's made-up ids aside. Excel opened the workbook written
here (`excel-cf-bars-check.ps1`, `test/cfbars-readback-excel.json`) and read each column's rule back: its kind, a bar's
colour, fill, border and automatic ends, an icon set's set.

| Rule | As Excel writes it |
|---|---|
| A data bar | A `cfRule` of type dataBar (min, max, its colour) with an `x14:id`; in the worksheet's extLst the `x14:cfRule` of that id: min and max length 0 and 100, automatic min and max, negative bars red (`FFFF0000`), a black axis. Gradient: a border in the bar's colour, the negative border red. Solid: `gradient="0"`, no border. |
| An icon set | A `cfRule` of type iconSet with the set's name (none for 3 Traffic Lights, the default) and its thresholds in percent: 0, 33, 67; 0, 25, 50, 75; 0, 20, 40, 60, 80. |
| Excel 2010's sets | 3 Stars, 3 Triangles and 5 Boxes: an `x14:cfRule` in the extLst alone, with its own priority and `xm:f` thresholds. |
| The extLst | Its conditional formats ext comes first among the sheet's exts. Its rules' priorities count with the others. |

What they draw follows Excel's documented rules, not measured (COM can't read back a bar's length or a cell's icon).
- **Bar ends:** the min and max are the lowest and highest values; automatic ones are 0 when that is further out. They can also be a number, a percent of the span, or a percentile.
- **Bar length:** between the rule's min and max lengths.
- **Axis:** with negative numbers and an automatic axis, the axis falls where 0 does, and negative bars run left from it in the negative colour.
- **Icons:** each number gets the last icon whose threshold it reaches. Thresholds can be a percent of the span, a percentile or a number.

## The Rules Manager (`src/ooxml/cfmanage.js`, `src/editor/cfmanager.js`)

Manage Rules... lists the rules for the selection or the whole sheet in priority order: each with its description, the
cells it applies to (editable), and Stop If True. New Rule..., Edit Rule..., Delete Rule, Move Up and Move Down are each
one step of Undo. New Rule offers Excel's six rule types: based on values (colour scales, data bars, icon sets), cells
that contain a cell value (all eight comparisons), specific text or a date occurring, top or bottom, above or below
average, unique or duplicate, and a formula. Edit Rule (or a double-click on a rule) opens the same box as Edit
Formatting Rule, set to the rule, its Format "(the rule's own format)" unless another is picked.

Measured over COM (`tools/excel-compare/excel-cf-manage.ps1`, `test/cfmanage-excel.json`): Excel's four rules on a
sheet, each change made here, 5 of 5 as Excel makes them.

| Change | As Excel makes it |
|---|---|
| Delete | The rule goes, with its extLst part (the ext and extLst too when left empty); the others keep their priorities, gaps and all. |
| A rule given priority p | Every other rule from p on one later. Move Up gives the rule above's priority; Move Down moves the rule below up. |
| Sent last | One more than the last. |
| Stop If True | `stopIfTrue="1"` after the priority. |
| Applies to | The rule moves to a `conditionalFormatting` of its own just after the one it was in. |
| Edit Rule | The rule stays in its place in its `conditionalFormatting`, its ranges and its format (dxfId) kept, its type and criteria new: greater than 5 to between 3 and 7, or to a formula; a formula to less than 4; the middle of three rules. |

Edit Rule was measured apart (`tools/excel-compare/excel-cf-edit.ps1`, `test/cfedit-excel.json`, FormatCondition.Modify:
4 of 4 as Excel writes them, `test/cfeditrule.test.js`). COM's Modify also gives the rule one more than the last
priority and sets Stop If True, as its Add does; the Rules Manager keeps an edited rule where it is in its list, so here
it keeps its priority and its own Stop If True (not measured through the box itself, which COM doesn't reach). A new
format picked is a new dxf. A data bar or icon set, or a rule edited into one, is replaced: the new rule on the same
ranges, given the old one's place in the order.

A formula rule is `type="expression"` with its formula, as Excel writes it (its own rules from COM, `stopIfTrue`
aside, which COM always sets).

## Not yet

A rule's number format and borders, custom thresholds and colours in New Rule (and so a data bar's own colours kept
when one is edited), and the other presets of Color Scales.
