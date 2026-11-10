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

## Not yet

Data bars and icon sets (stored in the sheet's extLst by Excel 2010 and later), date rules ("a date occurring")
checked against today's date, a rule's number format and borders, and the Conditional Formatting menu to add, edit and
manage rules.
