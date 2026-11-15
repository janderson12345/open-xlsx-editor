# Sort

How cells are sorted (`src/ooxml/sortops.js`), and what Excel was seen to do. Measured over COM: 60 cases on a new
workbook in a hidden instance (`tools/excel-compare/excel-steps.ps1` on `sort-cases.json`, from `sort-cases.mjs`,
giving `test/sort-excel.json`), sorted with `Worksheet.Sort` as Excel's Sort box applies one, and with `Range.Sort`
and `xlGuess` where the headers are guessed (the Sort object's own guess was seen never to take a header; Range.Sort's
does, as the A to Z button does). `npm test` (`test/sort.test.js`) sorts the same cells and expects Excel's in every
cell (formula and value, number format, bold, fill, bottom border), the merged ranges, a formula on another sheet and a
defined name, or Excel's refusal in its words: 60 of 60. Sorted by the editor, saved and opened in Excel
(`excel-readback.ps1`), 60 of 60 read back as Excel made them.

## In the editor

- Home > Editing > Sort & Filter: Sort A to Z, Sort Z to A, Custom Sort....
- Sort A to Z and Z to A sort by the active cell's column. One cell: the data around it (Excel's CurrentRegion: grown
  while any filled cell touches it, side or corner). Several cells with data beside them (the region around them is
  wider): Excel's Sort Warning first, Expand the selection chosen, or Continue with the current selection.
- Custom Sort... opens Excel's Sort box over the same range: levels (Sort by, Then by: a column and an order, A to Z
  or Z to A for text, Smallest to Largest for numbers, Oldest to Newest for dates), Add Level, Delete Level, Copy
  Level, Move Up, Move Down; My data has headers (ticked as Excel guesses; the columns then named by their headers);
  Options...: Case sensitive, Sort top to bottom or left to right.
- One step of the undo history; the selection stays. Not on a read-only workbook; `features: { editing: false }`
  turns them off.

## Excel's rules

| Rule | Seen |
|---|---|
| Ascending: numbers (dates and times are numbers), then text, then FALSE, then TRUE, then errors (all alike). Blank cells go last, ascending or descending; a formula giving "" is text, before all other text. Descending is the same order turned round, blanks still last. | -1, 0, 2.5, 3, "", " ", "10", a, b, B, FALSE, TRUE, #N/A, #DIV/0!, then the blank. |
| Ties keep the order they were in, either way. | b before B (row 3 before 10), ascending and descending. |
| Text: case left aside unless Case sensitive is ticked; then small letters first, compared from the left. | a, a, A, A, ab, aB, Ab, AB, b, B. |
| Signs before digits before letters; digits one by one; the signs in Excel's order (space, ! " # $ % & ( ) * , . / : ; ? @ [ \ ] ^ _ \` { \| } ~ ’ + < = > § °), not their codes. | " a", !x, #x, (x), [x], _x, ~x, 10, 1a, 2, a. x.y before x1. |
| Apostrophes, hyphens and en dashes are passed over, then break ties after the word without (' before - before –). A non-breaking space is a space, after a plain one. | coop, co-op, cop. x, 'x, -x, –x. |
| Accents after the plain letter; æ as ae, ß as ss. | e, é, f. n, ñ, o. |
| Numbers stored as text sort as text, unless the key says to sort them as numbers. | As numbers: -3, 1, "1.5", "2", 9, "10", x. |
| Each cell moves with its format (bold, fill, number format), but borders stay where they were. | A bottom border under C4 stayed under C4. |
| A formula moves as a copy moves it: relative references follow the row, absolute ones stay. Formulas outside the range that read it, on the sheet, another sheet or in a name, aren't changed. | =B3*2 moved to row 2 is =B2*2; =B4 is =B3; =$B$2 stays; H2 =B2 and Sheet2's =Sheet1!B2 still read B2. |
| Several keys: the first decides, the next breaks its ties. | Team A to Z then Score largest first: blue 7, blue 7, green 1, red 9, red 5, red 5. |
| Left to right: columns ordered by a row. | 1, 3, a, b along row 1, the row under following. |
| Headers guessed only when the first row has fewer than two blank cells, and some column has text over a number, or a cell of the first row is formatted otherwise than the one under it (bold, fill, number format, border, either way). | Qty over 3: headers. 5 over 3, m over c, 7 over c, Name over a blank: none. Bold m over c: headers. Two blank columns: none, even with Qty over 3 or bold. |
| Merged cells: refused unless every cell of the range is in a merged range of one size; then each sorts whole. | "To do this, all the merged cells need to be the same size." A2:B2, A3:B3, A4:B4 sorted. |
| An array formula in the range: refused. | "You can't change part of an array." |
| One empty cell with nothing around it: refused. | "This can't be applied to the selected range. Select a single cell in a range and try again." |
| Sort On Cell Color (Font Color): the rows whose cell has that fill (font colour) On Top, or On Bottom, each group keeping its order; further levels within. | Yellow on top: fig, plum (yellow), then pear, kiwi, apple, date. Yellow, then green, then A to Z: fig, plum, apple, date, kiwi, pear. |
| A custom list as the order: its words first, in its order whatever their capitals, the values not in it after (before, descending), blanks last. | High, Medium, Low over Low, High, None, Medium, high, blank, Low: High, high, Medium, Low, Low, None. Descending: None, Low, Low, Medium, High, high. |

Sort On and custom lists were measured apart (`sort2-cases.mjs`, `test/sort2-excel.json`, SortFields.Add's SortOn,
SortOnValue.Color and CustomOrder: 6 of 6 cases). In the Sort box, each level's Sort On is Cell Values, Cell Color or
Font Color (a colour that column has, On Top or On Bottom); Order's Custom List... takes the list typed, "High, Medium,
Low" (Excel's box picks from its own lists).

## The sort saved with the sheet

Excel saves the last sort of a sheet as its `sortState`, and so does the editor after every sort (the Sort box, the A
to Z and Z to A buttons, a filter's), as one step of Undo with the sort (`src/ooxml/sortstate.js`). Measured over COM
on 17 sorts of a made-up table, saved and read back (`tools/excel-compare/excel-sortstate.ps1`,
`sortstate-cases.json`, `test/sortstate-excel.json`; `test/sortstate.test.js` makes each and compares: 17 of 17):

| Rule | Excel's file |
|---|---|
| The ref is the rows sorted, without the header row; with Left to Right, the columns. | A1:C6 with headers: `<sortState ref="A2:C6">` |
| One sortCondition per level, its ref the key column's data rows (the key Excel's Sort box passes; given one cell over COM, Excel writes that cell). | `<sortCondition descending="1" ref="C2:C6"/><sortCondition ref="A2:A6"/>` |
| Z to A: `descending="1"`. Match case: `caseSensitive="1"`. Left to Right: `columnSort="1"`, the key a row. | `<sortState columnSort="1" ref="B2:C6"><sortCondition ref="B2:C2"/>` |
| Sort On Cell Color or Font Color: `sortBy` and the colour as a new dxf (a solid fill over `indexed="65"`; or the body font with every property given and the colour). | `<sortCondition sortBy="cellColor" ref="B2:B6" dxfId="0"/>` |
| A custom list: `customList`, its entries joined by commas. | `<sortCondition ref="C2:C6" customList="red,green,brown,purple"/>` |
| A filtered range's sort: the one sortState is the autoFilter's last child; otherwise it follows the autoFilter in the sheet. A new sort replaces the last. | `<autoFilter ref="A1:C6"><sortState ref="A2:C6">...` |
| The A to Z and Z to A buttons save the same form as the Sort box. Excel 365 declares the xlrd2 namespace on it (written so here). | |

A sort inside a table is not saved here: Excel keeps that sortState in the table's own part.

## Not yet

Sorting by a conditional format's icon, Excel's question about numbers stored as text
(sorted as text here, as without it), sorting a table (ListObject), the cells' menu's Sort, a table's saved sort, and
moving a saved sort's ranges when rows or columns are inserted or deleted. (Within a filter only the rows shown are sorted, as Excel does: docs/FILTER.md.) Notes, hyperlinks, conditional formats and data validation stay where they are.
(Filter, and the Data tab's Sort & Filter group, are done: docs/FILTER.md.)
