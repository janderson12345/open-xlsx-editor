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

## Not yet

Filter (AutoFilter's buttons; autofilters in files are kept), sorting by cell colour, font colour or icon, custom
lists as an order, Excel's question about numbers stored as text (sorted as text here, as without it), sorting a
table (ListObject) or within a filter, the Data tab's Sort & Filter group, the cells' menu's Sort, and the sort
state Excel saves with the sheet. Notes, hyperlinks, conditional formats and data validation stay where they are.
