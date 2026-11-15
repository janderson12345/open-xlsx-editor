# Data tools

Fill > Series, Flash Fill, Remove Duplicates, Text to Columns (`src/ooxml/datatools.js`, `flashfill.js`) and Move or
Copy > Create a copy (`src/ooxml/sheetcopy.js`); their boxes in `src/editor/datatoolboxes.js`. Each is one step of
Undo.

Measured over COM (`tools/excel-compare/excel-datatools.ps1`: 29 cases of `Range.DataSeries`, `RemoveDuplicates`,
`TextToColumns` and `FlashFill`, each on a new workbook of made-up cells, and a sheet copied with `Worksheet.Copy`;
`test/datatools-excel.json`). The same cells here and the same tool: every cell of A1:H12 as Excel left it, 29 of 29.

## Fill > Series (Home > Editing > Fill)

| Rule | Seen |
|---|---|
| Each row (or column) of the selection goes on from its first cell: + step (Linear), x step (Growth). | 1 by 2: 1, 3 ... 11. 2 growth 3: 2, 6, 18, 54, 162. |
| A stop value ends it there; the cells past it are left. | 5 by 3, stop 15: 5, 8, 11, 14. |
| Months and years from the first date, the day kept or the month's last. | Jan 31 by months: Feb 28, Mar 31, Apr 30. Feb 29 2024 by years: Feb 28 2025, Feb 28 2026. |
| Weekdays skip Saturday and Sunday. | Fri Nov 13: Mon 16, Tue 17 ... |
| Trend fits a straight line (Growth: a growth curve) to the values there and writes it over them too. | 1, 3, 4: 1.1667, 2.6667, 4.1667, 5.6667 ... |
| A date's format goes with it. | |

## Flash Fill (Data > Flash Fill, Ctrl+E)

The column's typed examples read as a pattern of the cells to their left (the table round the active cell), carried
down the rest. A pattern is pieces in a row: a whole cell, a word, a word's first letter, characters at fixed places,
or fixed text, each in its own case. The pattern that makes every example and leans least on fixed text and places
is used.

| Example | Excel filled |
|---|---|
| Avery Chen: Avery | Sam, Jo, Rin |
| Avery Chen: A. Chen | S. Lee, J. Park |
| ab-12: AB-12 | CD-34, EF-56 |
| Avery, Chen: Chen, Avery | Lee, Sam; Park, Jo |
| INV-0042-X: 0042 (typed, so 42) | 7, 0: the two characters where 42 was, made numbers |

Excel's Flash Fill learns from more than these five patterns; others here may differ from Excel's.

## Remove Duplicates (Data > Remove Duplicates)

| Rule | Seen |
|---|---|
| Rows alike in the columns ticked go, but the first; text is compared case aside, numbers by value. | pear/red and Pear/red: one. 1 and 1.0 typed: one. |
| The rows under move up within the range; cells beside it stay. | D2 kept its place. |
| Excel's message: "n duplicate values found and removed; k unique values remain." | |

## Text to Columns (Data > Text to Columns)

| Rule | Seen |
|---|---|
| Each cell split at the delimiters ticked (tab, semicolon, comma, space, another); quotes keep a piece whole. | "Smith, John",42: Smith, John and 42. |
| Treat consecutive delimiters as one; an empty piece leaves its cell. | one,,three: one, (nothing), three. |
| Each piece read as typed: numbers, dates (with their format), TRUE, percentages; 007 is 7. | 007, 3/4/2026, TRUE, -5%. |
| The destination may be elsewhere; the source then stays. | |
| Fixed width: each cell cut at the column breaks (character places from the left), each piece the characters between two breaks with the spaces at either end trimmed; an empty piece leaves its cell; a break past a cell's end gives it nothing there. | `ab   cd  e` at 3 and 8: ab, cd, e. `XYZ0000999Y`: XYZ, 9, 99Y. |
| Fixed width splits a number's cell too, from the text it shows; each piece read as typed, as Delimited's. | 1234567 at 3: 123 and 4567. `2026-01-05`: a date, m/d/yyyy. |
| Column data format Text: the piece kept as text, in the @ format. | 007 and 0042 stay as typed; 12 too. |
| Column data format Do not import column: left out, the columns after it moved left. | a,3/4/2026,gone,12.50 with the third skipped: four pieces in three columns. |
| Column data format Date (MDY, DMY, YMD, MYD, DYM, YDM): the piece's three numbers taken in that order, apart by / - or ., a two-digit year as a typed one (00 to 29 the 2000s); the m/d/yyyy format. A piece that isn't a date that way is read as typed. | DMY: 3/4/2026 is 3 April, 13/4/2026 13 April. YMD: 26-3-4 is 4 March 2026. MDY: 12/31/99 is 1999; 2026-03-04 is 4 March. |

The box asks for the Original data type (Delimited or Fixed width), as Excel's wizard does. Fixed width takes the
column breaks typed as character places (Excel's wizard has a ruler to click), over a preview of the first five cells
with a | at each break. Column data format is a list for each column the first five cells split into (Excel's wizard
picks a column in its preview). Not here: Excel's own guess at the breaks, and Advanced's decimal and thousands
separators.

## Move or Copy > Create a copy (a sheet tab's menu)

| Rule | Seen |
|---|---|
| The copy is named "Data (2)" (the first number free), before the sheet chosen. | |
| Its part is the sheet's text but for a new `xr:uid`, its tab not selected. | Excel's sheet2.xml against sheet1.xml. |
| The sheet's local names are copied as the copy's; a workbook-wide name referring to the sheet gets a copy local to the copy, referring to it. | Everywhere (=Data!$A$1): also Everywhere on Data (2), ='Data (2)'!$A$1. |

Notes, their VML boxes (a block of shape ids of their own), drawings and charts are copied with it, the charts
referring to the copy. A sheet with tables, PivotTables, threaded comments or controls isn't copied yet: they need
names and ids of their own.
