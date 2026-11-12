# Opening and saving a .csv

A .csv opens as Excel opens one (`src/ooxml/csv.js`): one sheet named after the file, each field taken as if it were
typed into its cell. Measured over COM: Workbooks.Open of test/csv-excel.csv (made-up fields), each cell's value,
number format and formula read back; 32 of 32 cells the same here (test/csv.test.js), and the editor shows them
(test/editing-formats.mjs).

## Excel's rules

| Field | Excel's cell |
|---|---|
| `42`, `007`, ` 7 `, `-12`, `(12)` | Numbers: 42, 7, 7, -12, -12; General |
| `3.14`, `1e5`, `12345678901234567890` | 3.14; 100000 shown 1.00E+05 (`0.00E+00`); 1.23E+19 |
| `50%`, `$1234.50`, `"1,234"` | 0.5 as `0%`; 1234.5 as currency; 1234 as `#,##0` |
| `2024-03-05`, `3/5/2024`, `13:45` | Dates as `m/d/yyyy`, a time as `h:mm` |
| `1/2`, `1-2`, `Mar-24` | Dates as `d-mmm` (the 2nd of January, the 24th of March, this year), as typed ones are |
| `TRUE`, `false`, `#N/A`, `#DIV/0!` | Booleans and errors |
| `=1+2` | A formula, worked out |
| `"quoted, comma"`, `"say ""hi"""`, a line break in quotes | Text: the comma kept, "" a quote, the line break kept (CR LF) |
| `  padded  `, `1 234`, `0001.50` | Text with its spaces; text; 1.5 |
| An empty field | No cell |

A field Excel wouldn't take as a formula is the text itself.

## Saving a sheet as a CSV

`saveCsv({ utf8 })` gives the sheet shown as Excel's Save As writes it (`csvBytes` in `src/ooxml/csv.js`), byte for
byte: test/csv-excel-utf8.csv (CSV UTF-8) and test/csv-excel-ansi.csv (CSV) are Excel's own saves of a sheet made
the same way in test/csv.test.js.

| Rule | Seen |
|---|---|
| Each cell as it shows, formatted. | `"1,234.50"`, `25%`, `3/5/2024`, `(5.00)`, `TRUE`, `#N/A`; a formula its value |
| A General number in 11 characters, whatever its column's width (no #s). | `0.333333333`, `1.23457E+12`, `0.000001`; 12345 in a column 3 wide |
| Text as it is; a typed apostrophe isn't written. | `'007`: `007`; `  lead` keeps its spaces |
| In quotes only with a comma, a quote or a line break; quotes doubled. | `"has, comma"`, `"say ""hi"""`, `"quote""mid"`; `semi;colon` bare |
| Rows end in CR LF; a line break in a cell is written as it is (LF). | `"two` LF `lines"` CR LF |
| The used range: each row as wide as it, empty cells empty fields. | `,x,,` |
| CSV UTF-8: a byte order mark, then UTF-8. CSV: code page 1252. | `café ü €`: `EF BB BF ... C3 A9`; `E9 FC 80` |

## Not yet

- The file is read as UTF-8 (a byte order mark left out); Excel reads a CSV without one in the system's code page, so
  text beyond ASCII can differ.
- Only commas separate fields (Excel uses the system's list separator); Text Import (Data > From Text/CSV) isn't here.
- `save()` still writes an .xlsx (a file opened from a CSV too); `saveCsv()` is the host's to call, and there is no
  File > Save As box. A character code page 1252 hasn't is written as ?, as Excel's CSV does (not measured).
