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

## Encoding, separator and locale

A CSV's bytes are read as UTF-16 or UTF-8 when a byte order mark says so, as UTF-8 when they are valid UTF-8, else as
Windows-1252 (Excel reads a file without a mark in the computer's code page, 1252 here; valid UTF-8 is read as such
here, so `café` in a UTF-8 file without a mark comes out right). The field separator is the one that splits most lines
into the same number of fields among comma, semicolon, tab and pipe, the locale's list separator on a tie (Excel takes
the computer's list separator only). Numbers and dates are read in the editor's locale (docs/LOCALE.md): in Germany
`1.234,5` is 1234.5 and `05.03.2024` the 5th of March. `csvWorkbook(bytes, name, today, { locale, sep, encoding, quote })`
takes each of them given.

## Text Import

The Text Import box (Excel's Text Import Wizard in one box; Data > From Text/CSV, and opening a .txt) reads a file with
settings chosen (`src/ooxml/textimport.js`). Measured over COM: Workbooks.OpenText of eleven made-up files
(`test/textimport-cases.js`, `tools/excel-compare/excel-textimport.ps1`, `test/textimport-excel.json`), each with its
settings; 108 of 108 cells the same value, type and number format here (`test/textimport.test.js`).

| Setting | Excel's rule, as seen |
|---|---|
| Encoding: UTF-8, UTF-16 (with its mark), Windows-1252 | `café`, `日本` in UTF-16; `€5` in 1252 |
| Delimited: tab, semicolon, comma, space, another character, several at once | a pipe as Other splits at each pipe |
| Text qualifier: `"`, `'` or none; doubled inside a field for one | `'it''s'`: it's; with `'`, `"q"` keeps its quotes; with none, `"b,c"` is `"b` and `c"` |
| Consecutive delimiters as one | `one   two` two fields; a delimiter first still makes an empty first field |
| Fixed width, columns starting where the breaks are | each field without the spaces around it: `Widget  ` is Widget |
| Start import at row | rows before it left out |
| General column | as if typed (`=1+2` a formula, `TRUE`, `#N/A`, `1e3`, `$12`, `50%`, `10:30`, `1/2`), with the decimal and thousands separators given: `1.234,5` is 1234.5 `#,##0.00`, `3,5%` 0.035 `0.00%`, `1.5` text; `€5` gets a euro format, `[$€-2] #,##0` (typed, it is a plain 5) |
| Text column | as it is, formatted `@` (`007`, the header too) |
| Date column, MDY, DMY or YMD | numbers split by `/`, `-` or `.`; a four-digit year first read year first in any order; m/d/yyyy, or d-mmm when there is no year (this year); a day the month hasn't (`31/2/2024`), another order (`13/5/2024` in MDY) or a word stays text |
| Skip | the column left out, the next one moving left |

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

- `save()` still writes an .xlsx (a file opened from a CSV too); `saveCsv()` is the host's to call, and there is no
  File > Save As box. A character code page 1252 hasn't is written as ?, as Excel's CSV does (not measured).
