# Opening a .csv

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

## Not yet

- The file is read as UTF-8 (a byte order mark left out); Excel reads a CSV without one in the system's code page, so
  text beyond ASCII can differ.
- Only commas separate fields (Excel uses the system's list separator); Text Import (Data > From Text/CSV) isn't here.
- Saving writes an .xlsx; Save As CSV isn't here.
