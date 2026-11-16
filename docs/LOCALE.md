# Other countries' settings

Excel takes the decimal and thousands separators, the order of day, month and year, the short date and the list
separator from the computer's Region settings; the file is the same whatever they are. Here they come from the
editor's `locale` option (a tag such as `'de-DE'`, or an object changing settings over one:
`{ tag: 'en-US', dateOrder: 'DMY' }`, `{ decimal: ',', group: '.' }`), else the browser's `navigator.language`, and
`editor.setLocale()` changes them (`src/ooxml/locale.js`). Like Excel's, the setting is the page's: every editor on a
page shows and reads typing in the last one set.

## What follows the locale

| What | en-US | de-DE | en-GB | fr-FR |
|---|---|---|---|---|
| A number shown (`#,##0.00`, General) | 1,234.50, 1234.5 | 1.234,50, 1234,5 | 1,234.50 | 1 234,50 |
| Built-in format 14 (and 22) | 3/5/2024 | 05.03.2024 | 05/03/2024 | 05/03/2024 |
| Typed `1.234,5` / `1,5` | text / text | 1234.5 (#,##0.00) / 1.5 | text / text | text / 1.5 |
| Typed `13/5/2024` | text | 13 May 2024 | 13 May 2024 | 13 May 2024 |
| Typed `5.3.2024`, `1.5` | text, 1.5 | 5 March 2024, 1 May | text, 1.5 | text, text |
| A formula typed and shown | `=ROUND(1.5,0)` | `=ROUND(1,5;0)` | `=ROUND(1.5,0)` | `=ROUND(1,5;0)` |
| An array constant | `{1,2;3,4}` | `{1.2;3.4}` | `{1,2;3,4}` | `{1.2;3.4}` |
| The formula bar's date and time | 3/5/2024 6:00:00 PM | 05.03.2024 18:00:00 | 05/03/2024 18:00:00 | 05/03/2024 18:00:00 |

The file keeps en-US: `ROUND(1.5,0)`, 1234.5, numFmtId 14. The calculation engine keeps en-US too (TEXT, VALUE), as
the file's cached values have to be the same wherever it is opened.

## Measured

- **Display with other separators** (`tools/excel-compare/excel-separators.ps1`, `test/separators-excel.json`): Excel's
  own "Use system separators" turned off and set to `,` and `.`, `,` and a no-break space, `.` and `'`, 13 formats ×
  8 values read back as Range.Text (the probe puts the settings back). 300 of 300 cells that show a value are the same
  here (`test/locale.test.js`; the other 12 are #s, a negative date or one past 9999). Rules: every digit
  separator in a number format follows the locale, in General and in scientific, percentage, fraction and time
  fraction (`mm:ss,0`) formats alike; literals in quotes don't (`"v1.5 "0.0` is `v1.5 1234,5`); a no-break space as
  thousands separator is shown as a plain space; dates in a custom code (`m/d/yyyy` given as a code, not format 14)
  are as the code says.
- **Typing** with other separators follows Excel's en-US rules with the separators swapped (a group wants three digits
  after it, a percentage, an exponent, currency); the date order and a dot between a date's parts are the Region
  settings', which can't be changed for Excel alone, so they are not measured. They follow Text Import's measured
  date columns (docs/CSV.md): a day that month hasn't is text, a four-digit year first is year, month, day in any
  order, a date without a year is this year's, shown d-mmm.

## Not yet

- Function names stay English (German Excel shows `RUNDEN`), as do month and day names, TRUE and FALSE, and the error
  values. Currency formats 5 to 8 and 41 to 44 stay in dollars.
- Boxes that read a typed value (filters, data validation, conditional formatting rules) read it en-US.
- Two editors on one page can't have different locales.
