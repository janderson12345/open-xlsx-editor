# Typed entries

How what is typed into a cell becomes its value and its number format (`src/ooxml/input.js` parseInput and
formatAfterEntry, applied by `src/ooxml/entry.js` to typing, pasting text and Replace), and what Excel was seen to do.
Measured over COM with United States settings: 141 entries put into cells as typing would (`Range.Formula = text`),
some into cells formatted first (`tools/excel-compare/excel-entries.ps1` on `entry-cases.json`, giving
`test/entry-excel.json`). `npm test` (`test/entry.test.js`) reads every one as Excel did, value, type and the format
the cell gets. Typed through the editor and saved, 121 of the first 124 read back in Excel as Excel made them (the other
three: a boolean's spelling in the comparison, a format the test set up first, and the `1E+400` oddity below).

## In the editor

- Enter, the formula bar, pasting text from another program and Replace all read an entry this way.
- The formula bar and F2 show a cell as Excel's formula bar does: a date as m/d/yyyy (with its time as h:mm:ss
  AM/PM), a time alone as h:mm:ss AM/PM, a percentage as one. Typed back, each gives the same value.
- An entry that gives the cell a format is one step of the undo history with its value.

## Excel's rules

| Rule | Seen |
|---|---|
| A date with a year is shown m/d/yyyy: m/d/yyyy, m/d/yy, m-d-yyyy, yyyy-mm-dd, yyyy/m/d. A two-digit year 00 to 29 is 2000 to 2029, 30 to 99 is 1930 to 1999. | `1/2/24`, `2024-01-02`, `1-2-2024`: 45293. `1/2/29`: 2029; `1/2/30`: 1930. |
| A month and a day without a year are in the current year, shown d-mmm: m/d, `Jan 5`, `5 Jan`, `5-Jan`, `January 5`, `Sept 5`. Even `3/4` is the 4th of March, not three quarters. | `1/2`: 2-Jan this year. `Jan-24`: 24-Jan. |
| When the second number can't be that month's day it is a year, the date the 1st, shown mmm-yy. | `12/2024`: Dec-24. `Jan-32`: Jan-32 (1932). `2/29` in a year without one: Feb-29 (2029). `Jan 0`: Jan-00. |
| A day, a month name and a year are shown d-mmm-yy; a month name and a day want a comma before the year; a month name and a year are shown mmm-yy. | `5-Jan-2024`, `15-Mar-24`, `Jan 5, 2024`: d-mmm-yy. `Mar 15 24`: text. `January 2024`: mmm-yy. |
| A date that doesn't exist, or before 1900 or after 9999, stays text. | `13/1/2024`, `2/30/2024`, `0/1/2024`, `12/31/1899`, `1/1/10000`, `1.2.2024`, `Sat 1/6/2024`. |
| Times: h:mm is shown h:mm, h:mm:ss h:mm:ss, 24 hours or more [h]:mm:ss, seconds with a fraction mm:ss.0; minutes or seconds past 59 still make a number, but General. | `10:30`: 0.4375. `25:30`: [h]:mm:ss. `10:30.5`: ten minutes 30.5 seconds. `0:90`: 0.0625, General. |
| AM and PM want a space before them and an hour of 12 or less; then h:mm AM/PM. | `10:30 PM`, `9 AM`, `9:00 am`: h:mm AM/PM. `10:30pm`, `9AM`, `13:00 PM`: text. |
| A date and a time are shown m/d/yyyy h:mm (General past 24 hours); the time needs a colon or AM/PM. | `1/2/2024 10:30 PM`: m/d/yyyy h:mm. `1/2/2024 10`: text. |
| Currency, $ before the number (a sign or brackets around it allowed): with decimals Excel's format 8, without 6. Other currency signs give a plain number. | `$1,234.50`, `$.5`: format 8. `$5`, `-$5`, `$-5`, `($5)`: format 6. `€5`: 5, General. `5$`: text. |
| Thousands separators: a comma wants at least three digits after it; #,##0, with decimals #,##0.00. | `1,234`, `(1,234)`: #,##0. `1,2345`: 12345. `1,23`: text. |
| A percentage (% before or after) is 0%, with decimals 0.00%; an exponent 0.00E+00; a whole number and a fraction # ?/?. | `50%`, `%50`, `-50%`: 0%. `50.5%`: 0.00%. `1e3`: 0.00E+00. `1 1/2`, `0 3/4`: # ?/?. |
| Brackets make a number negative; spaces round an entry are dropped; `.5`, `5.` and `+5` are numbers. | `(5)`: -5, General. |
| A cell with a number format keeps it, except that a date typed into a percentage or a time-only format, or a time into a date-only one, takes the entry's format. A cell formatted as Text keeps what is typed as text. | `1/2/2024` into 0.00: 45293.00; into 0%: m/d/yyyy; into h:mm: m/d/yyyy; into d-mmm-yy: kept. `10:30` into m/d/yyyy: h:mm. Into @: text. |
| An error value typed as Excel names it is that error, in any case: #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A, #GETTING_DATA. With a space around it, or another name, it stays text. | `#n/a`: #N/A, an error (ISNA is TRUE). `#N/A `, ` #N/A`, `#N/A!`, `#NA`, `#SPILL!`, `#CALC!`: text. |
| Replace enters its result the same way. | 1/1/2024 with 1 replaced by 9: 9/9/2024. 50% with 50 replaced by 60: 60%. |

## Not yet

Other currencies (other countries' separators and date order are in docs/LOCALE.md), and `1E+400`: too big
for a number, Excel keeps it as text but gives the cell the scientific format; here it stays General. A fraction with
a denominator of 10 or more is given # ??/?? (not measured).
