# Filter

How AutoFilter works (`src/ooxml/filterops.js`, the buttons in `src/editor/filterbuttons.js`), and what Excel was seen
to do. Measured over COM: 13 cases of `Range.AutoFilter` on a new workbook in a hidden instance, each then saved by
Excel so that its own XML could be read (`tools/excel-compare/excel-filter.ps1` on `filter-cases.json`, from
`filter-cases.mjs`, giving `test/filter-excel.json`). `npm test` (`test/filter.test.js`) does the same steps here and
expects Excel's hidden rows, its filter range, its `autoFilter` element (all but its random `xr:uid`) and its
`_xlnm._FilterDatabase` name: 13 of 13. Saved by the editor and opened in Excel, 13 of 13 read back with Excel's
filter, range and hidden rows.

## In the editor

- Home > Editing > Sort & Filter > Filter (Ctrl+Shift+L) turns the filter on over the data around the active cell
  (what Sort would take), or off. Clear takes every column's filter off; Reapply filters again after edits.
- A button in each heading cell opens the column's list (Alt+Down on the heading cell from the keyboard): Sort A to Z
  and Z to A (Smallest to Largest for numbers), Clear Filter From the column, a search box, and the column's values to
  tick, with (Select All) and (Blanks). OK keeps the rows ticked; a filtered column's button is marked.
- Each change is one step of the undo history. Not on a read-only workbook; `features: { editing: false }` turns it
  off.

## Excel's rules

| Rule | Seen |
|---|---|
| Turned on from one cell, the filter covers the data around it (Excel's CurrentRegion). | B3 in a table A1:C7: `<autoFilter ref="A1:C7"/>`. |
| A row is kept when its cell's shown text is one of the values ticked, case aside; blanks only when (Blanks) is ticked. Other rows are hidden (`hidden="1"` on the row). | `apple` keeps apple and Apple. `1` and `10` keep 1 and 10, not 2.5. Formatted 0.00, the values are `1.00` and `2.50`. |
| The values are written in order, blanks as an attribute. | `<filters><filter val="green"/><filter val="red"/></filters>`; `<filters blank="1"><filter val="blue"/></filters>`. |
| Each filtered column has its own `filterColumn` (`colId` from the range's first column); a row must pass all. | Team red or blue and Qty 1: only row 2 shown. |
| Clearing a column's filter takes its `filterColumn` away, the others stay. | |
| Excel keeps its range in a hidden name local to the sheet, written when the filter is on and left when it's off. | `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$C$7</definedName>` |
| Turning the filter off takes the element away and shows every row of its range. | |

The list itself (its order, one entry for texts alike but for case, values only from rows the other columns let
through) and Clear Filter From follow Excel's interface, which COM doesn't reach. A filter of a kind not read here
(icons, date groups) is kept as Excel wrote it, and its rows stay as they were.

## Text, Number, Date and Color filters (`src/ooxml/filtermore.js`, `src/editor/filterbox.js`)

A column's list has Text Filters..., Number Filters... or Date Filters... (by what the column holds): Excel's Custom
AutoFilter with two conditions (And or Or: equals, does not equal, begins and ends with, contains and does not contain,
greater and less than; * and ? wildcards), and the choices that take no value. Those are Top 10 (items or percent, top
or bottom) and Above or Below Average for numbers; today, this week, last month, this quarter, this year and the like
for dates; and the column's fill and font colours.

Measured over COM (`tools/excel-compare/excel-filter-more.ps1`: 31 filters of `Range.AutoFilter` on the same made-up table,
its dates days either side of =TODAY(), Excel's day recorded; `test/filter-more-excel.json`). Here the same table and
day, the same filters: 31 of 31 hide Excel's rows and write Excel's autoFilter, and the colour filters Excel's dxf.

| Filter | As Excel writes and applies it |
|---|---|
| Equals a value without wildcards | A list of that value, as a ticked value is. |
| Other conditions | `customFilters` (`and="1"` for And) of `customFilter`s: `operator` (none for equals) and `val`, wildcards kept. Numbers compare with numbers, text with text by its letters, case aside; a blank never passes a comparison. |
| Top 10 | `top10` (`top="0"` for bottom, `percent="1"`), `val` the number asked for, `filterVal` the last value kept: n% of the count, at least one (50% of 7: 3); ties kept. |
| Above and below average | `dynamicFilter` with the average as `val`. |
| Dates | `dynamicFilter` with `val` and `maxVal` the days the span covers, from today (a week from Sunday). Applied again, they span from the day it is. |
| By colour | `colorFilter` pointing to a dxf of the colour as a solid fill's foreground, white behind; `cellColor="0"` for a font colour. |

## Advanced Filter (`src/ooxml/advfilter.js`, `src/editor/advfilterbox.js`)

Data > Advanced...: filter the list in place or copy to another location, with a criteria range and Unique records
only. Data > Clear shows the rows an in-place Advanced Filter hid.

Measured over COM (`tools/excel-compare/excel-advfilter.ps1`: 7 cases of `Range.AdvancedFilter`; `test/advfilter-excel.json`):
7 of 7 hide or copy Excel's rows and write Excel's names.

| Rule | Seen |
|---|---|
| The criteria's headings name the list's columns; a row's cells must all hold (And), any row may (Or); a blank cell holds for anything. | Team red and Qty >2, or Team blue: apple and fig hidden. |
| Text alone means "begins with", case aside; ="=x" the value x exactly; >, >=, <, <=, <> compare. | a: apple, Apple, avocado. =pear: both pears. |
| Unique records only drops a row the same as one before it. | The second pear, 2, blue hidden. |
| Copied: the heading and the rows passing, at the cell given; nothing hidden. | |
| Names: `_xlnm._FilterDatabase` (hidden) for the list, `_xlnm.Criteria` for the criteria, `_xlnm.Extract` for the copy's heading row. | Extract: `Sheet1!$I$1:$K$1`. |

## Not yet

Date grouping in the list, sorting while rows are filtered out (Excel sorts only the rows shown), the filter growing as
rows are added below, tables' own filters, and filtering by an icon.
