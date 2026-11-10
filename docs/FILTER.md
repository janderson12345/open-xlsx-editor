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
(custom conditions, top 10, colours, date groups) is kept as Excel wrote it, and its rows stay as they were.

## Not yet

Text, number and date filters (custom conditions, top 10, above average), filtering by colour, date grouping in the
list, sorting while rows are filtered out (Excel sorts only the rows shown), the filter growing as rows are added
below, tables' own filters, and Advanced Filter.
