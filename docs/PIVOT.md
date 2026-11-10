# PivotTables: Refresh

Pivot tables are kept as they are when a workbook is saved, and their sources follow a sheet rename. They can also be
**refreshed**, as Excel's Refresh does (`src/ooxml/pivot.js`):

- **Alt+F5** refreshes the PivotTable at the active cell.
- **Ctrl+Alt+F5**, or **Data > Refresh All**, refreshes every PivotTable in the workbook.

Each refresh is one step that Ctrl+Z takes back.

## What a refresh does (measured)

A refresh reads the source range again (a sheet range, a defined name or a table). Then:

- **Cache:** each field's shared items are rebuilt in the order the values first come in the data. A number field
  that isn't on an axis gets its minimum and maximum again, and every record is written again. Pivot items match
  regardless of case, as Excel's do.
- **Item order:** each field keeps its items in the order they were shown. A value that is new to the field comes
  after them, even when it would sort first: "Central" is added after "South". A field sorted ascending or descending
  is sorted again.
- **Layout:** the pivot is laid out again in the compact form Excel 365 uses:
  - "Row Labels", with nested row items in one column, and each group's subtotal on its own row
  - the column field's items across, under "Sum of …" and "Column Labels"
  - several values side by side, their captions in the header row
  - with only a column field, the value's caption labels the data row
  - the grand total row and column
- **Values:** Sum, Count, Average, Max, Min, Product, Count Numbers, StdDev, StdDevp, Var and Varp.
- **What is rewritten:** the cells, the location, each field's items, and the row and column items. Everything else
  in the parts (styles, captions, extensions) stays as Excel wrote it.

## Measured against Excel (10 November 2026)

`tools/excel-compare/pivot-cases.mjs` makes eight cases from one table of made-up sales: region, product, quarter,
units and price. `excel-pivot.ps1` has a hidden Excel do each one:

1. build the pivot table at Pivot!A3, from its fields in rows, columns and values
2. read its cells and save the workbook
3. change the source (row 2's units to 100, its region to a new one, Central)
4. refresh it, then read and save again

The workbooks stay outside the repository, because Excel writes its user's name in the cache. `pivot-parts.mjs`
copies only the three pivot parts of each into `test/pivot-excel.json`, without `refreshedBy`, and refuses to write if
the workbook's author appears anywhere in them.

The eight cases:

- rows Region, sum of Units
- rows Region then Product
- rows Product, columns Quarter
- rows Region, sum of Units and average of Price
- count of Product
- max of Units and min of Price
- columns Region only
- rows Quarter then Region, columns Product

| Check | Result |
|---|---|
| Laid out again from Excel's own pivot, unchanged | 8 of 8: the same cells, and the same location, field items, row and column items, shared items and records as Excel's |
| Refreshed after the change | 8 of 8: the same cells as Excel's refresh, and the same rewritten XML |
| The editor's refreshed workbooks opened in Excel (`pivot-ours.mjs`, `excel-pivot-check.ps1`) | 8 of 8 opened, and Excel's own refresh of each changed nothing |

## Not refreshed yet

These pivots are left as they are, and the editor says why:

- report filters, hidden items, calculated fields or items, grouped fields (dates by month, numbers by range)
- a heading of the source renamed (Excel works that into the fields)
- the tabular and outline forms, grand totals or subtotals turned off
- more than one column field, several values with a column field, values on rows

A source with an empty heading gets Excel's "The PivotTable field name is not valid" message.

Also not handled:

- **Dates:** a date field is cached as numbers, where Excel caches dates.
- **Missing items:** an item no longer in the data is dropped from the field. That case wasn't measured.
- **Number formats:** a value field's number format isn't applied to cells the refresh newly fills.
- **Building and changing pivots:** creating a PivotTable, or moving its fields, can't be done here yet.
- **Pivot styles:** the PivotTable style's colours aren't drawn on the grid.

Tests:

- `test/pivot.test.js`: the 8 cases unchanged and refreshed, against Excel's cells and parts, and the refusals.
- `test/editing.mjs`: Alt+F5 after the source changes, Ctrl+Z, and Data > Refresh All.
- `test/a11y.mjs`: axe on the Data tab and its menu.
