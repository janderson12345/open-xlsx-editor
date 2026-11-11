# PivotTables: Insert and Refresh

## Insert > PivotTable (measured)

**Insert > PivotTable** makes a new PivotTable (`src/ooxml/pivotnew.js`, the boxes in `src/editor/pivotbox.js`):

1. **PivotTable from table or range** (Excel 365's box): the Table/Range starts as Excel fills it. That is the table the
   active cell is in, by name, or the current region round it, as `Sheet!$A$1:$E$19`. Then New Worksheet (a sheet
   before the one shown, the pivot at A3, as Excel puts it) or Existing Worksheet and a Location.
2. **PivotTable Fields**: each field to Rows, Columns or Values, a value summarized by Sum, Count, Average, Max or Min.
   A field starts on Sum when it holds only numbers, Count otherwise, as Excel starts it. Rows and Columns take their
   fields in the list's order. Excel's field list is a pane where fields are dragged; here it is a box.

The parts are written as Excel writes a PivotTable it has just made, and laid out by the same code as a refresh:

- a cache over the source (a table by its name, otherwise the sheet and range), its shared items only for the fields
  on rows and columns, and its records
- the table in the compact form with PivotStyleLight16, each field's items sorted A to Z (as a new pivot's are, though
  the field isn't marked sorted), the values named "Sum of Units", "Count of Product" and so on
- the parts wired in as Excel wires them: the sheet to the table, the table to the cache, the workbook's
  `pivotCaches` to the cache by its id
- the row labels aligned left and the columns fitted, as Excel leaves a new one

**Measured:** each of the 8 cases below made here from the same data and fields gives Excel's own parts, byte for byte
(ids, the refreshed date and Excel 2026's data-source extension aside), and Excel's cells (`test/pivotnew.test.js`).
Saved and opened in Excel (`tools/excel-compare/pivotnew-ours.mjs`, `excel-pivot-check.ps1`): 8 of 8 opened, and
Excel's own refresh changed nothing.

**What it refuses:** a source whose headings have an empty one (Excel's "The PivotTable field name is not valid"), a
source of headings alone ("Reference isn't valid."), a pivot over another ("A PivotTable report cannot overlap another
PivotTable report."), and the layouts the refresh doesn't make (no row or column field, more than one column field,
several values with a column field). Over cells with something in them, Excel's question comes first; its wording is
from Excel's dialog, which COM doesn't reach.

**Undo:** a new PivotTable is one step that Ctrl+Z takes back, the new sheet it went on with it.

## Refresh

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
- **Changing pivots:** moving a PivotTable's fields once it is made can't be done here yet; the pivot's filter buttons
  (Excel's pivotButton on its header cells) aren't written on a new one.
- **Pivot styles:** the PivotTable style's colours aren't drawn on the grid.

Tests:

- `test/pivot.test.js`: the 8 cases unchanged and refreshed, against Excel's cells and parts, and the refusals.
- `test/editing.mjs`: Alt+F5 after the source changes, Ctrl+Z, and Data > Refresh All.
- `test/a11y.mjs`: axe on the Data tab and its menu, and on Insert > PivotTable's two boxes.
- `test/pivotnew.test.js`: the 8 cases made new, against Excel's parts and cells; saved, opened and refreshed; the
  refusals.
- `test/editing.mjs` also: Insert > PivotTable from a table typed in, Region to Rows and Units to Values, on a new sheet.
