# Outlines: Group, Ungroup, Subtotal

How rows and columns are grouped into an outline, collapsed and expanded, and how Data > Subtotal makes one
(`src/ooxml/outline.js`, `src/ooxml/subtotal.js`), and what Excel was seen to do. Each rule was measured over COM in a
hidden Excel (`tools/excel-compare/excel-outline.ps1`: `Range.Group`, `Ungroup`, `Outline.ShowLevels`,
`Range.ShowDetail`, `ClearOutline`, `AutoOutline`, `Outline.SummaryRow` and `SummaryColumn`, `Range.Subtotal`,
`RemoveSubtotal`, each on a new workbook of made-up cells, saved and read back). `npm test` (`test/outline.test.js`, 38
cases; `test/subtotal.test.js`, 11 cases) expects what Excel wrote, to the attribute; `test/outline.mjs` clicks the
outline bar's buttons and the Subtotal box in a browser.

## In the editor

- Data > Outline: Group (Shift+Alt+Right) and its menu's Auto Outline; Ungroup (Shift+Alt+Left) and its menu's Clear
  Outline; Subtotal; Show Detail; Hide Detail; Settings (Summary rows below detail, Summary columns to right of detail).
- Whole rows or columns selected are grouped at once; any other range asks Rows or Columns, as Excel's Group box does.
- Show Detail, Hide Detail and Clear Outline go by the selection: whole columns, the columns; whole rows, the rows;
  otherwise the rows when the active cell's row is in a row group, the columns when not. Clear Outline with one cell
  selected clears the sheet's whole outline.
- The outline bar: left of the row headings and over the column headings, the level buttons 1, 2, 3... in the corner, and
  for each group a line along its rows (columns) and a − at its summary row, a + while its rows are hidden.
- Ctrl+8 hides the outline symbols (the bar goes, the outline stays) and shows them again: a view setting saved with the
  file, as Gridlines is, so not a step of the history.
- Every other action is one step of the undo history, and every one is a named command (`group`, `ungroup`,
  `outlineLevel`, `outlineToggle`, `showDetail`, `hideDetail`, `clearOutline`, `autoOutline`, `outlineSettings`,
  `outlineSymbols`, `subtotal`, `removeSubtotals`).
- Rows and columns an outline hides are hidden lines like any other: the status bar, filters, SUBTOTAL(10x), Unhide
  and printing treat them as hidden.

## Excel's rules: grouping

| Rule | Seen |
|---|---|
| Grouping gives each line one more `outlineLevel`; sheetFormatPr's `outlineLevelRow` (`outlineLevelCol`) holds the deepest. Nothing else changes, hidden lines stay hidden. | Rows 2:4: `outlineLevel="1"` on each, `outlineLevelRow="1"`. Rows 2:7 then 3:4: levels 1, 2, 2, 1, 1, 1. |
| A group is a run of lines at its level or deeper: groups that touch are one. | Rows 2:3 and 4:5 grouped: one group 2:5. |
| Seven levels at most; an eighth Group is refused, nothing changed. | The 8th nested Group failed ("Unable to get the Group property"); `outlineLevelRow="7"`. |
| A group's summary line is the one after it, or before it with `<outlinePr summaryBelow="0"/>` (`summaryRight="0"` for columns). Setting the default back takes outlinePr away, and sheetPr when it is left empty. | SummaryRow above: `<sheetPr><outlinePr summaryBelow="0"/></sheetPr>`; set below again: no sheetPr. |
| Collapsing hides the group's lines (`hidden="1"`) and writes `collapsed="1"` on the summary line, which gets a `row` (`col`) element for it if it has none. | Rows 2:4 at level 1: rows 2 to 4 `hidden="1"`, `<row r="5" collapsed="1">`. Columns B:D: `<col min="5" max="5" width="9.140625" collapsed="1"/>`. |
| A column hidden by the outline is written `width="0" hidden="1" customWidth="1"`, unless it has a width of its own, which it keeps. | B and D `width="0"`; C at 20 keeps `width="20.7109375" hidden="1" customWidth="1"`. |
| Level button n (ShowLevels) hides every line at level n or deeper and shows the rest; a group of level n or deeper is marked collapsed where its summary line stays in view. | Two levels, button 2: rows 3:4 hidden, row 5 collapsed. Button 1: rows 2:7 hidden, row 8 collapsed and row 5 not (it is hidden itself). Button 3 after 1: all shown, no collapsed left. |
| Hide Detail at a summary line hides the group it sums up; at a line inside a group, the innermost group holding it. | ShowDetail = False on row 5 (summary of 3:4, inside 2:7): only 3:4 hidden. On row 3 of 2:4: 2:4 hidden. |
| Show Detail at a collapsed summary shows its lines but leaves groups inside it that are collapsed hidden; at a hidden line it changes nothing. | Inner then outer hidden, outer shown: rows 2, 5, 6, 7 shown, 3:4 still hidden, row 5 still collapsed. |
| Ungroup takes one level off; hidden lines and the summary's collapsed stay. Ungroup of lines not grouped is refused. | A collapsed 2:4 ungrouped: rows still `hidden="1"`, `<row r="5" collapsed="1">`, no `outlineLevelRow`. |
| Clear Outline takes the levels away, shows the lines the outline hid and the collapsed marks go; the columns come back at the default width, keeping `customWidth`. | `<col min="2" max="3" width="9.140625" customWidth="1"/>`, no outlineLevelCol. |
| Ctrl+8 (the window's DisplayOutline) off writes `showOutlineSymbols="0"` on the sheetView and on sheetPr's outlinePr, even on a sheet without an outline; on again takes both away. | `<sheetPr><outlinePr showOutlineSymbols="0"/></sheetPr>`, `<sheetView tabSelected="1" showOutlineSymbols="0" workbookViewId="0"/>`. |
| Auto Outline makes the outline from formulas: a formula over its own column's cells just above it groups those rows, one over its own row's cells left of it groups those columns. | SUM(A1:A3) in A4, SUM(A1:B1) in C1 and so on: rows 1:3 and columns A:B at level 1. |

## Excel's rules: Subtotal

| Rule | Seen |
|---|---|
| At each change in the chosen column a row is inserted under the group holding "<value> Total" in that column (bold, a format with the bold font) and `SUBTOTAL(n,C2:C3)` over the group's own rows in each column chosen; a Grand Total row under the list over all its rows. | Region East, East, North, West...: row 4 "East Total", C4 `=SUBTOTAL(9,C2:C3)`; row 11 "Grand Total", `=SUBTOTAL(9,C2:C9)`. |
| SUBTOTAL's numbers and the words: Sum 9 Total, Count 3 Count, Average 1 Average, Max 4 Max, Min 5 Min, Product 6 Product, Count Numbers 2 Count, StdDev 7, StdDevp 8, Var 10, Varp 11 (each its own name). | "East Product", "East Count" (Count Numbers), "Grand StdDevp". |
| The list is outlined: its rows at level 2, the subtotal rows at 1, the grand total at none. | `outlineLevelRow="2"`. |
| Summary below data off: the subtotal rows go over their groups, the Grand row under the headings, and `summaryBelow="0"` is written. | Row 2 Grand Average over C4:C11, row 3 East Average over C4:C5. |
| Subtotals added to subtotals (Replace current subtotals off) go inside them: next to their group's rows, the rows a level deeper, the new grand total nearest the list's rows. | Max then Min: rows 2:3 at 3, East Min at 2, East Max at 1; Grand Min row 14 over Grand Max row 15, `SUBTOTAL(5,C2:C11)`. |
| Replace current subtotals: the SUBTOTAL rows there deleted first. | Max then Sum of Price with Replace: only the Sum rows. |
| Page break between groups: a manual break after each group's subtotal row, the last one after the grand total. | `<brk id="4"/>`, `id="6"`, `id="11"` with `max="16383" man="1"`. |
| Remove All deletes the rows holding SUBTOTAL, clears the list's outline and takes its page breaks away. | The list back as typed, no outlineLevelRow, no rowBreaks. |
| The list is the selection, or the table round the active cell. | B3 selected: the same result as A1:D7. |

## The outline bar

| Rule | Seen |
|---|---|
| The bar takes 19 pixels for each level button (levels + 1) and 5 more, beside the row headings and over the column headings alike; none without an outline or with its symbols off (sheetView `showOutlineSymbols="0"`). | `Window.PointsToScreenPixelsX(0)` and `Y(0)` at 100% and 96 dpi: 43px with one level, 62, 81, 100, 119, 138, 157 with up to seven; 0 with DisplayOutline off. |

Not measured (Excel's glyphs aren't in its PDF, and COM doesn't reach them): the buttons' own size (15px here, in a
19px slot), the line's look and the tick at its far end. The bar isn't drawn over frozen rows or columns.

## Not yet

- The Settings box's Automatic styles and Create and Apply Styles.
- Rows inserted inside a group take the format of the row above but not its outline level.
- Subtotal's cells take the formats of the row above them (Excel's own formats there weren't measured, the probe's cells
  had none); Excel's message for a list without headings.
