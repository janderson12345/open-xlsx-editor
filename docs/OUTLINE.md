# Outlines: Group, Ungroup, Subtotal

How rows and columns are grouped into an outline, collapsed and expanded, and how Data > Subtotal makes one
(`src/ooxml/outline.js`, `src/ooxml/subtotal.js`), and what Excel was seen to do. Each rule was measured over COM in a
hidden Excel (`tools/excel-compare/excel-outline.ps1`: `Range.Group`, `Ungroup`, `Outline.ShowLevels`,
`Range.ShowDetail`, `ClearOutline`, `AutoOutline`, `Outline.SummaryRow` and `SummaryColumn`, `Range.Subtotal`,
`RemoveSubtotal`, `Outline.AutomaticStyles`, `Range.ApplyOutlineStyles`, each on a new workbook of made-up cells, saved and read back). `npm test` (`test/outline.test.js`, 51
cases; `test/subtotal.test.js`, 11 cases; `test/subtotalformats.test.js`, 4 cases; `test/outlinestyles.test.js`, 9 cases) expects what Excel wrote, to the attribute; `test/outline.mjs` clicks the
outline bar's buttons and the Subtotal box in a browser.

## In the editor

- Data > Outline: Group (Shift+Alt+Right) and its menu's Auto Outline; Ungroup (Shift+Alt+Left) and its menu's Clear
  Outline; Subtotal; Show Detail; Hide Detail; Settings (Summary rows below detail, Summary columns to right of detail, Automatic
  styles; its Create makes the outline from the formulas, its Apply Styles styles the selection's summary cells).
- Whole rows or columns selected are grouped at once; any other range asks Rows or Columns, as Excel's Group box does.
- Show Detail, Hide Detail and Clear Outline go by the selection: whole columns, the columns; whole rows, the rows;
  otherwise the rows when the active cell's row is in a row group, the columns when not. Clear Outline with one cell
  selected clears the sheet's whole outline.
- The outline bar: left of the row headings and over the column headings, the level buttons 1, 2, 3... in the corner, and
  for each group a line along its rows (columns) and a − at its summary row, a + while its rows are hidden. With frozen
  panes, the frozen rows' and columns' part of the bar stays in place with them, as the headings do.
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
| A row (column) inserted takes the outline level of the one before it, the row above (column left), as it takes its format; it is shown even inside a collapsed group, and the group it joins is no longer collapsed. | Rows 2:4 grouped, a row inserted at 3: new row 3 level 1; at 2: level 0; at 5 (the summary): level 1. Collapsed, a row inserted at 3: shown, level 1, row 6 without `collapsed`. Column inserted at C in B:D: B:E level 1. |
| Auto Outline makes the outline from formulas: a formula over its own column's cells just above it groups those rows, one over its own row's cells left of it groups those columns. | SUM(A1:A3) in A4, SUM(A1:B1) in C1 and so on: rows 1:3 and columns A:B at level 1. |

## Excel's rules: Subtotal

| Rule | Seen |
|---|---|
| At each change in the chosen column a row is inserted under the group holding "<value> Total" in that column (bold, a format with the bold font) and `SUBTOTAL(n,C2:C3)` over the group's own rows in each column chosen; a Grand Total row under the list over all its rows. | Region East, East, North, West...: row 4 "East Total", C4 `=SUBTOTAL(9,C2:C3)`; row 11 "Grand Total", `=SUBTOTAL(9,C2:C9)`. |
| A subtotal row takes the formats of the list row next to it, as an inserted row does: the one over it with Summary below data, the one under it with the summaries above (the headings' format never); its own format and height, and each cell whose format isn't the row's (an empty cell made for it). The label adds bold to its cell's format. | Region filled yellow, Units 0.00, Item italic: A4 bold and yellow, B4 italic (empty), C4 0.00. Above: the Grand row italic as the first row of data, not bold as the headings. A bold 30pt row: East Total's row bold, 30pt. |
| SUBTOTAL's numbers and the words: Sum 9 Total, Count 3 Count, Average 1 Average, Max 4 Max, Min 5 Min, Product 6 Product, Count Numbers 2 Count, StdDev 7, StdDevp 8, Var 10, Varp 11 (each its own name). | "East Product", "East Count" (Count Numbers), "Grand StdDevp". |
| The list is outlined: its rows at level 2, the subtotal rows at 1, the grand total at none. | `outlineLevelRow="2"`. |
| Summary below data off: the subtotal rows go over their groups, the Grand row under the headings, and `summaryBelow="0"` is written. | Row 2 Grand Average over C4:C11, row 3 East Average over C4:C5. |
| Subtotals added to subtotals (Replace current subtotals off) go inside them: next to their group's rows, the rows a level deeper, the new grand total nearest the list's rows. | Max then Min: rows 2:3 at 3, East Min at 2, East Max at 1; Grand Min row 14 over Grand Max row 15, `SUBTOTAL(5,C2:C11)`. |
| Replace current subtotals: the SUBTOTAL rows there deleted first. | Max then Sum of Price with Replace: only the Sum rows. |
| Page break between groups: a manual break after each group's subtotal row, the last one after the grand total. | `<brk id="4"/>`, `id="6"`, `id="11"` with `max="16383" man="1"`. |
| Remove All deletes the rows holding SUBTOTAL, clears the list's outline and takes its page breaks away. | The list back as typed, no outlineLevelRow, no rowBreaks. |
| The list is the selection, or the table round the active cell. | B3 selected: the same result as A1:D7. |

## Excel's rules: automatic styles

| Rule | Seen |
|---|---|
| Automatic styles is outlinePr's `applyStyles="1"`; off takes it away. | `<sheetPr><outlinePr applyStyles="1"/></sheetPr>`. |
| A group's summary row takes the built-in style RowLevel_L of its group's level L (`builtinId="1" iLevel="L-1"`), a summary column ColLevel_L (`builtinId="2"`): level 1 bold, level 2 italic, deeper the Normal font. The style's own format: its font, nothing else applied; the cell's: the style's font, no apply flags. | `<cellStyle name="RowLevel_2" xfId="3" builtinId="1" iLevel="1"/>`, `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>`. |
| Where a styled row and column meet, the lower level wins, the row on a tie. | E5 (row level 2, column level 1): ColLevel_1; E8 (both level 1): RowLevel_1. |
| Apply Styles styles every cell of the range on a summary row or column, empty ones made; one cell, only it. | A1:H12: A5:H5, A8:H8 and E1:E12; A5 alone: A5. |
| Group with Automatic styles on gives the new group's summary row the style as its own format (`s` with `customFormat`), and its cells. | Rows 2:4: `<row r="5" s="1" customFormat="1">`, A5 RowLevel_1. |
| Subtotal with Automatic styles on does the same for its rows, instead of the bold label. | East Total's row RowLevel_2 (italic), Grand Total's RowLevel_1 (bold). |
| Create (Auto Outline) with Automatic styles on styles the cells of the used range. | A4:C4 RowLevel_1, C1:C3 ColLevel_1. |

Excel writes a new font for RowLevel_3 and deeper even though it is the Normal font; here the Normal font's own entry is
used, which looks the same.

## The outline bar

| Rule | Seen |
|---|---|
| The bar takes 19 pixels for each level button (levels + 1) and 5 more, beside the row headings and over the column headings alike; none without an outline or with its symbols off (sheetView `showOutlineSymbols="0"`). | `Window.PointsToScreenPixelsX(0)` and `Y(0)` at 100% and 96 dpi: 43px with one level, 62, 81, 100, 119, 138, 157 with up to seven; 0 with DisplayOutline off. |

Not measured (Excel's glyphs aren't in its PDF, and COM doesn't reach them): the buttons' own size (15px here, in a
19px slot), the line's look and the tick at its far end.

## Not yet

- Excel's message for a list without headings.
