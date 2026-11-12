# Frozen panes, zoom, gridlines and headings

How the sheet's view is read and written (`src/ooxml/view.js`, read by `sheet.js`) and shown (the frozen panes and
zoom in `src/editor/grid.js`, the View tab in `src/editor/viewtab.js`), and what Excel was seen to do. Measured over
COM: Excel's own ribbon commands (`CommandBars.ExecuteMso`: FreezePanes, FreezeTopRow, FreezeFirstColumn) and
`Window.FreezePanes`, `SplitRow`, `SplitColumn`, `Zoom`, `DisplayGridlines`, `DisplayHeadings`, on a new workbook in a
hidden instance, and the files it saved. `npm test` (`test/view.test.js`) expects what Excel wrote; a workbook written
here reads back in Excel with the same frozen rows and columns, zoom and active cell on each sheet.

## In the editor

- View > Freeze Panes: Freeze Panes (above and left of the active cell), Freeze Top Row, Freeze First Column; once
  frozen, the first entry is Unfreeze Panes, as Excel's menu turns.
- View > Zoom: Zoom... (10% to 400%), 100%, Zoom to Selection; Ctrl+wheel zooms by 10% a notch.
- View > Show: Gridlines, Formula Bar and Headings, on or off (Excel's check boxes; here buttons that stay pressed
  while on). Gridlines and Headings are the sheet shown's and saved with it; the Formula Bar is the editor's, for every
  sheet, and not in the file, as Excel keeps it as its own setting. With the headings off they take no room and aren't
  drawn, as in Excel, but are still there for screen readers.
- Page Layout > Sheet Options: Gridlines and Headings, each with View (the same switch as View > Show) and Print.
- None of these is in the undo history (Excel's aren't either: they change the view, not the cells), but all except
  the Formula Bar are saved.
- Frozen rows stay under the column headings and frozen columns beside the row headings while the rest scrolls; a
  thin line marks where. Clicks, keys, resizing and selecting work across the panes as on the sheet.

## Excel's rules

| Rule | Seen |
|---|---|
| Freeze Panes freezes the rows above and the columns left of the active cell. It is written as a `pane` element: `xSplit` (columns), `ySplit` (rows), `topLeftCell` (the first cell of the scrolled pane), `activePane` and `state="frozen"`. | C5: `<pane xSplit="2" ySplit="4" topLeftCell="C5" activePane="bottomRight" state="frozen"/>`. |
| With rows and columns frozen there is a selection for each of the three other panes: the top-right one in row 1 of the first scrolled column, the bottom-left one in column A of the first scrolled row, the active cell's in the bottom-right one. | C5: `topRight` C1, `bottomLeft` A5, `bottomRight` C5. |
| Freeze Top Row writes only `ySplit`, with `activePane="bottomLeft"`; Freeze First Column only `xSplit`, with `topRight`; one selection each, in that pane. | `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>`. |
| Unfreezing takes the pane and the panes' selections away. | `<selection activeCell="C4" sqref="C4"/>`. |
| The zoom is `zoomScale` and `zoomScaleNormal`, 10 to 400; 100 writes neither. | 150: `zoomScale="150" zoomScaleNormal="150"`. 9 and 401 refused. |
| Gridlines off is `showGridLines="0"` on the sheet view, Headings off `showRowColHeaders="0"`; on again takes the attribute away. | Both off: `<sheetView showGridLines="0" showRowColHeaders="0" tabSelected="1" workbookViewId="0"/>`; on again: `<sheetView tabSelected="1" workbookViewId="0"/>`. |
| The Formula Bar (`Application.DisplayFormulaBar`) isn't in the file. | The saved workbook the same as one with the bar shown. |
| Page Layout > Sheet Options' Print boxes are `printOptions`: `gridLines="1"`, `headings="1"`, the element gone when both are off. | `<printOptions headings="1" gridLines="1"/>` before `pageMargins`; both off, no element. Set here from Page Layout > Sheet Options (`src/ooxml/printopts.js`), used for printing (docs/PRINT.md). |

A file frozen through Excel's object model (`SplitRow`, then `FreezePanes`) has `state="frozenSplit"`: read the same.

## Not yet

Freeze Panes with A1 active (Excel then freezes the middle of the window; here it asks for another cell), freezing
while scrolled (Excel freezes the rows and columns in view above and left of the active cell; here those from the
sheet's start), a frozen pane that starts further down than row 1 (`topLeftCell` of the sheet view), split panes
(shown unsplit), the Ruler (Page Layout view only in Excel), New Window, Arrange and the status bar's zoom slider.
Excel scrolls by whole rows and columns; this grid scrolls by pixels. Zoom to Selection's percent is worked out here,
not measured against Excel's.
