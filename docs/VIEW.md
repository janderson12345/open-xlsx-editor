# Frozen panes and zoom

How the sheet's view is read and written (`src/ooxml/view.js`, read by `sheet.js`) and shown (the frozen panes and
zoom in `src/editor/grid.js`, the View tab in `src/editor/viewtab.js`), and what Excel was seen to do. Measured over
COM: Excel's own ribbon commands (`CommandBars.ExecuteMso`: FreezePanes, FreezeTopRow, FreezeFirstColumn) and
`Window.FreezePanes`, `SplitRow`, `SplitColumn`, `Zoom`, on a new workbook in a hidden instance, and the files it saved.
`npm test` (`test/view.test.js`) expects what Excel wrote; a workbook written here reads back in Excel with the same
frozen rows and columns, zoom and active cell on each sheet.

## In the editor

- View > Freeze Panes: Freeze Panes (above and left of the active cell), Freeze Top Row, Freeze First Column; once
  frozen, the first entry is Unfreeze Panes, as Excel's menu turns.
- View > Zoom: Zoom... (10% to 400%), 100%, Zoom to Selection; Ctrl+wheel zooms by 10% a notch.
- Neither is in the undo history (Excel's aren't either: they change the view, not the cells), but both are saved.
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

A file frozen through Excel's object model (`SplitRow`, then `FreezePanes`) has `state="frozenSplit"`: read the same.

## Not yet

Freeze Panes with A1 active (Excel then freezes the middle of the window; here it asks for another cell), freezing
while scrolled (Excel freezes the rows and columns in view above and left of the active cell; here those from the
sheet's start), a frozen pane that starts further down than row 1 (`topLeftCell` of the sheet view), split panes
(shown unsplit), New Window, Arrange and the status bar's zoom slider. Excel scrolls by whole rows and columns; this
grid scrolls by pixels. Zoom to Selection's percent is worked out here, not measured against Excel's.
