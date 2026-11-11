# Charts

How charts are drawn (`src/ooxml/charts.js` reads them, `src/editor/chartview.js` draws them), inserted
(`src/ooxml/chartnew.js`, `chartxml.js`) and edited (`src/ooxml/chartedit.js`). A chart's data follows its cells
(a value typed in redraws it); a chart not edited is saved as it was.

## Inserting a chart (Insert > Charts)

Fourteen kinds, by Excel's menus: clustered, stacked and 100% stacked column and bar; line, with or without markers;
pie and doughnut; area and stacked area; scatter with markers, straight lines or smooth lines. One cell selected: the
table round it (its current region), as Excel takes it.

Measured over COM (`tools/excel-compare/excel-chart-insert.ps1`: `Shapes.AddChart2` of each kind in Excel's default
style on a small made-up table, saved, `test/chartnew-excel.json`):

- Each kind's chart part is Excel's text for text (14 of 14), but for the two axis ids and the series' unique ids,
  which Excel makes up afresh each time (made up here too). The chart's style and colour parts Excel adds alongside
  are not written: Excel opens the chart without them and draws it in the same default style.
- The anchor: where Excel puts a 5 by 3 inch chart from its top-left cell (from F2: to M16, 304800 and 76200 EMU in).
  Insert > Charts puts it two columns past the selection, a row below its top.
- Excel opened the workbook of all fourteen written here (`excel-chart-readback.ps1`, `test/chartnew-readback-excel.json`)
  and read each as the type it made, with its two series' formulas and its title.

Which cells are series, as Excel takes them (ten shapes of selection; each series' formula from `AddChart2`):

| Rule | Seen |
|---|---|
| The first row holds the names (or categories) when none of its cells is a number; the first column likewise. A blank corner makes both labels. | Region, Sales, Cost over names and numbers: names B1 and C1, categories A2:A5. Years under a "Year" heading: the years are numbers, so data. Years under a blank corner: categories. |
| Series run down the columns when the numbers have more rows than columns; else, ties included, along the rows. | 4 rows by 2: two series. 2 rows by 6: two series along the rows. 3 by 3: three series along the rows. |
| A title Excel makes itself: the series' name with one series, or for a pie (which draws only its first series), else "Chart Title". | |

## Editing a chart

Click a chart to pick it (outlined, eight handles; Ctrl+Alt+5 from the keyboard, Tab to the next). Drag to move it,
a handle to size it, the arrow keys to move it a cell; Delete deletes it; Enter or its menu's Edit Chart... opens the
box for its kind, title, legend and data labels. Each is one step of Undo.

Measured over COM on a clustered column chart (`tools/excel-compare/excel-chart-edit.ps1`, `test/chartedit-excel.json`):
15 of 15 edits written here as Excel writes them, text for text.

| Edit | As Excel writes it |
|---|---|
| A title's text | A `c:tx` of rich text in the title (14pt, Excel's grey); `autoTitleDeleted` 0. |
| No title | No `c:title`; `autoTitleDeleted` 1. |
| Legend moved, or none | `legendPos` b, t, l, r; none: no `c:legend`. |
| Data labels (Show) | Each series a `c:dLbls` of values (9pt, 75% grey, leader lines in an extension), the chart's own `showVal` 1. |
| Change Chart Type | The series kept, drawn as a new chart of that kind is, but: the old gap width kept; pie and doughnut slices without outlines; a doughnut's hole 50%; a scatter chart's hidden line 25400 EMU. The old axes kept and turned: positions for bar, `crossBetween` midCat for area and scatter, a category axis becoming the scatter's x value axis. The title, legend and blanks setting as they were. |

Moved or sized: the anchor's two corners, as the cells and offsets under the new rectangle (not yet measured against
Excel's own rounding). Deleted: the anchor, its relationship and the chart part with what only it uses.

Not yet: editing a chart part not written with the `c:` prefix (refused with a message), the data range, series
one at a time, axis titles, secondary axes, combo charts and the 2016 kinds (waterfall, histogram, treemap, ...).

## What is read

From each sheet's drawing, every two-cell anchor whose graphic frame points to a chart part: where it sits (its two
corner cells and the offsets into them), its kind (column, bar, line, area, pie, doughnut, scatter; clustered, stacked,
100% stacked), its title (its own text, a cell, or, with none, the series' name when there is one series), its legend's
place, and its series: names, categories and values read from the cells they point to in this workbook, else from the
file's cache. `test/fixtures/charts.xlsx` is written as Excel writes chart parts; opened in Excel over COM, it reads
four charts of the kinds, series and titles the editor reads, at the same places.

## Excel's automatic value axis

Measured over COM: 34 sets of numbers, each drawn as a chart in a hidden Excel and its value axis read
(`Axes(xlValue).MinimumScale`, `MaximumScale`, `MajorUnit`: `tools/excel-compare/excel-axes.ps1` on `axes-cases.json`,
giving `test/axes-excel.json`). `npm test` (`test/charts.test.js`) scales each as Excel did: 34 of 34.

| Rule | Seen |
|---|---|
| The axis starts at 0 when the numbers' spread is more than a sixth of the largest (of the smallest, all negative); exactly a sixth isn't more. | 100 to 322: 0 to 350. 500 to 600: 440 to 620. |
| Otherwise it starts half the spread below the smallest number. | 98 to 102: 96 to 103. |
| The far end gets 5% of the spread more (mixed signs: both ends). | 1 to 5: to 6. -5 to 5: -6 to 6. |
| The major unit is the smallest of 1, 2 and 5 times a power of ten that takes both ends out to whole units in no more than ten. | 1 to 21: 2 would make 12, so 5 (0 to 25). 1 to 19: 2 (0 to 20). |
| One number, or all alike: from 0. All zero: 0 to 1 by 0.1. | 42: 0 to 45 by 5. 7, 7, 7: 0 to 8. |
| Column, bar and line charts scale alike. Stacked charts scale the stacks' sums, 100% stacked run 0% to 100%. | Stacked bars of three regions: 0 to 450 by 50, as Excel's. |

## How they look

In the look of Excel's default chart style: a white area with a light grey border, the title in grey 14pt, grey 9pt
axis labels in the series' number format, light grey gridlines, the series in the theme's accent colours (a series'
own colour when it has one), bars spaced by the chart's gap width and overlap, pie slices clockwise from the top, a
bar chart's first category at the bottom. The pictures are hidden from screen readers; each chart is given to them as a
figure after the sheet, its kind and title as the caption and its data as a table.

## Not yet

Scatter, bubble, radar, stock, surface, combo and the 2016 kinds (waterfall, histogram, treemap...); data labels,
markers' own shapes, trendlines, error bars, secondary axes, a logarithmic or date axis, axis titles, a chart's own
fonts, fills and layout; selecting, moving and resizing a chart; charts on chart sheets; and a title borrowed from
the series drawn bold and smaller, as Excel draws it.
