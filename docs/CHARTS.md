# Charts

How charts are drawn (`src/ooxml/charts.js` reads them, `src/editor/chartview.js` draws them). Read only: the chart
parts are saved as they are, and a chart's data follows its cells (a value typed in redraws it).

## What is read

From each sheet's drawing, every two-cell anchor whose graphic frame points to a chart part: where it sits (its two
corner cells and the offsets into them), its kind (column, bar, line, area, pie, doughnut; clustered, stacked, 100%
stacked), its title (its own text, a cell, or, with none, the series' name when there is one series), its legend's
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
