# Sparklines

How sparklines are read (`src/ooxml/sparklines.js`) and drawn (`src/editor/sparkgeom.js` works out the shapes,
`src/editor/sparkview.js` draws them as an SVG in the cell, on screen and in the printed pages; `src/pdf.js` draws the
same paths in the PDF export). A sparkline follows its cells: a value typed in redraws it. Sparklines are not edited
yet (no Insert > Sparklines, no Sparkline tab); they are saved back as they came.

## What Excel writes

Measured over COM (`tools/excel-compare/excel-sparklines.ps1`: in an invisible Excel, 15 groups built with
`Range.SparklineGroups.Add` and each option set through COM, the workbook saved and its sheet's extension read back,
`test/sparklines-xml-excel.json`):

- The groups are in the worksheet's `extLst`, an `ext` with uri `{05C60535-1F16-4fd2-B633-F4F36F0B64E0}` holding
  `x14:sparklineGroups`; each `x14:sparklineGroup` holds its colours and `x14:sparklines`, each sparkline its source
  (`xm:f`, `Sheet1!B2:M2`) and its cell (`xm:sqref`). A group over O11:O12 is one group of two sparklines.
- The type is written only when it isn't a line: `type="column"`, and win/loss is `type="stacked"`.
- An option is written only when on, or not its default: `markers`, `high`, `low`, `first`, `last`, `negative`,
  `displayXAxis`, `displayHidden`, `rightToLeft`, `lineWeight` (0.75 unwritten), `minAxisType` and `maxAxisType`
  (`group`, `custom`; individual unwritten) with `manualMin` and `manualMax`, `dateAxis="1"` with the dates' range as
  an `xm:f` in the group before `x14:sparklines`.
- `displayEmptyCellsAs`: `gap` and `span` written; zero (the schema's default) left out.
- All eight colours are always written (`colorSeries`, `colorNegative`, `colorAxis`, `colorMarkers`, `colorFirst`,
  `colorLast`, `colorHigh`, `colorLow`), as `rgb` when set through COM; the ribbon's styles write theme colours (Style
  1: accent 1 darkened by half for the series and markers, accent 2 for negative, accent 1 lightened 40% for first
  and last, accent 1 for high and low, black for the axis). Groups are written newest first.

Every group and option read back as COM reported it (`test/sparklines.test.js`).

## How Excel draws them

Excel prints each sparkline as a 200 dpi picture of its cell. Pictures from Excel's own PDFs, measured pixel by pixel
(`tools/excel-compare/sparkimages.cjs` takes the pictures out, `spark-features.py` measures them): the probe's sheets
(cells 8, 20 and 45 characters wide by 12, 24, 48 and 96 points; lines of 0.25 to 6 points) and
`test/fixtures/sparklines.xlsx` (28 sparklines, one per option, `test/sparklines-excel.json`).

| Rule | Measured |
|---|---|
| The plot sits 1.92px (4 picture pixels) in from the cell's sides, 3.84px (8) below its top, 4.08px (8.5) above its bottom, whatever the cell's size. | First bars at 4.32, 7.2 and 13.92px in cells 65, 156 and 344px wide; tops 3.84px down at every height. |
| The width is cut into one slot per value; a line's points are at the slots' middles; a column fills the middle 1/1.4 of its slot. | 4 values in 344px: bars 60.64px wide every 84.96px; line points at 44.4, 129.5, 214.3, 299.2. |
| A date axis cuts the width into one slot per day from the first date to the last, each point in its day's slot. | 12 dates over 43 days: points from 4.5px, 5.28px a day. |
| Right to left mirrors the slots: the first value at the right. | |
| Vertically the greatest value is at the plot's top, the least at its bottom (the sparkline's own, the group's, or the custom ends). All one value: a line along the top. | |
| A column runs from zero when the axis crosses it, else from the end nearer zero: over positive numbers from the least (whose bar is a hairline, a picture pixel tall), over negative ones from the greatest. Bars past a custom end stop at it. | "Column, points", "Column, negative", "Custom max 8". |
| Win/loss: wins from the middle up to the top, losses down to the bottom, nothing for 0. | |
| Markers are diamonds, wider with the line: 3.36px up to 0.75pt, 4.8 at 1.5pt, 7.68 at 2.25pt, 9.12 at 3pt, 12.96 at 6pt. The line is the weight wide. | Sheet Weights. |
| Which colour a point takes, each over the ones before: markers (or the series for columns), negative, first, last, high, low. Every point at the high or low value is one. Negative points get markers even with markers off. | A first point that is the high one: green (high); a last that is the low one: magenta (low); a negative low point: low. |
| Empty cells: gaps break the line (a point left alone draws nothing), zero plots them as 0, connect passes over them. A text cell is plotted as 0. | "Empty as gaps", "Text in the data". |
| Hidden rows and columns are left out (no slot) unless the group shows hidden data. | Hidden column N. |
| The axis, when asked for, only where zero is on the vertical axis: none over positive numbers alone. Across the plot, a pixel wide, over the line. | "Line, axis, positive". |

## How close

`test/sparklines.test.js` draws each of the fixture's 28 sparklines at the size Excel printed its cell and compares
the shapes with Excel's pictures: 120 bars, every edge within 0.9px of Excel's (mean 0.30px), each bar Excel's colour;
166 line segments, each middle within 1.27px of Excel's line (mean 0.36px; the steep ones differ most, where Excel's
antialiasing steps); 19 coloured markers within 0.62px of Excel's centres, their width within 0.48px; the axis within
0.2px. On screen the same sizes are used in CSS pixels.

## Not yet

Inserting and editing sparklines (Insert > Sparklines, the Sparkline tab); sparklines don't follow rows and columns
inserted or deleted, and a group's source on another sheet isn't renamed with it. Not measured: an axis when the least
value is exactly 0, a column sparkline of one value repeated, error values in the data (treated as empty).
