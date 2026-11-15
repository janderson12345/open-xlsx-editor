# Sparklines

How sparklines are read (`src/ooxml/sparklines.js`) and drawn (`src/editor/sparkgeom.js` works out the shapes,
`src/editor/sparkview.js` draws them as an SVG in the cell, on screen and in the printed pages; `src/pdf.js` draws the
same paths in the PDF export). A sparkline follows its cells: a value typed in redraws it. Insert > Sparklines makes
them and the Sparkline tab changes them (`src/ooxml/sparkedit.js` writes the XML; `src/editor/acts-sparklines.js`,
`sparklinetab.js` and `sparkboxes.js` are the ribbon and its boxes); a sheet's sparklines not changed are saved back as
they came.

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

## Inserting and editing

Measured over COM (`tools/excel-compare/excel-sparkedit.ps1`, `test/sparkedit-excel.json`): `SparklineGroups.Add` for
16 pairs of location and data, and the Sparkline tab's edits run through COM from the same start (a line group
O2:O4 over B2:M4 and a column O6 over B6:M6), each saved. `test/sparkedit.test.js` replays each edit here and gets
Excel's saved XML for all 13, its made-up ids aside.

| Rule | Measured |
|---|---|
| The location is one row or column. One cell takes data of one row or column, whole. n cells cut the data into its rows when it has n rows, else into its columns when it has n columns; a square is cut into rows whichever way the location runs. Anything else is refused. | O2:O3 over B2:M3: rows. B6:M6 over B2:M3: 12 columns of 2. O2:O4 and B6:D6 over B2:D4: rows. O2:O3 over B2:M2, O2:P3, and O2 over B2:M3 refused (COM: 0x800A03EC). |
| A piece of one cell is written B2:B2; data of one cell B2. Another sheet's name is quoted as a formula quotes it. | O2:O13 over B2:B13; 'My data'!B2:M2. |
| A new group comes first in the list, with empty cells shown as gaps and Excel's colours (series FF376092, the rest FFD00000, the axis black), its type written unless a line. A sparkline already in one of its cells leaves its group. | Add over O3 of O2:O4: O3 a column group of its own first, O2 and O4 still one group. |
| Group: the selected sparklines move into the active cell's group, after its own, taking its settings; the rest of their old groups stay. | Group of O3 and O6 at O6: O3 joins O6's group; at O3: O6 joins O2:O4's. |
| Ungroup: each selected sparkline becomes a group of its own with its group's settings, put first in turn. | O3: groups O3, then O6, then O2 and O4 together. O2:O3: O3, O2, O6, O4. |
| Clear Selected Sparklines takes the selected ones out (a group left empty goes); Clear Selected Sparkline Groups takes every group with one selected. | |
| Settings are written in the schema's order, defaults left out; a custom end writes manualMin or manualMax, the other end not. | lineWeight 1.5 with a custom minimum of -2 and empty cells as zero: `manualMin="-2" lineWeight="1.5" minAxisType="custom"`. |
| A theme colour is written theme and tint, Excel's theme numbering (Accent 1 is 4). | ThemeColor 5, TintAndShade -0.25: `theme="4" tint="-0.249977111117893"`. |
| Edit Group Location & Data keeps the group's settings; Edit Single Sparkline's Data changes one sparkline's source. | Modify to P2:P4 over C2:M4: rows. |

Excel opened the editor's files and read them back as written (`tools/excel-compare/excel-sparkedit-check.ps1`,
`test/sparkedit-readback-excel.json`): a workbook with no extLst given line, column and win/loss groups with every
option and theme, standard and custom colours, and the browser suite's file after the Sparkline tab's edits.

The ribbon: Insert > Sparklines (Line, Column, Win/Loss) opens Create Sparklines, its Data Range filled with the
selection when it holds values, else its Location Range with the selection ($Q$2); a shape Excel refuses keeps the box
open with the reason. The Sparkline tab shows while the active cell holds a sparkline: Edit Data (group location and
data, a single sparkline's data, Hidden & Empty Cells), Type, Show (High, Low, Negative, First, Last Point, Markers for
a line), Sparkline Color (with Weight) and Marker Color (Home's theme and standard colours), Axis (general or date axis,
Show Axis, right to left, each end automatic, the same for all, or custom), Group, Ungroup and Clear. Each change is one
step of the history.

Not measured, where COM can't reach (the boxes are modal): the boxes' own words and what they fill in, and the
colours a sparkline inserted from Excel's ribbon gets; the editor gives a new group the colours Excel's object model
gives (above), which are the gallery's Dark #5. Changes are refused on a protected sheet.

## The style gallery

Excel's 36 sparkline styles can only be applied from its ribbon: a sparkline group has no Style property in COM and
the gallery isn't a command ExecuteMso runs. So `tools/excel-compare/excel-sparkstyles.ps1` makes one Excel of its own
visible (with the owner's agreement) and invokes each gallery item through Windows UI Automation (no keys, no mouse),
saving a copy after each (`test/sparkstyles-excel.json`). What was found:

- A style sets the group's eight colours and nothing else (its attributes were unchanged).
- The 24 accent styles are four rows of six accents: Darker 50% (series accent darkened 50%, negative the next
  accent, markers as the series, first and last lightened 40%, high and low the accent), Darker 25% (series darkened
  25%, negative the next accent, markers and points the next accent darkened 25%), plain (series the accent, markers
  and points it darkened 25%), Lighter 40% (series lightened 40%, negative Background 1 darkened 50%, markers lightened 80%,
  first and last darkened 25%, high and low darkened 50%). Accent 6's "next accent" is accent 1. The axis is always
  black.
- Dark #1 and #2 are theme colours; Dark #3 to #6 and Colorful #1 to #4 fixed colours; Colorful #5 and #6 a different
  theme colour for each part. Theme colours are written with Excel's tints as text (0.39997558519241921).

The Sparkline tab's Style gallery shows the 36 by their series colour and high and low points' colours, named as
Excel names them ("Sparkline Style Accent 1, Darker 50%"), and gives the chosen style to the active cell's group and
every group with a sparkline selected (which groups Excel's gallery takes when several are selected wasn't measured).
`test/sparkedit.test.js` applies each of the 36 and gets the colours Excel wrote.

## Not yet

Sparklines don't follow rows and columns
inserted or deleted, and a group's source on another sheet isn't renamed with it. Not measured: an axis when the least
value is exactly 0, a column sparkline of one value repeated, error values in the data (treated as empty).
