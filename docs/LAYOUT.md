# Layout: matching Excel

Every rule here is to be measured against **Excel's own PDFs** (`tools/excel-compare`), never another program's
output. A change is kept only if the total doesn't get worse and every workbook that got worse is
understood. Rules marked *measured* were fitted on Excel's PDFs; the rest are starting points from the specification.

The probe for printed geometry is `test/fixtures/widths.xlsx`: 40 column widths (text in row 1 at each column's left
edge, a number in row 2 at its right edge) and 20 row heights (text in column A).

## Rules in use

- **Column width.** Stored width `w` in characters of the maximum digit width `mdw`, padding included:
  `px = trunc(((256 w + trunc(128 / mdw)) / 256) mdw)` (ECMA-376 Part 1, 18.3.1.13). `mdw` is fixed at 7px
  (Calibri 11 at 96 dpi) until the Normal style's font is read. `src/ooxml/units.js`. On screen; still to be checked
  against Excel's own screen.
- **Printed column width** (*measured*). Screen pixels x 1.075 (`PRINT_COL_SCALE`): 30 columns of the probe from 8 to
  285 screen pixels all within 0.15px (64 prints as 68.8, 145 as 155.9). Measured with Excel 365 exporting through
  the default printer's driver (a 600 dpi inkjet): the factor may belong to the printer, which a second machine
  would show. Run 2: cell text x within 1px 18/9027 -> 9057/9059, page counts 1/4 -> 3/4.
- **Printed cell padding** (*measured*). Text starts 3.2px into its cell; right-aligned text ends 1.9px before the
  cell's right edge (every column of the probe). On screen still 2px each side.
- **Default column width.** `defaultColWidth` when given; else `baseColWidth` (8) characters plus 5px padding, rounded
  up to a multiple of 8px: 64px for Calibri 11. The multiple of 8 is from [MS-OI29500]'s note on `baseColWidth`.
- **Row height.** Points to pixels at 96 dpi, rounded: 15pt is 20px. `customHeight` and the tallest font in the row
  are not used yet (all rows use `ht` or the sheet default). On screen; still to be checked against Excel's screen.
- **Printed row height** (*measured*). Whole device pixels of a 600 dpi printer (6.25 to a CSS pixel) of
  pt x 1.3565 px, truncated (`printRowPx`): 15pt prints as 127 device pixels, 20.32px. Fitted on the probe's 21 rows
  of 8 to 60pt (baselines all within 0.12px of one offset from their row's bottom; plain pt x 1.353 spreads 0.43px)
  and on large.xlsx's 43 steps of 15pt rows. Same printer caveat as the columns. Run 3: page counts 3/4 -> 4/4, cells
  on the same page 9059 -> 100147, both ways within 1px 51 -> 100145.
- **Printed baseline** (*measured*, Calibri 11). 4.15px above the row's bottom edge. In print the cell takes the row's
  full height and 0.15px of bottom padding lifts Edge's baseline (4px above the bottom of an 18px line box) to it.
- **Gridlines.** The right and bottom pixel of each cell, inside its width and height.
- **Cell text.** Bottom-aligned, 2px side padding, numbers right, text left, booleans and errors centred, no wrapping;
  text runs on to the right over empty cells. Line height is Calibri's hhea ascent + descent + gap (1.2207 em),
  through `--oxe-cell-lh` only.
- **General format.** Up to 11 characters; scientific from 1E+11 and below 0.0001. A number too wide for its column
  is clipped; Excel shows fewer digits or `###` instead.

## Where it stands

The four generated fixtures, Excel 365 and Edge on Windows (runs in `..\runs\fixtures`):

```
1b (before)  page counts 1/4 on same page 9027/85124 x within 1px   18/9027 (0.2%)   y within 1px 32/9027 (0.4%) both  3/9027 (0.0%)
2  (columns) page counts 3/4 on same page 9059/100147 x within 1px 9057/9059 (100.0%) y within 1px 51/9059 (0.6%) both 51/9059 (0.6%)
3  (rows)    page counts 4/4 on same page 100147/100147 x within 1px 100145/100147 (100.0%) y within 1px 100147/100147 (100.0%) both 100145/100147 (100.0%)
```

Run 3 repeated with no change (3b) gave the same report to the cell: no noise seen on the fixtures.

The two cells still off across: a centred TRUE (our text 3.4px wider than Excel's) and a number in an 8px column
that doesn't fit, which Excel starts at the left instead of right-aligning.

The 10 of Excel's 100,157 cells not matched at all, and why:

- General format: Excel shows 0.000012345 as `0.000012345` (the rule here gives `1.2345E-05`), and fits
  123456789012 into a 64px column as `1.23E+11` where this clips `1.23457E+11`.
- A number too wide for its column is filled with `#` in Excel (the probe's 1.29-character column).
- Text that reaches its cell's edge here (ours is wider: fonts are not mapped yet, "Amount" 47.7px in Excel against
  52.5) reads as one piece with the next cell's in the extraction (`c3c4`, `ColumnMColumnN`).

Next: the General format's switch-over points and narrow-column digits, then fonts (text widths). All of this so far
is on generated fixtures; the numbers that matter come from real workbooks.

The fixtures are a start; the numbers that matter come from real workbooks (`..\samples`).

## Open differences

- Styles (fonts, fills, borders, alignment, number formats) are not read yet: everything shows in Calibri 11 General.
- Gridlines show under text that runs on into empty cells; Excel hides them there.
- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.
- Printing: no scaling or fit-to-page, print titles, manual page breaks, print area, headers and footers yet.

## Tried and dropped

(Nothing yet.)
