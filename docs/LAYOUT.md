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
- **General format** (*measured*). Probe `test/fixtures/general.xlsx`: 43 values in columns of 30, 8.43, 6, 4 and 2.5
  characters, all 215 as Excel shows them (`test/general-excel.json`, checked by `npm test`). The room is the
  characters of the digit width that fit the cell less its 5px padding (`colChars`: 64px holds 8), at most 11 plus a
  minus sign. Fixed notation when it shows the value exactly, else whichever form has room for more significant
  digits, fixed on a tie; rounding is decimal, half up, on 15 digits; a cell where nothing fits is filled with `#`;
  a value too small to show is `0` (`-0` when negative). Merged cells use the merged width; print uses the screen
  width, so both show the same text. The textbook rule (scientific below 1E-4, also GridCraft's) is wrong: Excel
  shows 0.00001 as `0.00001` in a wide column.

## Where it stands

The generated fixtures, Excel 365 and Edge on Windows (runs in `..\runs\fixtures`; general.xlsx joined at run 4):

```
1b (before)  page counts 1/4 on same page 9027/85124 x within 1px   18/9027 (0.2%)   y within 1px 32/9027 (0.4%) both  3/9027 (0.0%)
2  (columns) page counts 3/4 on same page 9059/100147 x within 1px 9057/9059 (100.0%) y within 1px 51/9059 (0.6%) both 51/9059 (0.6%)
3  (rows)    page counts 4/4 on same page 100147/100147 x within 1px 100145/100147 (100.0%) y within 1px 100147/100147 (100.0%) both 100145/100147 (100.0%)
4  (General) page counts 5/5 on same page 100364/100364 x within 1px 100363/100364 (100.0%) y within 1px 100364/100364 (100.0%) both 100363/100364 (100.0%)
```

Run 3 repeated with no change (3b) gave the same report to the cell: no noise seen on the fixtures.

The one cell still off: a centred TRUE (our text 3.4px wider than Excel's). The 8 of Excel's 100,372 cells not
matched at all are text that reaches its cell's edge here (ours is wider: fonts are not mapped yet, "Amount" 47.7px
in Excel against 52.5), which the extraction reads as one piece with the next cell's (`c3c4`, `ColumnMColumnN`).

Next: fonts (the Normal style's font, Carlito for Calibri where Calibri is missing, text widths). All of this so far
is on generated fixtures; the numbers that matter come from real workbooks (`..\samples`).

## Open differences

- Styles (fonts, fills, borders, alignment, number formats) are not read yet: everything shows in Calibri 11 General.
- Gridlines show under text that runs on into empty cells; Excel hides them there.
- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.
- Printing: no scaling or fit-to-page, print titles, manual page breaks, print area, headers and footers yet.

## Tried and dropped

(Nothing yet.)
