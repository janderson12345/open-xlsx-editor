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
  are not used yet (all rows use `ht` or the sheet default).
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
```

The two cells still off across: a centred TRUE (our text 3.4px wider than Excel's) and a number in an 8px column
that doesn't fit, which Excel starts at the left instead of right-aligning.

Next, from the same probe:

- Printed rows: 15pt steps 20.3px (20 here), 60pt 81.3px, about x 1.015 but not as cleanly as columns (20pt steps
  27.0). That decides rows per page: large.xlsx is 114 rows of pages in Excel, 112 here.
- Text baseline in a row: Excel 4.3px above the row's bottom edge, 5px here.
- Excel's text is narrower than ours in the same font ("Amount" 47.7px against 52.5): fonts are not mapped yet.

The fixtures are a start; the numbers that matter come from real workbooks (`..\samples`).

## Open differences

- Styles (fonts, fills, borders, alignment, number formats) are not read yet: everything shows in Calibri 11 General.
- Gridlines show under text that runs on into empty cells; Excel hides them there.
- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.
- Printing: no scaling or fit-to-page, print titles, manual page breaks, print area, headers and footers yet.

## Tried and dropped

(Nothing yet.)
