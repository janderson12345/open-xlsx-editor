# Layout: matching Excel

Every rule here is to be measured against **Excel's own PDFs** (`tools/excel-compare`), never another program's
output. A change is kept only if the total doesn't get worse and every workbook that got worse is
understood. Nothing below has been measured yet: these are the starting points, from the specification.

## Rules in use (unmeasured)

- **Column width.** Stored width `w` in characters of the maximum digit width `mdw`, padding included:
  `px = trunc(((256 w + trunc(128 / mdw)) / 256) mdw)` (ECMA-376 Part 1, 18.3.1.13). `mdw` is fixed at 7px
  (Calibri 11 at 96 dpi) until the Normal style's font is read. `src/ooxml/units.js`.
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

Run 1 (2026-11-08), the three generated fixtures, Excel 365 and Edge on Windows, before any rule was fitted:

```
TOTAL workbooks 3 page counts 1/3 cells matched 85041/100057 on same page 8957/85041 x within 1px 3/8957 (0.0%) y within 1px 8/8957 (0.1%) both 2/8957 (0.0%)
```

What run 1 shows, to be turned into rules one at a time:

- Excel's printed columns are wider than the screen pixels the specification's formula gives: a default column is
  68.8px on Excel's page (64 here), column B (width 20.7109375) 155.9px (145 here), about 7.5% more.
- Printed rows step 20.3px for 15pt (20 here).
- Excel's text is narrower than ours in the same font (Calibri on Excel's page; "Amount" 47.7px against 52.5).
- So basic.xlsx's 8 columns take 2 pages across in Excel and fit on 1 here, and large.xlsx has 342 pages against 224.

The fixtures are a start; the numbers that matter come from real workbooks (`..\samples`).

## Open differences

- Styles (fonts, fills, borders, alignment, number formats) are not read yet: everything shows in Calibri 11 General.
- Gridlines show under text that runs on into empty cells; Excel hides them there.
- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.
- Printing: no scaling or fit-to-page, print titles, manual page breaks, print area, headers and footers yet.

## Tried and dropped

(Nothing yet.)
