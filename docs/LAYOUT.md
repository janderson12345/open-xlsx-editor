# Layout: matching Excel

Every rule here is to be measured against **Excel's own PDFs** (`tools/excel-compare`, not built yet), never another
program's output. A change is kept only if the total doesn't get worse and every workbook that got worse is
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

## Open differences

- Styles (fonts, fills, borders, alignment, number formats) are not read yet: everything shows in Calibri 11 General.
- Gridlines show under text that runs on into empty cells; Excel hides them there.
- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.

## Tried and dropped

(Nothing yet.)
