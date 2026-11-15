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
- **Printed row height** (*measured*). Depends on the workbook's Normal style font (`printRowPx`):
  - Aptos Narrow: whole device pixels of a 600 dpi printer (6.25 to a CSS pixel) of pt x 1.3565 px, truncated: 15pt
    prints as 127 device pixels, 20.32px. widths.xlsx's 21 rows of 8 to 60pt (baselines within 0.12px of one offset;
    plain pt x 1.353 spreads 0.43px) and large.xlsx's 43 steps of 15pt rows. Run 3: page counts 3/4 -> 4/4.
  - Calibri: the same of pt x 1.291: 15pt is 121 device pixels, 19.36px, read exactly off fills.xlsx's fill
    boundaries; fonts.xlsx's rows of 8 to 60pt step within about 0.1px of it. (A first fit of plain pt x 1.288 on
    text baselines drifted 0.04px a row; run 12 replaced it.)
  - Other Normal fonts: points, unscaled, until measured. No formula relating the two factors was found (columns
    print at x 1.075 under both). Same printer caveat as the columns.
- **Fonts** (*measured*). A cell's typeface is its font's, except that a font of scheme `minor` or `major` is the
  theme's body or heading font whatever its `name` says. A workbook without a theme part gets Excel 365's default
  theme, whose body font is Aptos Narrow (basic.xlsx prints in Aptos Narrow though its fonts say Calibri). Stacks add
  a metric-compatible stand-in (Calibri: Carlito, Cambria: Caladea, Arial, Times New Roman, Courier New: Liberation)
  and for Aptos Narrow, which browsers don't have (Office keeps it in its own cloud-font cache), Open Metric Fonts'
  Morrow Sans Narrow when the page serves it (README, Fonts), else Calibri: as wide within a pixel. Each cell format
  is a CSS class (`cellstyles.js`). Runs 5 -> 9: text width within 1px 0.3% -> 99.0% of all cells; fonts.xlsx 30 ->
  53 of 55.
- **Morrow Sans Narrow served** (*measured*, run "final" in
  `..\runs\morrow3`: print-ours with `OXE_FONTS=1`). It
  has Aptos Narrow's widths and Windows metrics, so Excel's placement of the text is worked out from those
  (`fonts.js` OPEN), with no stand-in offset. Chromium places a web font by its hhea metrics, though, not the Windows
  ones it uses for an installed font: placed by the Windows ones, every baseline was 0.95px high (large.xlsx 61,460 of
  100,000 within 1px). basic.xlsx and large.xlsx: width within 1px 99,028 -> 100,046 of 100,049, across and down
  unchanged (all within 1px); the ten workbooks 100,962 -> 101,981 of 102,095 in width, the rest of the report the
  same, fills and borders too. The cells are placed again when a served font finishes loading (`fontsChanged`).
  Release 1.0.1, its faces split by alphabet and loaded as FontFace objects from `fonts: { omf }` (`omffonts.js`;
  run "v101" in `..\runs\omf101`): the same report to the cell. Installed rather than served, Morrow Sans Narrow was
  placed by its Windows metrics here and every baseline was 0.95px high again (large.xlsx 61,460 of 100,000; a
  workbook made from a CSV too): Chromium places a font by its typo metrics when it sets USE_TYPO_METRICS (both Morrow
  fonts and Aptos Narrow do; Calibri doesn't), installed or served. Placed so (run "typo" in `..\runs\typo`, Morrow
  installed): the ten workbooks 102,082 of 102,095 within 1px down again, large.xlsx all of them.
- **Printed baseline** (*measured*). Excel puts all the bottom-aligned text of a row on one baseline: the Windows
  descent of the row's tallest font plus about 0.9px above its printed row boundary (rows.xlsx, whose fills give
  every boundary exactly; Calibri 11, 14, 20, Arial 14, Times New Roman 16, Cambria 20 within about 0.3px). Excel's
  printed grid starts 0.7px lower than ours (the half-point origin, decor.js), and text is placed in our grid, so the
  offset here is 0.15px (Aptos Narrow, drawn in Calibri: 0.3), fitted by centring 1615 Calibri and 100,000 Aptos
  Narrow baselines; top and middle targets take the 0.7px too. Smaller text in a row is lifted onto its baseline
  (`rowheights.js rowDescents`); a styled empty cell's font counts.
- **Where Chromium puts a baseline** (*measured* in Edge, a zero-size marker on the baseline at several line
  heights): floor((L - (ascent + descent)) / 2) + ascent below the top of a line box of L px, with the font's Windows
  ascent and descent each rounded to a whole pixel. Exact for Calibri 11 to 36, Cambria, Times New Roman, Arial and
  Courier New; the line height is solved from it (`fonts.js cellLineHeight`), and every font's baseline then lands
  where it is aimed within 0.01px. An earlier model with the hhea metrics drifted with size (Calibri 36 3.7px low,
  Arial 14 2px high, even Calibri 11 0.82px low); its errors had been absorbed into the offsets above.
- **Padding for large fonts** (*measured*, rows.xlsx). Excel's cell padding grows with the font: 0.13px more each side
  for every pixel of size above 11pt (Calibri 20 starts 1.6px further in, Calibri 36 4.2px, Cambria 20 1.8px).
  Runs 27 -> 30: text within 1px across 102082 -> 102093 of 102095; rows.xlsx both ways 36 -> 45 of 47, fonts.xlsx
  35 -> 43 of 56, no workbook worse.
- **Automatic row heights** (*measured*, rows.xlsx: all 20 rows' fills within 1px). A row without customHeight is as
  tall as its tallest cell whatever its `ht` says (ht 30 printed at 15pt; ht 10 with a 20pt cell at 26.25pt): the
  font's row height (`fontRowPx`: round(max(Windows ascent + descent, hhea line) x size + 1.6 + 0.035 x size) px,
  fitted on Calibri 11, 14, 20, 36, Cambria 20, Times New Roman 16, Arial 14, all exact), times the lines of wrapped
  text, never below the sheet's default. Styled empty cells count, merged cells don't. Wrapped lines are counted in
  the screen column's width with the browser's font measurement, and step by the same row height.
- **Colours** (*measured*). Theme indexes swap the first two pairs (0 lt1, 1 dk1, 2 lt2, 3 dk2, then the accents);
  tints work in Windows' integer HLS (0-240, Microsoft KB Q29240) and reproduce all 8 of fills.xlsx's tints exactly,
  where floating-point HLS misses 5 by a step; indexed colours from the legacy palette or the workbook's own.
  `theme.js resolveColor` is the one resolver. Patterned fills are drawn as the blend of their two colours for now.
- **Fills and borders** (*measured* in print, fills.xlsx). Excel's printed grid starts half a point (0.67px across,
  0.7px down) beyond margin + sizes; fills run boundary to boundary; border lines are 1.28px a unit (thin 1, medium
  2, thick and double 3) and hair 0.16px, centred on the boundary; the left edge owns both left corners, the top and
  bottom edges start after it and run past the right one. Dotted and hair print solid; dashes are 3.84 on 1.28 off
  (medium dashed 11.52/3.84, dash-dot 11.52/3.84/3.84/3.84), laid along the page. In print solid lines are filled
  boxes (Chromium rounds a border's width down, a box's edges to the nearest pixel). On screen: 1px a unit on the
  gridline pixel, CSS dashes (unmeasured). A filled cell hides its gridlines; the sheet stays white in dark mode.
  Runs 10 -> 13: Excel's fill and border shapes within 1px 32 -> 122 of 153; the rest are double lines (Excel
  shortens the inner line at corners), dash pieces cut by Chromium's whole-pixel snapping, and the hatched fill.
- **Gridlines.** The right and bottom pixel of each cell, inside its width and height.
- **Cell text.** 2px side padding on screen (3.2 and 1.9 in print); general alignment puts numbers right, text left,
  booleans and errors in the middle. Size, line height and baseline shifts only through `celltext.js`'s custom
  properties, on an inner span (`.oxe-tx`).
- **Alignment and wrapping** (*measured* in print, align.xlsx). Centred text is centred on the cell; right-aligned
  text ends where numbers do; an indent level is 10.08px from either side (9 on screen, unmeasured). Top-aligned
  text's baseline is the Windows ascent + 0.8px below the row's top; middle-aligned text centres one line's ascent +
  descent in the row. Text runs on over empty neighbours: left text to the right, right text to the left, centred
  both ways as far as the nearer side allows; numbers and wrapped text never. Wrapped lines step by the font's
  default row height (Calibri 11: 15pt, 19.36px in print), not its line height, and break where Excel's do (text
  widths already match). Run 17: align.xlsx 32 of 32 pieces matched, 30 within 1px both ways. Fill, justify and
  distributed are drawn left, rotation and shrink-to-fit not yet.
- **General format** (*measured*). Probe `test/fixtures/general.xlsx`: 43 values in columns of 30, 8.43, 6, 4 and 2.5
  characters, all 215 as Excel shows them (`test/general-excel.json`, checked by `npm test`). The room is the
  characters of the digit width that fit the cell less its 5px padding (`colChars`: 64px holds 8), at most 11 plus a
  minus sign. Fixed notation when it shows the value exactly, else whichever form has room for more significant
  digits, fixed on a tie; rounding is decimal, half up, on 15 digits; a cell where nothing fits is filled with `#`;
  a value too small to show is `0` (`-0` when negative). Merged cells use the merged width; print uses the screen
  width, so both show the same text. The textbook rule (scientific below 1E-4, also GridCraft's) is wrong: Excel
  shows 0.00001 as `0.00001` in a wide column.
- **Number formats** (*measured*). Probe `test/fixtures/numfmt.xlsx`: 27 built-in formats and 30 custom codes with 14
  values, all 798 pairs as Excel shows them (`test/numfmt-excel.json`, in `npm test`; `format.js`). Sections by sign
  or condition, a minus first and none when every digit shown is zero, decimal half-up rounding, separators through
  zero padding, scaling commas after the decimals too, engineering notation (0 in `##0.0E+0` is `000.0E+0`),
  fractions by continued fractions within the denominator's digits, dates in the 1900 system with its 29 February
  (and 1904), times rounded to the second or to the fractions shown, `[h]` `[mm]` elapsed, `@` and text sections.
  Built-ins 37-40 pad positives with `_)` as Excel does, not ECMA-376's space. In the view (run 21, the probe's 1583
  pieces all within 1px across): `_x` is a hidden x, `?` a blank a digit wide, a fraction with no fractional part
  keeps the room of its space, digits and slash, `*x` repeats x across the room left, a format's colour beats the
  font's, a number in `@` sits left, negative dates and formatted numbers wider than their cell show centred #s (as
  many as fit, measured in the cell's font).

## Where it stands

The generated fixtures, Excel 365 and Edge on Windows (runs in `..\runs\fixtures`; general.xlsx joined at run 4):

```
1b (before)  page counts 1/4 on same page 9027/85124 x within 1px   18/9027 (0.2%)   y within 1px 32/9027 (0.4%) both  3/9027 (0.0%)
2  (columns) page counts 3/4 on same page 9059/100147 x within 1px 9057/9059 (100.0%) y within 1px 51/9059 (0.6%) both 51/9059 (0.6%)
3  (rows)    page counts 4/4 on same page 100147/100147 x within 1px 100145/100147 (100.0%) y within 1px 100147/100147 (100.0%) both 100145/100147 (100.0%)
4  (General) page counts 5/5 on same page 100364/100364 x within 1px 100363/100364 (100.0%) y within 1px 100364/100364 (100.0%) both 100363/100364 (100.0%)
9  (fonts)   page counts 6/6 on same page 100423/100423 x within 1px 100418/100423 (100.0%) y within 1px 100405/100423 (100.0%) both 100402/100423 (100.0%) width within 1px 99390/100423 (99.0%)
14 (fills)   page counts 7/7 on same page 100436/100436 x within 1px 100431/100436 (100.0%) y within 1px 100420/100436 (100.0%) both 100417/100436 (100.0%) width within 1px 99403/100436 (99.0%)
14 shapes    fills and borders within 1px 122/153 (79.7%), colour only 26, missing 5
17 (align)   page counts 8/8 on same page 100465/100465 x within 1px 100459/100465 (100.0%) y within 1px 100447/100465 (100.0%) both 100443/100465 (100.0%) width within 1px 99420/100465 (99.0%)
21 (numfmt)  page counts 9/9 on same page 102048/102048 x within 1px 102042/102048 (100.0%) y within 1px 101963/102048 (99.9%) both 101959/102048 (99.9%) width within 1px 100924/102048 (98.9%)
27 (rows)    page counts 10/10 on same page 102094/102094 x within 1px 102082/102094 (100.0%) y within 1px 102067/102094 (100.0%) both 102063/102094 (100.0%) width within 1px 100961/102094 (98.9%)
30 (fonts)   page counts 10/10 on same page 102095/102095 x within 1px 102093/102095 (100.0%) y within 1px 102082/102095 (100.0%) both 102080/102095 (100.0%) width within 1px 100962/102095 (98.9%)
```

Run 3 repeated with no change (3b) gave the same report to the cell: no noise seen on the fixtures. fonts.xlsx joined
at run 5; text width is measured from run 5 (run 4 had 109 of 100,364 within 1px). Run 5 printed in 12pt by mistake
(the cell text size was set on the screen grid only, and print pages sit outside it): its baselines were right by
accident, and runs 6-8 show the real work.

From run 10 text is matched page by page (Excel's PDF keeps text clipped past a page's edge). From run 17 pieces a
space apart are joined (Chromium writes wrapped words as runs of their own); widths.xlsx shows 4 fewer matches for
it, its 1-2 character columns now reading as one piece with their neighbours, not a change in layout. numfmt.xlsx
joined at run 18: 85 of its baselines are 1.1px off, 4 rows a page at the same two places on every page, Chromium's
whole-pixel snapping landing just past the line. Next (GridCraft's `crates/numfmt` is a reference, checked
against Excel). All of this so far is on generated fixtures; the numbers that matter come from real workbooks
(`..\samples`).

## Open differences

- Text rotation, shrink-to-fit, conditional formats and East Asian date formats are not drawn yet.
- Row and column styles (a whole column's fill) are not applied to cells that aren't in the file.
- The maximum digit width is fixed at 7px (right for Calibri 11 and Aptos Narrow 11); other Normal fonts need theirs.
- Gridlines show under text that runs on into empty cells; Excel hides them there.
- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.
- Printing: no scaling or fit-to-page, print titles, manual page breaks, print area, headers and footers yet.

## Tried and dropped

(Nothing yet.)
