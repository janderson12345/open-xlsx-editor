# Layout: matching Excel

Every rule here is to be measured against **Excel's own PDFs** (`tools/excel-compare`), never another program's
output. A change is kept only if the total doesn't get worse and every workbook that got worse is
understood. Rules marked *measured* were fitted on Excel's PDFs; the rest are starting points from the specification.

The probe for printed geometry is `test/fixtures/widths.xlsx`: 40 column widths (text in row 1 at each column's left
edge, a number in row 2 at its right edge) and 20 row heights (text in column A).

## Rules in use

- **Column width** (*measured*). Stored width `w` in characters of the maximum digit width `mdw`, padding included:
  `px = trunc(((256 w + trunc(128 / mdw)) / 256) mdw)` (ECMA-376 Part 1, 18.3.1.13), `src/ooxml/units.js`.
- **Maximum digit width** (*measured*, `tools/excel-compare/excel-digitwidth.ps1`, `test/digitwidth-excel.json`). The
  workbook's Normal style font sets it (`workbook.mdw`, `src/ooxml/digitwidth.js`): Excel 365 over COM, the Normal
  style given 41 fonts and sizes, columns of 1, 2, 5, 10 and 20 characters read back in points and as saved, the
  default width and AutoFit. Calibri 8-9: 6, 10-11: 7, 12: 8, 14: 10, 16: 11, 18: 12, 20: 14, 24: 16; Aptos Narrow 11:
  7; Aptos 10, 11, 12: 7, 8, 9; Arial 8 to 16: 6, 7, 7, 8, 9, 11, 12; Times New Roman 10, 11, 12, 14: 6, 7, 8, 9;
  Cambria 10, 11, 12: 7, 8, 9; Segoe UI 9, 10, 11: 6, 7, 8; Verdana 8, 10, 12: 7, 8, 10; Courier New 10, 11, 12: 8, 9,
  10 (and Tahoma, Georgia, Consolas, MS Sans Serif). The padding grows with it: 2 x ceil(mdw / 4) + 1 pixels (5 up to
  an mdw of 8, 7 to 12, 9 to 16), so n characters are n x mdw + padding, stored as that over mdw cut to 1/256 (Arial
  12's 10 characters: 97px, `10.77734375`), and the default column 8 characters plus padding rounded up to a multiple
  of 8 (56px for an mdw of 6, 80 for 9, 128 for 14, 144 for 16). AutoFit pads its text by the padding and 2px more.
  All 41 x (5 widths, stored widths, default width, StandardWidth, AutoFit) as Excel has them (`npm test`). Excel's
  widths come from Windows' hinted glyphs: worked out from the font files alone (the digits' advance at the size in
  pixels, rounded), 32 of the 36 installed fonts are right and Calibri 8 and 14, Arial 14 and Times New Roman 10 a
  pixel off; so the measured table comes first and any other font is measured in the browser (a canvas, 0-9, rounded
  to the nearest pixel, `src/editor/digitmeasure.js`). Grid, column headings, typed and dragged widths, AutoFit, the
  Width box and tip, CELL("width"), pasted column widths, print (screen pixels x 1.075) all count in it. Browser check:
  `test/display-diffs.mjs`, the headings of seven fonts (mdw 6 to 16) at Excel's pixels.
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
- **Gridlines under text running on** (*measured* in print, `tools/excel-compare/excel-runover.ps1`: Excel's PDF with
  Gridlines: Print, `test/runover-excel.json`). Text running on into empty cells leaves out the column gridlines it
  crosses in its row: left text the ones to its right, right text to its left, centred both ways, stopped where the
  text stops (a neighbour with something in it, or the text's own end: nine zeros cross an 8.43 column's edge, eight
  don't); the row's own gridline stays. Measured with the text's width in the browser (`cells.js runGaps`): on screen
  21 of 21 left out and the other 210 drawn (read off a screenshot) and in print the same as Excel's PDF
  (`test/display-diffs.mjs`).
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
  distributed are drawn left.
- **Turned, stacked and shrunk text** (*measured* in print, `test/fixtures/rotate.xlsx`: 37 samples, Excel's PDF in
  `..\excel-pdfs
otate`, runs in `..
uns
otate`). General alignment of turned text goes by the angle, not the
  value: 1 to 89 degrees up on the left, 90 up and every angle down on the right, 90 down on the left (a number
  too); stacked letters are centred. Stacked letters are a line each, stepping by the font's default row height as
  wrapped lines do (19.36px for Calibri 11; 17.96 with the vertical writing mode used before), each centred, and set
  left or right each centred in a column as wide as a line is tall. Text turned short of 90 degrees keeps 0.23 em more
  from the cell's sides than its line box turned would (1.7px at 89, 1.1 at 45, 0.5 at 15); at the bottom it sits
  0.18 em times the angle's cosine lower and 0.04 em higher, at the top 0.14 em lower (`cells.js` placeTurned).
  Shrink to fit sizes the text for the room less 1.7px and down to a whole printer pixel at 600 dpi (Excel printed
  11pt's 92 dots as 76, 67, 75 and 28 for four texts), keeping the full-size baseline. A row Excel sizes for text
  turned 90 degrees is its width times 1.0645 plus 7px (130.5pt for a 26-letter line; 1.14 times gave 134.25).
  Runs 1 -> 8: rotate.xlsx within 1px both ways 11/67 -> 64/68 (x 34 -> 67, y 18 -> 65), width 58 -> 63. Left:
  'Row sized at 45' (Excel's row 63pt, here 66: the rows below it 4.6px lower), down 89 degrees 1.4px low, a number
  turned down 45 1.1px low, down 90 set right 1.6px right, the long shrunk text 1.3px right. On screen the same rules
  are used but not measured, and shrink to fit isn't stepped there. PDF export (src/pdf.js) draws turned text flat.
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
- **East Asian built-in dates** (*measured*, `tools/excel-compare/excel-eastasia.ps1`, `test/eastasia-excel.json`).
  Built-in ids 27-36 and 50-58 mean a different code in each East Asian locale and a file gives only the id. Excel 365
  en-US reads them as m/d/yyyy (the short date, as format 14, so following the locale) and 32-35 as h:mm:ss (266 cells
  over COM); in ja-JP, zh-CN, zh-TW (and zh-HK, zh-MO) and ko-KR they are ECMA-376's table of codes (18.8.30;
  `src/ooxml/eastasia.js`): Windows' Region can't be changed for Excel alone, so which code each id takes in those
  locales is the specification's, and how Excel draws each code was measured (840 cells): `[$-411]` codes count
  Japanese eras (g R/H/S/T/M, gg 令/平/昭/大/明, ggg 令和..., e the era's year, ee two digits; 1989-01-07 is 昭和64,
  1989-01-08 平成1, 2019-05-01 令和1, 元 only with `[$-ja-JP-x-gannen]`), `[$-404]e` is the Gregorian year (no
  Republic of China calendar without its calendar id), 上午/下午 is AM/PM, other locale tags change nothing. `e` alone is
  the year and `g` nothing, as in Excel. The locale is the editor's (`docs/LOCALE.md`).

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

## Pictures, shapes and icons

Read from the sheet's drawing parts (`src/ooxml/drawobjects.js`), painted as paths, pictures and text boxes in the
sheet's pixels (`src/editor/drawpaint.js`, the outlines in `src/ooxml/presets.js`) and drawn over the cells
(`src/editor/drawview.js`), with the charts, in the drawing's order. The same boxes are printed (`print.js`) and
exported (`pdf.js`). The probe is `test/fixtures/drawings.xlsx` (`drawings.mjs`, everything made in code): Excel
printed it to PDF (`tools/excel-compare/excel-print.ps1`) and its pictures (`images.cjs`), fills and lines
(`shapes.cjs`) and text (`cells.cjs`) were compared with ours.

- **Anchors.** A two-cell anchor goes from its first corner to its second whatever its editAs: Excel's PDF drew the
  editAs="oneCell" picture over hidden row 6 that much shorter (116.16px, six printed rows), though its extent is
  seven rows. A one-cell anchor is its corner and its extent; an absolute anchor its position and extent, in pixels
  of the sheet (at print, not scaled with the printed columns: the 120px picture printed 120px). Offsets past a
  cell's size stop at its edge. Lines (connectors, line) may be 0 one way; anything else of no width or height
  (all its rows or columns hidden) isn't drawn.
- **Turned shapes.** Turned 45 to 135 degrees (or 225 to 315) the anchor holds the shape's frame upright, its width
  and height swapped about the middle: the right arrow at 45 degrees over B18:D22 printed 184.6px square, as an arrow
  of the anchor's height by its width turned is; not swapped, 201.7px. Under 45 degrees the frame is the anchor's.
  Flips mirror the outline and the picture; text isn't mirrored (turned over for flipV).
- **Pictures.** srcRect crops each side by its fraction of the picture (negative: a margin); Excel's PDF has the
  cropped part at the frame, ours the whole picture placed so that the same part fills the frame, clipped to it: the
  same picture on the page within 0.4px. A picture's geometry (a rounded or oval frame) clips it. SVG icons draw
  from their SVG on screen and in browser print, and from their PNG fallback in the PDF export; GIF, BMP and WebP
  are turned to PNG for the export. EMF, WMF and TIFF draw as a grey box.
- **Colours.** The file's own, else the theme's through the shape's style (fillRef, lnRef: the theme's line widths,
  fontRef for text). lumMod and lumOff work on HSL's luminance in fractions, not the integer HLS of cells' tints:
  accent 2 (ED7D31) at 75% printed C55A11 (integer HLS gives C65911). shade and tint work in linear RGB: accent 1
  shaded 50% printed 2F528F, as here. alpha is kept; a gradient is drawn in its middle stop's colour.
- **Arrowheads.** Medium heads (Excel's default) are three times the line's width long and wide, and no less than 8px:
  on lines of 1.33 and 2px Excel printed 8 by 8px (here within 0.3px).
- **Text.** Insets 0.1in across and 0.05in down (bodyPr's lIns, tIns...), anchored top, middle or bottom, paragraphs
  left, centred or right, lines the font's own height apart (18px for 11pt Calibri at 100% spacing, not 1.2em); 11pt
  in the theme's body font when the file says nothing.

Results (16 November 2026, Excel 365 and Edge): pictures within 0.8px in place and size in browser print, within 0.6px
in the PDF export; the shapes' text 40 of 40 pieces within 1px across and down, both ways; fills 12 of 18 within 1px
in browser print (13 in the export), the others Excel's lines written as filled outlines where ours are strokes (their
ends within 1.1px), its SVG icon drawn as paths, and the cloud.

Not drawn yet, or drawn otherwise: presets other than those in `presets.js` (the cloud, callouts, stars but star5,
brackets...) are their bounding box; custom geometry's guides are taken as 0; gradients, pattern fills, a shape's
picture fill, shadows, glow and 3-D are left out; vertical text is turned but not laid out in columns; text in a turned
box isn't in the PDF export (it is in browser print); SmartArt, slicers, ink and form controls aren't drawn.

## Open differences

- The grid ends 100 rows and 10 columns past the used range; Excel's goes on to row 1,048,576 and column XFD.
- Printing (print area, print titles, scale, fit to pages, page breaks, headers and footers): docs/PRINT.md.

## Tried and dropped

(Nothing yet.)
