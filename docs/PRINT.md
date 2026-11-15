# The print model

Excel's Page Setup, applied when a sheet is printed or exported to PDF (`src/editor/print.js`,
`src/ooxml/printarea.js`, `src/ooxml/headerfooter.js`). The page geometry underneath this (column widths, row heights,
fonts, padding, baselines) is in docs/LAYOUT.md.

| Setting | Where it's stored | Applied |
|---|---|---|
| Print area | `_xlnm.Print_Area` | Only that range is printed, instead of the used range. A print area of several ranges prints only its first. |
| Rows and columns to repeat | `_xlnm.Print_Titles` (`$1:$2`, `$A:$A`) | Printed at the top (left) of every page that starts after them, for that page's columns (rows), and taken off the room for the page's own rows (columns). |
| Manual page breaks | `rowBreaks`, `colBreaks` (`man="1"`) | A page starts at each. |
| Scale | `pageSetup scale` | Columns and rows print at the scale, each rounded to whole printer dots (600 dpi). |
| Fit to pages | `pageSetUpPr fitToPage`, `fitToWidth`, `fitToHeight` (0: as many as it takes) | The largest whole percentage, 10 to 100, at which the print area fits. |
| Header and footer | `headerFooter`; pictures in the VML drawing of `legacyDrawingHF` | Left, centre and right sections with Excel's codes (`&P`, `&N`, `&D`, `&T`, `&F`, `&A`, `&Z`, `&G` a picture, bold, italic, underline, strikethrough, font, size, colour), and different first and even pages. |

## Measured rules

The probes are `test/fixtures/print-area.xlsx`, `print-scale.xlsx`, `print-fit.xlsx` and `print-headers.xlsx`
(`make-fixtures.mjs`), plus `print-scale.xlsx` at 40, 50, 60, 70, 75, 80 and 90% (made from it for the measurement,
outside the repository). Excel printed each to PDF, and the text of both sets of PDFs was compared as for the layout
(`tools/excel-compare`).

- **Scaled sizes.** A row prints as its printed height in printer dots times the scale, rounded to a whole dot. A
  15pt Calibri row is 121 dots, so 48, 61, 73, 85, 91, 97 and 109 dots at those seven scales. Columns do the same:
  430 dots at 100% is 323 at 75%. The page's cells are drawn in a box zoomed by the scale, with these exact sizes
  before the zoom.
- **The room on a page.** Excel paginates into the printable height and width less one point. At 60%, 77 rows of 73
  dots fit on a 9.5in page, though 78 would fill only 5,694 of its 5,700 dots. At 90%, 10 columns of 387 dots fit
  across 7.1in, where 11 would fill 4,257 of 4,260. At 100%, 47 Calibri rows of 121 dots fit.
- **Fit to page.** Excel takes the whole percentage below the exact one: 24 columns across one page is 41.28%, and
  Excel prints at 41%.
- **Text in scaled cells** sits further in and lower than the scale alone puts it: 1.4px × (1 − scale) across (the
  3.2px left padding becomes 1.4 + 1.8 × scale px) and 1.2px × (1 − scale) down. In the zoomed box the browser draws
  baselines up to a pixel higher than placed, half a pixel on average, so half a pixel more is added.
- **Center on page.** Each page's cells (its print titles with them) sit in the middle of the room between its
  margins, a point short of it each way: centered horizontally, 1, 3 and 8 columns, portrait and landscape, every cell
  1.2 to 1.3px left of the middle; vertically the same above it (workbooks made in code, printed by Excel;
  `..\runs\center` and `center2`: 252 of 252 cells within 1px across and down). Read from printOptions'
  horizontalCentered and verticalCentered.
- **Printed gridlines** (Sheet Options > Gridlines: Print, printOptions gridLines) are hairlines (0.19px in Excel's
  PDF) at every row and column edge of each printed block of cells, print titles included; Excel draws an outline
  round the printed range too, over the same edges (a workbook made in code, `..\runs\gridprint`: all 13 of its
  gridlines within 1px, the text where it was).
- **Printed headings** (Sheet Options > Headings: Print, printOptions headings): the column letters in a row of the
  default height above the cells, the row numbers in a column 4.96px plus 7.52px a digit wide (three digits at least:
  27.52px to row 999, 35.04px past it) left of them, both taken from every page's room; each label centred, 1.4px
  right of and 1.1px below the middle of its box, in the Normal style; rules of 1.28px round them and an outline
  round it all (workbooks made in code, 9 to 1,200 rows, `..\runs\heads`: 864 of 864 text pieces within 1px both
  ways, 299 of 299 rules, the same page counts).
- **A print area of several ranges** prints each on pages of its own, in the order the name lists them, each from
  the top-left of its page, print titles over each (workbooks made in code: two ranges down, two across, three out of
  order, two with a title row; `..\runs\areas`: 116 of 116 text pieces within 1px both ways, the same page
  counts). With fit to page, the scale is the first range's here (not measured).
- **Print titles** print only for the page's own columns. With a print area of B3:H70, rows 1:2 print as B1:E2 above
  the B:E pages.
- **Headers and footers.** They use the Normal style's font and size, scaled with the sheet unless `scaleWithDoc` is
  off. A section's baseline is 1.2px + 1.023em below the header margin, or 2.27px + 0.295em above the footer margin,
  in the section's largest size (16pt in one section moves that section alone). The left section starts 1.25px inside
  the left margin, the right section ends 1.25px inside the right margin, and the centre is centred on the page. (At
  100% that is 16.2px and 6.6px at 11pt; the same at 40, 50, 60, 75 and 90%: workbooks made in code, `..\runs\hfscale`,
  292 of 292 text pieces within 1px both ways, where the rule measured at 100% alone, 1.05em and 0.436em, left 11 of
  them 1.1 to 3.2px out at the smaller scales.)
- **Headers and footers of several lines** (a line break in a section): the lines a printed row of their size apart
  (20.3px at 11pt, 25.4px at 14pt); a header's first baseline 1.014em below the header margin, a footer's last 0.361em
  above the footer margin and the lines before it above; each line centred in a centre section, right-aligned in a
  right one; a left or right section of several lines starts at the margin itself, not 1.25px in (workbooks made in
  code, one to three lines, `..\runs\hf2`: 198 of 198 text pieces within 1px both ways, single-line headers as
  they were). The same at a scale, the row and the em scaled with the sheet: at 40, 50, 60, 75 and 90%, and 14pt at
  50% (`..\runs\hfscale2`: 279 of 279 text pieces within 1px both ways; two more present in both PDFs left unpaired by
  the comparison's reading order).
- **Different first page and different odd and even pages** (headerFooter differentFirst, differentOddEven): the
  first page takes firstHeader and firstFooter, even pages evenHeader and evenFooter, placed as any page's (workbooks
  made in code, three pages each way and both, `..\runs\hf3`: 1,187 of 1,187 text pieces within 1px both ways).
- **Pictures in headers and footers** (`&G`; `ooxml/hfpictures.js`): each section's picture is a shape of the VML
  drawing that legacyDrawingHF names, its id the section's (LH, CH, RH, LF, CF, RF, with FIRST or EVEN after it for
  the first and even pages: LHFIRST, RFEVEN), its style the size in points. Excel prints it 1.075 times as wide and
  1.015 times as tall (60 by 30pt as 85.92 by 40.64px), 1.92px wider again for text before it and for text after it.
  Alone in its section, its top is on the header margin, or its bottom on the footer margin, at the margin itself
  (no 1.25px inset) or centred. Beside text, its bottom is 0.183em below the text's baseline, and the line that much
  further from the margin (2.7px at 11pt). The other sections stay where they were (workbooks made in code with a PNG
  made in code, alone, beside text, first and even pages, `..\runs\hfpic`: 532 of 532 text pieces within 1px
  both ways, browser print and PDF export; all 9 pictures (`tools/excel-compare/images.cjs`) within 0.5px in place
  and size, Chrome's print rounding them to whole pixels, PDF export within 0.15px).
- **Pictures in sections of several lines, and at a scale** (`editor/printhf.js`): the section's lines are boxes one
  under the other from the header margin (up from the footer margin), at the margin itself. A text line's baseline is
  1.014em below its box's top and its box a printed row tall, as in any section of several lines; a picture's line has
  its baseline the picture's height below its top. The picture's bottom is on that baseline, or 0.183em below it when
  its line has any characters, its line break included; the box then goes on below the baseline as a text line's does,
  and it ends at the baseline when the line has none. The 1.92px for text beside a picture counts the text on its own
  line. With the sheet printed at a scale, pictures, the drop and the 1.92px scale with it when the header scales with
  the sheet (half at 50%) and not when scaleWithDoc is off (workbooks made in code, `..\runs\hfpic2`: two and three
  lines, a picture above and below text and between words, at 50% and at 75% with scaleWithDoc off; all 9 pictures
  within 0.5px in place and size, print and export; text pieces within 1px both ways, 122 of 122 since the one-line
  rule at a scale above).

## Setting them (measured)

Page Layout > Page Setup's Print Area and Breaks set what the rules above read (`ooxml/pagesetupedit.js`), written as
Excel writes them. Measured over COM in a hidden Excel: the ribbon's Insert Page Break, Remove Page Break and Reset
All Page Breaks through ExecuteMso (its Print Area commands don't run in a hidden Excel, so PageSetup.PrintArea, which
sets the same name), then the file saved (`test/pagesetupedit.test.js`).

| Excel's rule | Excel's file |
|---|---|
| Insert Page Break at the active cell: a manual break before its row and one before its column; in row 1 only the column's, in column A only the row's, at A1 none. | C5: `<rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks>` and `<colBreaks ...><brk id="2" max="1048575" man="1"/></colBreaks>`. |
| Remove Page Break: the break above the active row and the one left of its column, those there. | Breaks at rows 5, 9 and columns C, E; at E3: E's alone goes. |
| Reset All Page Breaks: every manual break, the elements too. | No rowBreaks or colBreaks left. |
| Margins: all six in pageMargins, inches; Excel's presets Normal (0.7 sides, 0.75 top and bottom, 0.3 header and footer), Wide (1, 0.5), Narrow (0.25 sides). | `<pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/>`. |
| Orientation and Size: orientation always written once pageSetup is, paperSize unless Letter (1). | `<pageSetup paperSize="9" orientation="portrait"/>` for A4; `orientation="landscape"`. |
| Scale to Fit: fitting is `fitToPage="1"` in sheetPr's pageSetUpPr, fitToWidth and fitToHeight written unless 1, Automatic as 0; a scale takes fit to pages off, sheetPr with it when left empty. | Width 1, Height Automatic: `fitToHeight="0"`; Width Automatic, Height 2: `fitToWidth="0" fitToHeight="2"`; 75%: `scale="75"`. |
| Custom margins (Margins > Custom Margins..., Excel's Page Setup box at its Margins tab) written as typed; Center on page in printOptions, before its headings and gridLines; Page order (the Sheet tab) as pageOrder, left out for Down, then over. | `<pageMargins left="0.6" right="1.25" top="0.33" bottom="2" header="0.1" footer="0"/>`, `<printOptions horizontalCentered="1" verticalCentered="1"/>`, `pageOrder="overThenDown"`. |
| Print Titles (Page Layout > Print Titles, Excel's Page Setup box at its Sheet tab): the sheet's own name _xlnm.Print_Titles, columns first, then rows; rows typed "1:2" or a range's rows; a lone row refused; neither, the name gone. | `'Q1 Plan'!$A:$B,'Q1 Plan'!$1:$2`; "B2:C4" for rows: `$2:$4`; "$3": refused. |
| Set Print Area: the selection, the sheet's own name _xlnm.Print_Area, $-absolute with the sheet's name; Add to Print Area (offered once there is one) adds the selection after a comma; Clear Print Area takes the name away. | `'My Data'!$B$2:$D$5`, `'My Data'!$B$2:$D$5,'My Data'!$F$7:$G$8`, one cell `'My Data'!$C$3`. |

Excel also writes its printer's own horizontalDpi, verticalDpi and a printerSettings part when page setup is touched;
the editor leaves those to the printer. Each is one step of the undo history. Not measured: Excel's joining of neighbouring ranges in Add to Print Area (here
they are listed as given). The Page
Setup box's messages for what it won't take are the editor's own, and only its Sheet tab's print area and titles are
there.

## Results (10 November 2026)

| | Browser print | PDF export |
|---|---|---|
| The four probes: texts within 1px across and down | 3,872 of 3,872 | 3,872 of 3,872 |
| The probe at seven scales | 8,820 of 8,820, page counts 7 of 7 | |
| The ten earlier fixtures | 102,078 of 102,093, page counts 10 of 10, fills and borders 142 of 173 (as before) | |

The ten earlier fixtures were checked so the print model didn't change them. The 15 cells that differ are the same
ones as before it (docs/LAYOUT.md).
