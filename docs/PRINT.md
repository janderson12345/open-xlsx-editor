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
| Header and footer | `headerFooter` | Left, centre and right sections with Excel's codes (`&P`, `&N`, `&D`, `&T`, `&F`, `&A`, `&Z`, bold, italic, underline, strikethrough, font, size, colour), and different first and even pages. |

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
- **Print titles** print only for the page's own columns. With a print area of B3:H70, rows 1:2 print as B1:E2 above
  the B:E pages.
- **Headers and footers.** They use the Normal style's font and size, scaled with the sheet unless `scaleWithDoc` is
  off. The header's baseline is 1.05em below the header margin, and the footer's 0.436em above the footer margin. The
  left section starts 1.25px inside the left margin, the right section ends 1.25px inside the right margin, and the
  centre is centred on the page.

## Results (10 November 2026)

| | Browser print | PDF export |
|---|---|---|
| The four probes: texts within 1px across and down | 3,872 of 3,872 | 3,872 of 3,872 |
| The probe at seven scales | 8,820 of 8,820, page counts 7 of 7 | |
| The ten earlier fixtures | 102,078 of 102,093, page counts 10 of 10, fills and borders 142 of 173 (as before) | |

The ten earlier fixtures were checked so the print model didn't change them. The 15 cells that differ are the same
ones as before it (docs/LAYOUT.md).

Not yet: a print area of several ranges (only the first prints), "Center on page", gridlines and row and column
headings in print, pictures in headers (`&G`), headers of more than one line, and the first-page and even-page
options beyond choosing their text (measured only with odd pages).
