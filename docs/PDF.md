# PDF export

`src/pdf.js` (`import { exportPdf } from 'open-xlsx-editor/pdf'`) makes a PDF of the workbook open in the editor
without the browser's print dialog. It uses the same print layout that printing uses (`print.js`), so the pages,
page breaks and paper size are the ones the browser would print. It uses pdf-lib and fontkit (MIT), adapted from
open-docx-editor's `pdf.js`.

```js
import { exportPdf } from 'open-xlsx-editor/pdf';
const bytes = await exportPdf(editor);                    // the sheet shown, as Excel's "Print Active Sheets"
const book = await exportPdf(editor, { sheets: 'all' });  // every visible worksheet, as "Print Entire Workbook"
```

Options: `sheets` (`'shown'` or `'all'`), `title` and `author` for the document information, and `localFonts: false` to
skip the installed fonts.

## What is drawn

- **Pages:** each printed page at its paper size and orientation, with the sheet's margins.
- **Fills and borders:** where the page paints them, in the same order.
- **Cell text:** as real text, which can be selected and searched. Each line is placed at the layout's position and
  clipped to its cell, as the page clips it. Underline and strikethrough are drawn too.
- **Fonts:** the installed fonts, embedded as subsets, read with the browser's Local Font Access (Chromium browsers ask
  once). Otherwise the PDF standard fonts are used, each piece of text scaled to the width the layout measured.

- **Headers and footers:** their text, as the cells', and their pictures (`&G`, PNG and JPEG) at the layout's places
  (docs/PRINT.md).

Print areas, titles, scaling and fit to page come from the print layout (docs/PRINT.md). Not drawn yet, as the print
layout doesn't have them either: charts, pictures on the sheet, and notes.

## Measured against Excel (10 November 2026)

Excel printed the ten fixtures to PDF (`excel-print.ps1`, kept in `excel-pdfs/fixtures` outside the repository). The
editor exported the same workbooks through `exportPdf`, run as `OXE_EXPORT=1 node tools/excel-compare/print-ours.mjs`
in Edge with the installed fonts allowed (Calibri, Cambria, Arial, Times New Roman and Courier New embedded).
`cells.cjs`, `compare.py`, `shapes.cjs` and `compare-shapes.py` then compared the two sets of PDFs, as for printing
(docs/LAYOUT.md).

| | exportPdf | Browser print (the run before) |
|---|---|---|
| Page counts the same as Excel's | 10 of 10 (large.xlsx: 342 of 342) | 10 of 10 |
| Cells on the same page, within 1px across and down | 102,079 of 102,097 (100.0%) | 102,080 of 102,095 (100.0%) |
| Text width within 1px | 96,968 of 102,097 (95.0%) | 100,962 of 102,095 (98.9%) |
| Fills and borders within 1px | 144 of 173 (83.2%) | 142 of 173 (82.1%) |

The cells that differ are in the same two workbooks as in printing, rows.xlsx and fonts.xlsx: 40 of 56 for fonts.xlsx
(43 when printed) and 44 of 46 for rows.xlsx (45). Text widths come out a little wider or narrower than Excel's where the embedded font's own widths
differ from the browser's. Each piece of text starts at the same place either way.

Speed, for large.xlsx's 342 pages and 100,000 cells: 11 seconds with the standard fonts in Chrome, about 2.5 minutes
with the installed fonts embedded in Edge.

Tests: `test/pdf.mjs` covers the sheet shown (page count, paper size, title, real text) and the whole workbook (its
hidden sheet left out), and checks that the editor is left as it was.
