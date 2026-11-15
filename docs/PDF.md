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
skip the installed fonts. The served fonts' folder is the editor's own (`fonts: { omf }`).

## What is drawn

- **Pages:** each printed page at its paper size and orientation, with the sheet's margins.
- **Fills and borders:** where the page paints them, in the same order.
- **Cell text:** as real text, which can be selected and searched. Each line is placed at the layout's position and
  clipped to its cell, as the page clips it. Underline and strikethrough are drawn too.
- **Fonts:** for each piece of text, the first font of its cell's font stack that can be had: an installed font, read
  with the browser's Local Font Access (Chromium browsers ask once), or the open-metric-fonts twin the page serves
  (`fonts: { omf }`, README, Fonts), each embedded as a subset. Otherwise the PDF standard fonts are used. Each piece
  of text is scaled to the width the layout measured.

## Served fonts

A cell in a Microsoft font the computer lacks is drawn on screen in its open-metric-fonts twin (Morrow Sans Narrow for
Aptos Narrow, Tamarack for Tahoma...), and the PDF embeds the same twin (`src/pdfomf.js`). Its files are fetched from
the editor's `fonts.omf` folder, the ones the browser drew from, so most come from the browser's cache: each face's
WOFF2 file of each alphabet a piece of text uses, embedded as a subset. A piece of text with letters from several
alphabets (Latin and Cyrillic) is drawn in one run, each letter in the file that holds it, as the browser picks the
files by their unicode-range. A letter the twin lacks altogether (Seaford's twin has no Cyrillic, as Seaford hasn't) is
drawn as "?", as in a standard font. With `localFonts: false`, or where Local Font Access isn't allowed, the twin is
used for a font that is installed too: the same widths, and the font's own letter shapes rather than Helvetica's.

pdf-lib's fontkit reads WOFF2 and subsets it: Tamarack's (Tahoma's twin) latin file, for one line of text, gave a
5.2 KB PDF, its TrueType file 5.5 KB. The face's single whole WOFF2 file (also in the bundle) came out at 1.4 MB, so the
per-alphabet files are used. A face the family doesn't have (Tahoma has no italic; the browser slants its regular) is
drawn upright. A twin's text needs no scaling to the layout's width except where the browser kerned a pair the
standard widths don't (Aptos Narrow's "Yes" 15.63pt unkerned, 14.91 as laid out: drawn at 95.4%).

- **Headers and footers:** their text, as the cells', and their pictures (`&G`, PNG and JPEG) at the layout's places
  (docs/PRINT.md).

Print areas, titles, scaling and fit to page come from the print layout (docs/PRINT.md). Not drawn yet, as the print
layout doesn't have them either: charts, pictures on the sheet, and notes.

## Measured against Excel (10 and 15 November 2026)

Excel printed the ten fixtures to PDF (`excel-print.ps1`, kept in `excel-pdfs/fixtures` outside the repository). The
editor exported the same workbooks through `exportPdf`, run as `OXE_EXPORT=1 node tools/excel-compare/print-ours.mjs`
in Edge with the installed fonts allowed (Calibri, Cambria, Arial, Times New Roman and Courier New embedded), and on 15
November also with open-metric-fonts 1.0.1 served (`OXE_FONTS=1`: Aptos Narrow, in basic.xlsx and large.xlsx, embedded
as its twin Morrow Sans Narrow instead of Calibri). `cells.cjs`, `compare.py`, `shapes.cjs` and `compare-shapes.py`
then compared the two sets of PDFs, as for printing (docs/LAYOUT.md).

| | exportPdf, twins served | exportPdf | Browser print, twins served |
|---|---|---|---|
| Page counts the same as Excel's | 10 of 10 (large.xlsx: 342 of 342) | 10 of 10 | 10 of 10 |
| Cells on the same page, within 1px across and down | 102,080 of 102,098 (100.0%) | 102,079 of 102,097 (100.0%) | 102,080 of 102,095 (100.0%) |
| Text width within 1px | 101,989 of 102,098 (99.9%) | 96,968 of 102,097 (95.0%) | 101,981 of 102,095 (99.9%) |
| Fills and borders within 1px | 144 of 173 (83.2%) | 144 of 173 (83.2%) | 142 of 173 (82.1%) |

With the twin embedded, large.xlsx's 100,000 cells are all within 1px in width, as printed. The cells that differ are
in the same two workbooks as in printing, rows.xlsx and fonts.xlsx: 40 of 56 for fonts.xlsx (43 when printed) and 44
of 46 for rows.xlsx (45). Text widths come out a little wider or narrower than Excel's where the embedded font's own
widths differ from the browser's. Each piece of text starts at the same place either way.

Speed, for large.xlsx's 342 pages and 100,000 cells: 11 seconds with the standard fonts in Chrome, about 2.5 minutes
with the installed fonts embedded in Edge.

Tests: `test/pdf.mjs` covers the sheet shown (page count, paper size, title, real text) and the whole workbook (its
hidden sheet left out), and checks that the editor is left as it was. `test/omf.mjs` covers a served twin: fetched only
for a font the computer lacks, embedded from the folder, its text real text.
