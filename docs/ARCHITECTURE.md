# Architecture

A browser viewer (and, to come, editor) for .xlsx files in plain JavaScript. Three layers that don't know about each
other's insides:

- `src/ooxml/` reads and writes the file. No DOM rendering; runs in Node too (unit tests use @xmldom/xmldom).
- `src/editor/` is the view: the grid, and later selection, editing and the ribbon.
- `src/engine/` will hold measured data (font metrics, Excel's layout constants). Empty so far.

## Module map

| Module | What it does |
|---|---|
| `src/index.js` | `mountXlsxEditor(host, { bytes, fileName, theme })`: opens a workbook and shows the sheet it opens on; `save()` gives the bytes back. Types in `src/index.d.ts`. |
| `src/ooxml/package.js` | The OPC package (the zip): every part kept byte for byte, relationships, content types; only changed parts rewritten. From open-docx-editor. |
| `src/ooxml/xml.js` | Parsing and serialising XML (injectable DOMParser for Node), namespace-aware lookups defaulting to SpreadsheetML, `_xHHHH_` decoding. From open-docx-editor. |
| `src/ooxml/units.js` | The one home for unit conversions: EMU, points, pixels, Excel's column width unit. From open-docx-editor, extended. |
| `src/ooxml/refs.js` | A1 references and ranges; rows and columns are 0-based in code. |
| `src/ooxml/workbook.js` | `Workbook`: sheets in order, shared strings, each sheet's model read on first use, `setCell`, `toBytes` (only dirty sheets serialised, with their XML declaration). |
| `src/ooxml/sheet.js` | A worksheet part to a sparse model: cells by row then column, row heights, column widths, merges, defaults. Implied `r` positions followed. |
| `src/ooxml/sheetedit.js` | Changing one cell in the worksheet's XML in place: implied positions pinned first, row and cell inserted in order, `spans` hint dropped when wrong, `dimension` grown. Formula cells refused until formulas and calcChain are modelled. |
| `src/ooxml/strings.js` | Shared string table and inline strings; phonetic runs skipped. |
| `src/ooxml/numfmt.js` | Display text of a value. Only the General format so far. |
| `src/editor/grid.js` | The virtualised grid: only rows and columns in view (plus overscan) in the DOM; sticky headings; text running on over empty cells; merged ranges; gridlines. An ARIA grid (rows, row and column headers, gridcells with indexes). |
| `src/editor/print.js` | The printed sheet: used range split into pages (columns across the printable width, rows down its height, the sheet's page order), each at its paper size with the sheet's margins; `@page` size set. Built by `preparePrint()` and on `beforeprint`. |
| `src/editor/axis.js` | Sizes and offsets of rows or columns; binary search from a scroll position to an index. |
| `src/editor/celltext.js` | The one place cell text size and line height are set (`--oxe-cell-fs`, `--oxe-cell-lh`). |
| `src/editor/editor.css` | Theme variables (light, dark, host-overridable) and the grid's layout. Geometry through custom properties; no `!important`. |

## Tools and tests

| File | What it does |
|---|---|
| `test/fixtures/make-fixtures.mjs` | Generates the test workbooks (deterministic zips). CI checks the committed files match. |
| `test/*.test.js` | Unit tests in Node (`npm test`): round trips byte for byte, edits, units, references, geometry, the checker. |
| `test/e2e.mjs`, `test/a11y.mjs` | Browser suites (`npm run test:e2e`): each fixture opened, scrolled, saved; axe in light and dark at 1280 and 320px. Both assert standards mode. |
| `tools/check-ooxml.mjs` | Structural check of .xlsx files (what makes Excel offer a repair). |
| `tools/excel-compare/` | Comparison with Excel's own PDFs: `excel-print.ps1` (Excel through COM, read-only), `print-ours.mjs` (the editor in Edge), `cells.cjs` (cell text and positions from both), `compare.py`, `diff.py`, `run.ps1`. See its README. |
| `tools/scan-bytes.mjs` | Control characters and byte order marks in source text (part of `npm run lint`). |

## How a file becomes a grid

1. `OpcPackage.fromBytes` unzips; the workbook is the target of the package's `officeDocument` relationship.
2. `Workbook` reads `xl/workbook.xml` (sheets, `date1904`) and the shared strings.
3. `wb.sheet(i)` parses that sheet's part once and keeps the XML document with the model.
4. `createGrid` turns the model into row and column axes (`sheetGeometry`), then draws the window the scroll
   position shows; scrolling redraws on the next animation frame when the window changes.
5. `save()` serialises only the sheets that `setCell` touched; every other part goes out as it came in.
