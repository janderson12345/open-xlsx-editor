# open-xlsx-editor

Open, edit and save Excel workbooks (.xlsx) in the browser. Plain JavaScript, no server, no Office.

[![Latest release](https://img.shields.io/github/v/release/janderson12345/open-xlsx-editor)](https://github.com/janderson12345/open-xlsx-editor/releases)
[![npm](https://img.shields.io/npm/v/open-xlsx-editor)](https://www.npmjs.com/package/open-xlsx-editor)
[![CI](https://github.com/janderson12345/open-xlsx-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/janderson12345/open-xlsx-editor/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

**[Live demo](https://janderson12345.github.io/open-xlsx-editor/)**: a sample workbook opens; drop your own .xlsx on
the page (it stays in your browser, nothing is uploaded).

- **Looks like Excel.** The grid is laid out as Excel lays it out. Column widths, row heights, fonts, fills, borders,
  number formats and conditional formatting are drawn as Excel draws them, and printing and PDF export follow Excel's
  page setup. All of this is measured against Excel's own PDFs.
- **Calculates like Excel.** 500 of Excel's 519 functions, dynamic arrays that spill as Excel 365's do, LET and
  LAMBDA, table references and 3D references. 1926 of 1926 test formulas give Excel's results.
- **Edits like Excel.** Excel's ribbon (classic or simplified), its keys and its dialog boxes: Format Cells, Find and
  Replace, Sort, Filter, Fill, Paste Special, Conditional Formatting, Data Validation, PivotTables, charts, notes and
  comments, hyperlinks, sheet and workbook protection, password-protected workbooks, Text Import and Save As.
- **Speaks your locale.** Numbers, dates and formulas are shown and typed with the reader's separators and date order
  (1.234,50, 05.03.2024, `=ROUND(1,5;0)` in Germany), and the file stays as Excel writes it. CSV and text files open
  and save in UTF-8, UTF-16 or Windows-1252 with any separator.
- **Keeps your file.** Saving rewrites only the parts you changed. Every other part, including what the editor doesn't
  understand (macros, add-in data, newer features), goes back byte for byte, and the result opens in Excel without
  repair.
- **Works with AI assistants.** An MCP server (`open-xlsx-editor-mcp`) and an agent interface read and edit workbooks.
  Every change an assistant makes can be reviewed cell by cell and taken back before saving (docs/AGENT.md).
- **Runs anywhere.** One script and one stylesheet, served from your own site. It fetches nothing while it runs, and
  works under a strict Content-Security-Policy. Light and dark themes; keyboard and screen reader support, checked
  with axe.

This is version 0.x: the API may still change between minor versions. `docs/FEATURES.md` lists, ribbon tab by ribbon
tab, what Excel does and what this editor does yet. The main things missing are:

- opening .xlsb and .xls files;
- running macros;
- PivotTable filters, grouping and calculated fields;
- the Spelling box;
- Check Accessibility;
- split windows;
- 19 functions, including FORECAST.ETS.

## Install

```sh
npm install open-xlsx-editor
```

## Quick start

```js
import { mountXlsxEditor } from 'open-xlsx-editor';
import 'open-xlsx-editor/style.css';

const bytes = new Uint8Array(await (await fetch('book.xlsx')).arrayBuffer());
const editor = mountXlsxEditor(document.getElementById('host'), { bytes, fileName: 'book.xlsx' });

const saved = editor.save(); // a Uint8Array: the .xlsx with your changes
```

The host element needs a height; the editor fills it. Without a bundler, use the prebuilt files in `dist/`
(`open-xlsx-editor.js`, an ES module, and `open-xlsx-editor.css`):

```html
<link rel="stylesheet" href="open-xlsx-editor.css">
<div id="host" style="height: 100vh"></div>
<script type="module">
  import { mountXlsxEditor } from './open-xlsx-editor.js';
  const editor = mountXlsxEditor(document.getElementById('host'));
</script>
```

Without `bytes`, `editor.open(bytes, fileName)` opens a workbook later. A workbook saved with a password asks for
it.

## The API

Everything the ribbon and the keys do is a named command, so a page can drive the editor too. The full API, with
types, is in `src/index.d.ts`.

```js
editor.enter('B2', '42');      // as if typed; editor.setCell('C3', 'text'), editor.undo(), editor.select('A1')
editor.enter('B3', '=B2*2');   // a formula; the cells that read B2 or B3 are worked out again
editor.showSheet('Notes');     // another sheet, by name or index, with the selection it had
editor.renameSheet('Notes', 'Q1 notes'); // formulas that read it follow; also insertSheet, moveSheet, hideSheet...
editor.command('toggleFont', 'b'); // what the ribbon's buttons do, on the selection; Ctrl+B does the same
editor.save({ as: 'xltx' });   // as a template, or a template as 'xlsx'
editor.saveCsv({ sep: ';', encoding: 'windows-1252' }); // the sheet shown as a CSV or text file (docs/CSV.md)
editor.saveAs();               // the Save As box (F12): .xlsx, .xlsm, .csv, .txt, .pdf; the file to onSave or downloaded
editor.importText(bytes, 'data.txt'); // the Text Import box, into a new workbook ({ into: 'sheet' }: at the active cell)
editor.setLocale('de-DE');     // numbers, dates and formulas shown and typed as in Germany (docs/LOCALE.md)
```

Without a page, in Node or a worker, `openWorkbook(bytes)` reads and edits a workbook through the agent interface
(docs/AGENT.md), and `Workbook.fromBytes(bytes)` gives the model the editor uses.

Two more entry points load only when imported:

- `open-xlsx-editor/pdf`: `exportPdf(editor)` makes a PDF of the workbook as Excel prints it (docs/PDF.md).
- `open-xlsx-editor/collab`: `connectWorkbook(editor, ydoc)` co-edits a workbook over Yjs (docs/COLLAB.md).

### Options

```js
const editor = mountXlsxEditor(el, {
  readOnly: true,                       // show the workbook only (no ribbon, nothing editable)
  features: { number: false },          // hide a tab by name, or a group by its label on any tab
  toolbar: true,                        // false hides the ribbon but keeps editing
  ribbon: 'simple',                     // start with the one-line ribbon ('classic': the full one); setRibbon() later
  statusBar: true,                      // false hides Average, Count and Sum; { minimum: true, maximum: true } adds those
  theme: 'dark',                        // 'light' or 'dark'; without it, the page's colour scheme
  fillLists: [['North', 'East', 'South', 'West']], // custom lists the fill handle goes round
  errorChecking: false,                 // no green error triangles; { rules: [...] } chooses Excel's rules
  onLink: (url) => window.open(url),    // a hyperlink clicked
  author: 'Avery Chen',                 // the name notes and comments are written under
  people: [{ name: 'Sam Lee', userId: 'sam.lee@example.com' }], // whom @ offers in a comment, besides the workbook's
  locale: 'de-DE',                      // country settings for numbers and dates; without it, the browser's language
  onSave: ({ fileName, bytes }) => {},  // Save As's file; without it, the browser downloads it
  onRefreshQuery: async ({ name, kind, connection, m, target }) => rows, // Data > Refresh: a query's rows (docs/QUERIES.md)
  exportPdf,                            // from 'open-xlsx-editor/pdf': Save As then offers PDF
  onChange: () => {},                   // the workbook changed
});
editor.setReadOnly(false);              // switch at any time
editor.setFeatures({ number: true });
```

Tab names: `home`, `insert`, `page-layout`, `formulas`, `data`, `review`, `view`. Group names are their labels in lower
case with dashes: `undo`, `clipboard`, `font`, `alignment`, `number`, `cells`, `editing`; on Page Layout,
`sheet-options`; on View, `show`, `zoom` and `window`. A group turned off also turns off its keys (`font: false`:
Ctrl+B does nothing; `clipboard: false`: Ctrl+C, X and V). F6 moves between the sheet and the ribbon, and Esc goes
back to the sheet.

## AI assistants (MCP)

```json
{ "mcpServers": { "excel": { "command": "npx", "args": ["-y", "-p", "open-xlsx-editor", "open-xlsx-editor-mcp", "--root", "/path/to/workbooks"] } } }
```

The server reads and edits the workbooks in one folder: sheets, cells, formulas, names, tables and notes. It saves to
a new file unless asked to overwrite, and the assistant shows its changes for review before saving. docs/AGENT.md has the tools and the review.

## Self-hosting: no CDN, a strict Content-Security-Policy

The editor is one script and one stylesheet, served from your own site. It fetches nothing while it runs (no CDN, no
font service, nothing sent anywhere), and writes its own styles through the CSSOM rather than as `<style>` elements
or `style=""` attributes. So it runs under a policy that allows nothing but your own origin, with no `'unsafe-inline'`
and no outside hosts:

```http
Content-Security-Policy: default-src 'none'; script-src 'self'; style-src 'self'; font-src 'self'; img-src 'self'; connect-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'
```

`frame-ancestors 'none'` refuses all framing; to show the editor in a frame of your own site, make it `'self'`.

This was measured in Edge with exactly this policy on every response. A host page with nothing inline opened
basic.xlsx, anchors.xlsx, charts.xlsx, notes.xlsx, print-headers.xlsx and formulas.xlsx, then edited, opened Format
Cells and Find, printed and saved. There were no blocked requests and none to another host. The page itself must also have
nothing inline (this repository's demo has an inline `<style>`, so it is not that page). Chromium logs one report
that blocks nothing: a `style-src-attr` report when the editor reads a workbook's own XML (a note's VML has `style=""`
attributes), whose values still come through.

### Fonts

Workbooks name Office fonts, and a browser draws only the fonts installed on the computer or served by the page.
Aptos Narrow, the body font of Excel 365's default theme, isn't installed with Windows, macOS or Linux, so without
help it is drawn in Calibri (or Carlito), which is close but not the same width; Tahoma, Verdana, Century Gothic,
Garamond and the other Microsoft fonts are missing on most computers that aren't Windows ones. [open-metric-fonts](https://github.com/open-metric-fonts/open-metric-fonts)
makes open (SIL OFL) fonts with exactly their character widths and line metrics, 151 families in release 1.0.1:
Morrow Sans Narrow for Aptos Narrow, Tamarack for Tahoma, Glade for Verdana, Compass for Century Gothic, Colines for
Garamond and so on. The editor loads them from a folder your site serves, the same folder the Word editor
(open-docx-editor) uses, so one copy serves every viewer of Office documents on the site:

1. **Serve the open-metric-fonts web bundle once, in a folder named for its version.** The editor is made for release
   `OMF_VERSION` (exported by the package; 1.0.1 now). The release's web bundle holds every package as
   `<package>/files/<face>-<alphabet>.woff2`, as the editor looks for them (about 70 MB on disk, every family and
   alphabet; a browser downloads only the files a workbook needs).

   ```bash
   curl -LO https://github.com/open-metric-fonts/open-metric-fonts/releases/download/v1.0.1/open-metric-fonts-web-1.0.1.zip
   ```

   Check it against the release's `SHA256SUMS.txt` and its build attestation before use:

   ```bash
   sha256sum -c SHA256SUMS.txt --ignore-missing
   ```

   ```bash
   gh attestation verify open-metric-fonts-web-1.0.1.zip --repo open-metric-fonts/open-metric-fonts
   ```

   ```bash
   unzip -q open-metric-fonts-web-1.0.1.zip && mkdir -p wwwroot/lib/fonts/omf/1.0.1 && cp -R open-metric-fonts-web-1.0.1/. wwwroot/lib/fonts/omf/1.0.1/
   ```

   Or only the families you want, from npm with your other dependencies, pinned (`npm install --save-exact
   @open-metric-fonts/morrow-sans@1.0.1 @open-metric-fonts/tamarack@1.0.1`; each font's package is named in
   `src/editor/omf.js`), each copied to `<folder>/<package>/` by your build; a font whose package isn't there falls
   back as before. Serve `.woff2` as `font/woff2`; the folder's name carries the version, so it can be cached as
   immutable. A site that already serves the bundle for another viewer points this editor at that copy rather than
   adding its own; a different release goes in as a sibling version folder.
2. **Point the editor at it**, by the URL the site serves it at (absolute, or relative to the page):

   ```js
   import { mountXlsxEditor, OMF_VERSION } from 'open-xlsx-editor';
   mountXlsxEditor(el, { fonts: { omf: `/lib/fonts/omf/${OMF_VERSION}/` } });
   ```

   (A bundle under `lib/<name>/` beside that folder can also say `new URL(`../fonts/omf/${OMF_VERSION}/`,
   import.meta.url).href`.)

The first time a workbook names one of the 151 fonts, the editor adds its twin's faces to the page's fonts, per
alphabet, as `FontFace` objects (no stylesheet, so `font-src 'self'` is all the policy needs), after the font itself in
the cell's font stack. The browser downloads a file only when a cell shows text in that font, the computer lacks the
font, and the file holds the text's characters: a computer with Tahoma installed never fetches Tamarack. When a file
arrives the editor places and measures the text again, Excel's placement worked out from the original's Windows
metrics (the twin's are the same) and the browser's from the twin's own. Without `fonts.omf`, or before a file
arrives, the stand-ins as before (Calibri for Aptos Narrow). `exportPdf` embeds the twins too (docs/PDF.md).

Measured on this repository's Windows machine, for the fonts installed on it, against their twins served from the
release 1.0.1 bundle:

- **Canvas widths in Chromium**: every face of the 66 installed families the release has a twin for, 174 faces as
  Chromium names them, each character of U+0020-007E and U+00A1-00FF at 100px and six cell-like strings at 11pt: 156
  faces the same within 0.01px for every character. Of the other 18, 12 are Sitka italics, at most 0.27px at 100px
  (on "î") and 0.007px on a whole string at 11pt, and 2 Bahnschrift SemiCondensed, at most 0.05px. Bahnschrift
  Condensed and SemiCondensed are installed only as instances of one variable font, which Chromium draws at the same
  widths for both, so Condensed differs (by up to 20px at 100px; the twin is the condensed one).
  Franklin Gothic Medium and its italic differ on the canvas though their files' advances are the same: Chromium
  draws another installed face for that name.
- **The font files**: 141 installed faces (by name ID 1 and the style bits, as Windows links a family's faces), every
  character both map, 81,950 characters: the same advance widths, and the same Windows ascent and descent.
- **Against Excel's PDFs** (Edge, `tools/excel-compare`): basic.xlsx and large.xlsx (Aptos Narrow) and the editor's
  print of them agree within 1px across, down and in text width for 100,049 of 100,049 cells; with Calibri in its
  place, 99,028 are within 1px in width. Text widths decide where wrapped text breaks, where a number shows `####`
  and what AutoFit picks, so these match too. `exportPdf` with the twin embedded: large.xlsx's 100,000 cells all
  within 1px in width too, and the ten fixtures 99.9% (95.0% when it drew Aptos Narrow in Calibri; docs/PDF.md).

Not yet:
- Row heights: a row sized to its font takes the font's metrics from `src/engine/metrics.js` (read from a Windows
  machine's fonts), else Calibri's (for Aptos Narrow that gives Excel's 15pt). The families not in that table get
  Calibri's row height even when drawn in their twin; the twin's metrics, the original's, aren't used for it yet,
  as that hasn't been measured against Excel.
- A face the family doesn't have (Tahoma's italic) is slanted by the browser, as for an installed font; `exportPdf`
  draws it upright.

The demo loads them from `demo/fonts/omf/` (a copy of the bundle, not in this repository) when the page is opened with
`?fonts`. `tools/omf-map.py` (ported from open-docx-editor's `build-omf.py`) writes the editor's map of them,
`src/editor/omf.js`, from an open-metric-fonts checkout; `OMF_VERSION` is the release it was made from.

## Documentation

- `docs/FEATURES.md`: what Excel does and what this editor does yet, by ribbon tab, and the functions still missing.
- `docs/ARCHITECTURE.md`: the module map. Start here to change the code.

Each of the others says how a part was matched to Excel, what was measured, and what isn't done yet:

- **Opening, saving, printing:** `docs/FIDELITY.md` (what survives a save), `docs/ENCRYPTION.md` (password
  workbooks), `docs/CSV.md` (a .csv opened and saved, Text Import), `docs/LOCALE.md` (other countries' settings), `docs/LAYOUT.md` (the grid as Excel lays it out),
  `docs/PRINT.md` and `docs/PDF.md`.
- **Calculation:** `docs/CALC.md` (formulas and functions), `docs/DYNAMIC.md` (spilling, LET, LAMBDA),
  `docs/NAMES.md` (defined names), `docs/AUTOSUM.md`.
- **Editing:** `docs/ENTRY.md` (typed entries), `docs/CLIPBOARD.md` (copy, cut, Paste Special), `docs/FILL.md`,
  `docs/FIND.md`, `docs/GOTOSPECIAL.md` (and several ranges selected), `docs/PROOFING.md` (AutoCorrect, spelling).
- **Cells, rows and sheets:** `docs/FORMATS.md`, `docs/CONDFORMAT.md`, `docs/VALIDATION.md`, `docs/SIZES.md`,
  `docs/ROWCOL.md`, `docs/SHEETS.md`, `docs/VIEW.md` (frozen panes, zoom, gridlines, headings), `docs/PROTECT.md`
  (Protect Sheet and Protect Workbook).
- **Data and objects:** `docs/SORT.md`, `docs/FILTER.md`, `docs/DATATOOLS.md` (Flash Fill, Remove Duplicates, Text
  to Columns...), `docs/PIVOT.md`, `docs/CHARTS.md`, `docs/NOTES.md` (notes and comments), `docs/LINKS.md`
  (hyperlinks), `docs/QUERIES.md` (Power Query queries and connections, refreshed through the host).
- **Working with others:** `docs/COLLAB.md` (co-editing), `docs/AGENT.md` (the agent interface and MCP server).

## Contributing, security, licence

- `CONTRIBUTING.md`: setup, the tests, and how changes are measured against Excel. `npm run dev` serves the demo at
  http://localhost:8080.
- `SECURITY.md`: how to report a security problem privately.
- `CHANGELOG.md`: what changed in each release.
- MIT licence (`LICENSE`); the libraries it bundles are listed in `THIRD-PARTY-NOTICES.md`.
- Sister project of [open-docx-editor](https://github.com/janderson12345/open-docx-editor), the same for Word documents.
