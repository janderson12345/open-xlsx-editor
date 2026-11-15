# open-xlsx-editor

Open and edit .xlsx files in the browser. Plain JavaScript, no server, no Office.

It opens a workbook with Excel's Home tab over it (font, borders, fill and font colour, alignment, number formats,
written to the file as Excel writes them), shows its sheets in a virtualised grid with their fonts, fills, borders,
alignment and number formats, the sheet tabs underneath (in their tab colours; Ctrl+Page Up and Page Down change sheet; insert, rename,
move, colour, hide and delete sheets from the tab menu, with formulas and names kept pointing at the right sheets),
prints a sheet as Excel lays it out (measured against Excel's own PDFs), and lets you edit cell values: select with the
keyboard or mouse, type over a cell or press F2, Delete, undo and redo. Formulas are worked out as Excel works them
out (500 of Excel's 519 functions, dynamic arrays that spill as Excel 365's do, and table references; 1914 of 1914 test formulas give Excel's results), again whenever what they read changes, and
can be typed, pointing at cells with the arrow keys or the mouse, on this sheet or another, as in Excel. Saving
rewrites only the parts you changed; every other part goes back byte for byte.
An AI assistant can read and edit workbooks too, through an MCP server (`open-xlsx-editor-mcp`) or the agent
interface, and every change it makes can be reviewed cell by cell and taken back before saving (docs/AGENT.md).
Sister project of open-docx-editor, built the same way.

```js
import { mountXlsxEditor } from 'open-xlsx-editor';
import 'open-xlsx-editor/style.css';

const editor = mountXlsxEditor(document.getElementById('host'), { bytes, fileName: 'book.xlsx' });
editor.enter('B2', '42');      // as if typed; editor.setCell('C3', 'text'), editor.undo(), editor.select('A1')
editor.enter('B3', '=B2*2');   // a formula; the cells that read B2 or B3 are worked out again
editor.showSheet('Notes');     // another sheet, by name or index, with the selection it had
editor.renameSheet('Notes', 'Q1 notes'); // formulas that read it follow; also insertSheet, moveSheet, hideSheet...
editor.command('toggleFont', 'b'); // what the ribbon's buttons do, on the selection; Ctrl+B does the same
const saved = editor.save(); // Uint8Array; save({ as: 'xltx' }) as a template, or a template as 'xlsx'
```

## Read-only mode and turning features off

As in open-docx-editor:

```js
const editor = mountXlsxEditor(el, {
  readOnly: true,                       // show the workbook only (no ribbon, nothing editable)
  features: { number: false },          // hide a tab by name, or a group by its label on any tab
  toolbar: true,                        // false hides the ribbon but keeps editing
  ribbon: 'simple',                     // start with the one-line ribbon ('classic': the full one); setRibbon() later
  statusBar: true,                      // false hides Average, Count and Sum; { minimum: true, maximum: true } adds those
});
editor.setReadOnly(false);              // switch at any time
editor.setFeatures({ number: true });
```

Tab names: `home`, `insert`, `page-layout`, `formulas`, `data`, `review`, `view`. Group names are their labels in lower
case with dashes: `undo`, `clipboard`, `font`, `alignment`, `number`, `cells`, `editing`; on Page Layout,
`sheet-options`; on View, `show`, `zoom` and `window`. A group turned off also turns off its keys (`font: false`:
Ctrl+B does nothing; `clipboard: false`: Ctrl+C, X and V). F6 moves between the sheet and the ribbon, Esc goes back to the sheet; the ⌃ at the right switches to a
one-line ribbon. `docs/FEATURES.md` lists what is there and what isn't yet.

- `npm run dev`: the demo at http://localhost:8080 (drag an .xlsx onto it).
- `CONTRIBUTING.md`: setup and tests. `docs/FEATURES.md`: what Excel does and what this editor does yet, by ribbon tab,
  and the functions still missing. `docs/ARCHITECTURE.md`: the module map. Each of the others says how a part was
  matched to Excel, what was measured, and what isn't done yet:
  - Opening, saving, printing: `docs/FIDELITY.md` (what survives a save), `docs/ENCRYPTION.md` (password workbooks),
    `docs/CSV.md` (a .csv opened and saved),
    `docs/LAYOUT.md` (the grid as Excel lays it out), `docs/PRINT.md` and `docs/PDF.md`.
  - Calculation: `docs/CALC.md` (formulas and functions), `docs/DYNAMIC.md` (spilling, LET, LAMBDA), `docs/NAMES.md`
    (defined names), `docs/AUTOSUM.md`.
  - Editing: `docs/ENTRY.md` (typed entries), `docs/CLIPBOARD.md` (copy, cut, Paste Special), `docs/FILL.md`,
    `docs/FIND.md`, `docs/GOTOSPECIAL.md` (and several ranges selected), `docs/PROOFING.md` (AutoCorrect, spelling).
  - Cells, rows and sheets: `docs/FORMATS.md`, `docs/CONDFORMAT.md`, `docs/VALIDATION.md`, `docs/SIZES.md`,
    `docs/ROWCOL.md`, `docs/SHEETS.md`, `docs/VIEW.md` (frozen panes, zoom, gridlines, headings), `docs/PROTECT.md`
    (Protect Sheet and Protect Workbook).
  - Data and objects: `docs/SORT.md`, `docs/FILTER.md`, `docs/DATATOOLS.md` (Flash Fill, Remove Duplicates, Text to
    Columns...), `docs/PIVOT.md`, `docs/CHARTS.md`, `docs/NOTES.md` (notes and comments), `docs/LINKS.md` (hyperlinks).
  - Working with others: `docs/COLLAB.md` (co-editing), `docs/AGENT.md` (the agent interface).
- `LESSONS.md`: what the Word editor taught. `CLAUDE.md`: the rules for working on this code. `docs/REUSE.md`: which
  Word editor modules carry over.

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
help it is drawn in Calibri (or Carlito), which is close but not the same width. [open-metric-fonts](https://github.com/open-metric-fonts/open-metric-fonts)
makes Morrow Sans Narrow and Morrow Sans: open (SIL OFL) fonts with exactly Aptos Narrow's and Aptos's character
widths and line metrics. The editor loads them from a folder your site serves, the same folder the Word editor
(open-docx-editor) uses, so one copy serves every viewer of Office documents on the site:

1. **Serve the open-metric-fonts web bundle once, in a folder named for its version.** The editor is made for release
   `OMF_VERSION` (exported by the package; 1.0.1 now). The release's web bundle holds every package as
   `<package>/files/<face>-<alphabet>.woff2`, as the editor looks for them; it needs `morrow-sans/` only.

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

   Or only `morrow-sans`, from npm with your other dependencies, pinned (`npm install --save-exact
   @open-metric-fonts/morrow-sans@1.0.1`), copied to `<folder>/morrow-sans/` by your build. Serve `.woff2` as
   `font/woff2`; the folder's name carries the version, so it can be cached as immutable. A site that already serves
   the bundle for another viewer points this editor at that copy rather than adding its own; a different release goes
   in as a sibling version folder.
2. **Point the editor at it**, by the URL the site serves it at (absolute, or relative to the page):

   ```js
   import { mountXlsxEditor, OMF_VERSION } from 'open-xlsx-editor';
   mountXlsxEditor(el, { fonts: { omf: `/lib/fonts/omf/${OMF_VERSION}/` } });
   ```

   (A bundle under `lib/<name>/` beside that folder can also say `new URL(`../fonts/omf/${OMF_VERSION}/`,
   import.meta.url).href`.)

The editor adds each face of Morrow Sans Narrow and Morrow Sans to the page's fonts, per alphabet, as `FontFace`
objects (no stylesheet, so `font-src 'self'` is all the policy needs). The browser downloads a file only when a cell
shows text in it, and when one arrives the editor places and measures the text again with that font's own metrics.
Without `fonts.omf`, or before a file arrives, Calibri stands in. Measured with release 1.0.1 served this way (Edge,
the comparison in `tools/excel-compare`): Excel's PDFs of basic.xlsx and large.xlsx (Aptos Narrow) and the editor's
print of them agree within 1px across, down and in text width for 100,049 of 100,049 cells; with Calibri in its
place, 99,028 are within 1px in width. Text widths decide where wrapped text breaks, where a number shows `####` and
what AutoFit picks, so these match too.

Not yet:
- Only the Aptos family is wired to its twin. open-metric-fonts' other families (Tahoma's, Verdana's, Century
  Gothic's, Garamond's and more) aren't yet in the editor's font stacks.
- `exportPdf` (`open-xlsx-editor/pdf`) embeds installed fonts only, so a served font is drawn there in a PDF standard
  font at the same widths. The browser's own Print uses the served font.

The demo loads them from `demo/fonts/omf/` (a copy of the bundle, or of `morrow-sans/`, not in this repository) when
the page is opened with `?fonts`. `tools/omf-faces.mjs` writes the editor's list of faces from a newer `morrow-sans`.

MIT licence.
