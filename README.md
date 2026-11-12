# open-xlsx-editor

Open and edit .xlsx files in the browser. Plain JavaScript, no server, no Office.

It opens a workbook with Excel's Home tab over it (font, borders, fill and font colour, alignment, number formats,
written to the file as Excel writes them), shows its sheets in a virtualised grid with their fonts, fills, borders,
alignment and number formats, the sheet tabs underneath (in their tab colours; Ctrl+Page Up and Page Down change sheet; insert, rename,
move, colour, hide and delete sheets from the tab menu, with formulas and names kept pointing at the right sheets),
prints a sheet as Excel lays it out (measured against Excel's own PDFs), and lets you edit cell values: select with the
keyboard or mouse, type over a cell or press F2, Delete, undo and redo. Formulas are worked out as Excel works them
out (486 functions, dynamic arrays that spill as Excel 365's do, and table references; 1553 of 1553 test formulas give Excel's results), again whenever what they read changes, and
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
const saved = editor.save(); // Uint8Array
```

## Read-only mode and turning features off

As in open-docx-editor:

```js
const editor = mountXlsxEditor(el, {
  readOnly: true,                       // show the workbook only (no ribbon, nothing editable)
  features: { number: false },          // hide a tab by name, or a group by its label on any tab
  toolbar: true,                        // false hides the ribbon but keeps editing
  statusBar: true,                      // false hides Average, Count and Sum of the selection
});
editor.setReadOnly(false);              // switch at any time
editor.setFeatures({ number: true });
```

Tab names: `home`, `view` (the other tabs come later). Group names are their labels in lower case with dashes: `undo`,
`clipboard`, `font`, `alignment`, `number`, `cells`, `editing`; on View, `zoom` and `window`. A group turned off also turns off its keys (`font: false`:
Ctrl+B does nothing; `clipboard: false`: Ctrl+C, X and V). F6 moves between the sheet and the ribbon, Esc goes back to the sheet; the ⌃ at the right switches to a
one-line ribbon. `docs/FEATURES.md` lists what is there and what isn't yet.

- `npm run dev`: the demo at http://localhost:8080 (drag an .xlsx onto it).
- `CONTRIBUTING.md`: setup and tests. `docs/ARCHITECTURE.md`: the module map. `docs/LAYOUT.md`: matching Excel's
  layout. `docs/CALC.md`: matching Excel's calculation. `docs/SHEETS.md`: matching Excel's sheet operations.
  `docs/FORMATS.md`: matching Excel's Home tab. `docs/CLIPBOARD.md`: matching Excel's copy, cut, paste and Paste Special.
  `docs/ROWCOL.md`: matching Excel's inserting and deleting of rows, columns and cells. `docs/SIZES.md`: widths, heights,
  hiding. `docs/FIND.md`: Find and Replace. `docs/VIEW.md`: frozen panes and zoom.
  `docs/ENTRY.md`: how typed entries are read. `docs/FILL.md`: Fill and the fill handle. `docs/SORT.md`: Sort. `docs/AUTOSUM.md`: AutoSum. `docs/FILTER.md`: Filter. `docs/NOTES.md`: notes and comments. `docs/CONDFORMAT.md`: conditional formatting. `docs/VALIDATION.md`: data validation. `docs/FIDELITY.md`: what survives a save, checked by Excel. `docs/CHARTS.md`: charts.
  `docs/FEATURES.md`: what Excel does and what this editor does yet, by ribbon tab.
- `LESSONS.md`: what the Word editor taught. `CLAUDE.md`: the rules for working on this code. `docs/REUSE.md`: which
  Word editor modules carry over.

MIT licence.
