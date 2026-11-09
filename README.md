# open-xlsx-editor

Open and edit .xlsx files in the browser. Plain JavaScript, no server, no Office.

It opens a workbook, shows its sheets in a virtualised grid with their fonts, fills, borders, alignment and number
formats, the sheet tabs underneath (in their tab colours; Ctrl+Page Up and Page Down change sheet; insert, rename,
move, colour, hide and delete sheets from the tab menu, with formulas and names kept pointing at the right sheets),
prints a sheet as Excel lays it out (measured against Excel's own PDFs), and lets you edit cell values: select with the
keyboard or mouse, type over a cell or press F2, Delete, undo and redo. Formulas are worked out as Excel works them
out (about 130 functions; 399 of 399 test formulas give Excel's results), again whenever what they read changes, and
can be typed, pointing at cells with the arrow keys or the mouse, on this sheet or another, as in Excel. Saving
rewrites only the sheets you changed; every other part goes back byte for byte. The ribbon comes next.
Sister project of open-docx-editor, built the same way.

```js
import { mountXlsxEditor } from 'open-xlsx-editor';
import 'open-xlsx-editor/style.css';

const editor = mountXlsxEditor(document.getElementById('host'), { bytes, fileName: 'book.xlsx' });
editor.enter('B2', '42');      // as if typed; editor.setCell('C3', 'text'), editor.undo(), editor.select('A1')
editor.enter('B3', '=B2*2');   // a formula; the cells that read B2 or B3 are worked out again
editor.showSheet('Notes');     // another sheet, by name or index, with the selection it had
editor.renameSheet('Notes', 'Q1 notes'); // formulas that read it follow; also insertSheet, moveSheet, hideSheet...
const saved = editor.save(); // Uint8Array
```

- `npm run dev`: the demo at http://localhost:8080 (drag an .xlsx onto it).
- `CONTRIBUTING.md`: setup and tests. `docs/ARCHITECTURE.md`: the module map. `docs/LAYOUT.md`: matching Excel's
  layout. `docs/CALC.md`: matching Excel's calculation. `docs/SHEETS.md`: matching Excel's sheet operations.
  `docs/FEATURES.md`: what Excel does and what this editor does yet, by ribbon tab.
- `LESSONS.md`: what the Word editor taught. `CLAUDE.md`: the rules for working on this code. `docs/REUSE.md`: which
  Word editor modules carry over.

MIT licence.
