# open-xlsx-editor

Open and edit .xlsx files in the browser. Plain JavaScript, no server, no Office.

It opens a workbook, shows the sheet it opens on in a virtualised grid with its fonts, fills, borders, alignment and
number formats, prints it as Excel lays it out (measured against Excel's own PDFs), and lets you edit cell values:
select with the keyboard or mouse, type over a cell or press F2, Delete, undo and redo. Saving rewrites only the sheet
you changed; every other part goes back byte for byte. Formulas are shown with their saved values and can't be edited
yet; more sheets, formulas and the ribbon come next.
Sister project of open-docx-editor, built the same way.

```js
import { mountXlsxEditor } from 'open-xlsx-editor';
import 'open-xlsx-editor/style.css';

const editor = mountXlsxEditor(document.getElementById('host'), { bytes, fileName: 'book.xlsx' });
editor.enter('B2', '42');      // as if typed; editor.setCell('C3', 'text'), editor.undo(), editor.select('A1')
const saved = editor.save(); // Uint8Array
```

- `npm run dev`: the demo at http://localhost:8080 (drag an .xlsx onto it).
- `CONTRIBUTING.md`: setup and tests. `docs/ARCHITECTURE.md`: the module map. `docs/LAYOUT.md`: matching Excel.
- `LESSONS.md`: what the Word editor taught. `CLAUDE.md`: the rules for working on this code. `docs/REUSE.md`: which
  Word editor modules carry over.

MIT licence.
