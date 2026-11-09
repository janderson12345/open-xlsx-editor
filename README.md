# open-xlsx-editor

Open and edit .xlsx files in the browser. Plain JavaScript, no server, no Office.

Early days: it opens a workbook and shows its first sheet read-only in a virtualised grid, and saves it back with
every part it didn't change byte for byte. Styles, number formats, editing in the page and printing come next.
Sister project of open-docx-editor, built the same way.

```js
import { mountXlsxEditor } from 'open-xlsx-editor';
import 'open-xlsx-editor/style.css';

const editor = mountXlsxEditor(document.getElementById('host'), { bytes, fileName: 'book.xlsx' });
const saved = editor.save(); // Uint8Array
```

- `npm run dev`: the demo at http://localhost:8080 (drag an .xlsx onto it).
- `CONTRIBUTING.md`: setup and tests. `docs/ARCHITECTURE.md`: the module map. `docs/LAYOUT.md`: matching Excel.
- `LESSONS.md`: what the Word editor taught. `CLAUDE.md`: the rules for working on this code. `docs/REUSE.md`: which
  Word editor modules carry over.

MIT licence.
