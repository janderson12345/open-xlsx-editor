// The demo page: open an .xlsx from the file input or by dropping it, show it read-only. `window.oxe` is the editor's
// API, which the browser tests drive.
import { mountXlsxEditor } from '../src/index.js';

const host = document.getElementById('host');
const status = document.getElementById('status');
const input = document.querySelector('input[type=file]');
// (?fonts: open-metric-fonts served from demo/fonts/omf/, a copy of its web bundle (at least morrow-sans) that isn't in
// the repository: README, Fonts. ?ribbon=simple or classic: the ribbon's layout to start with.)
const params = new URLSearchParams(location.search);
const editor = mountXlsxEditor(host, { ...(params.has('fonts') ? { fonts: { omf: './fonts/omf/' } } : {}), ...(params.get('ribbon') ? { ribbon: params.get('ribbon') } : {}) });
window.oxe = editor;

async function openFile(file) {
  try {
    // (A workbook saved with a password: the editor asks for it, and opens it once given.)
    if (!(await editor.open(new Uint8Array(await file.arrayBuffer()), file.name))) { status.textContent = `${file.name} was not opened`; return; }
    const n = editor.workbook.sheets.length;
    status.textContent = `${file.name}: ${n} sheet${n === 1 ? '' : 's'}`;
  } catch (e) {
    status.textContent = `Could not open ${file.name}: ${e.message}`;
  }
}

input.addEventListener('change', () => input.files[0] && openFile(input.files[0]));
document.addEventListener('dragover', (e) => { e.preventDefault(); document.body.classList.add('drop'); });
document.addEventListener('dragleave', () => document.body.classList.remove('drop'));
document.addEventListener('drop', (e) => {
  e.preventDefault();
  document.body.classList.remove('drop');
  const file = e.dataTransfer.files[0];
  if (file) openFile(file);
});
