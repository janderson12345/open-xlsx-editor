// The Save As box: the file's name and its type, as Excel's Save As offers them: Excel Workbook (.xlsx), Excel
// Macro-Enabled Workbook (.xlsm, which keeps a workbook's macros; an .xlsx can't), CSV (a separator and an encoding
// chosen: CSV UTF-8, CSV, or others), Text (Tab delimited) (.txt, Windows-1252, or Unicode Text in UTF-16), and PDF
// when the host gives the editor a PDF export. A CSV or text file holds only the sheet shown: the box says so, as
// Excel warns when the workbook has more sheets, and that features such as formulas and formats aren't kept. Resolves
// to { fileName, type, sep, encoding, bom } or null.
import { modal, button } from './dialogs.js';
import { make, select, labelled } from './formatbox-kit.js';

const TAB = String.fromCharCode(9);
const TYPES = [['xlsx', 'Excel Workbook (*.xlsx)'], ['xlsm', 'Excel Macro-Enabled Workbook (*.xlsm)'], ['csv', 'CSV (*.csv)'], ['txt', 'Text (Tab delimited) (*.txt)'], ['pdf', 'PDF (*.pdf)']];
const SEPARATORS = [[',', 'Comma (,)'], [';', 'Semicolon (;)'], [TAB, 'Tab'], ['|', 'Pipe (|)']];
const CSV_ENCODINGS = [['utf-8-bom', 'UTF-8 (CSV UTF-8, with a byte order mark)'], ['utf-8', 'UTF-8, no byte order mark'], ['windows-1252', 'Western European (Windows-1252)'], ['utf-16le', 'Unicode (UTF-16)']];
const TXT_ENCODINGS = [['windows-1252', 'Western European (Windows-1252)'], ['utf-16le', 'Unicode Text (UTF-16)'], ['utf-8-bom', 'UTF-8, with a byte order mark'], ['utf-8', 'UTF-8, no byte order mark']];

/** A file name's base (no folder, no extension). */
export const baseName = (name) => String(name || '').replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '');

/** The type a file name's extension gives, if the box offers it. */
const typeOf = (name) => {
  const ext = /\.([^.]+)$/.exec(name || '')?.[1]?.toLowerCase();
  return { xlsx: 'xlsx', xlsm: 'xlsm', csv: 'csv', txt: 'txt', tsv: 'txt', pdf: 'pdf' }[ext] || null;
};

/**
 * The box. `fileName`: the open file's name (its base kept, its type chosen first); `macros`: the workbook has a VBA
 * project (.xlsm chosen first, and .xlsx said to lose it); `sheets` and `sheetName`: how many sheets and the one shown;
 * `pdf`: PDF offered; `listSep`: the locale's list separator, a CSV's first separator.
 */
export function saveAsBox(root, { fileName = '', macros = false, sheets = 1, sheetName = '', pdf = false, listSep = ',' } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-saveas-box');
    const name = make('input', { type: 'text', className: 'oxe-dialog-field oxe-saveas-name', spellcheck: false, value: baseName(fileName) || 'Book1' });
    const offered = TYPES.filter(([v]) => v !== 'pdf' || pdf);
    const first = macros ? 'xlsm' : typeOf(fileName) && offered.some(([v]) => v === typeOf(fileName)) ? typeOf(fileName) : 'xlsx';
    const type = select(offered, first);
    const sep = select(SEPARATORS, SEPARATORS.some(([v]) => v === listSep) ? listSep : ',');
    const csvEnc = select(CSV_ENCODINGS, 'utf-8-bom');
    const txtEnc = select(TXT_ENCODINGS, 'windows-1252');
    const csvEl = make('div', { className: 'oxe-saveas-options' }, labelled('Field separator:', sep), labelled('Encoding:', csvEnc));
    const txtEl = make('div', { className: 'oxe-saveas-options' }, labelled('Encoding:', txtEnc));
    const note = make('p', { className: 'oxe-dialog-text oxe-saveas-note', role: 'status' });
    const sync = () => {
      const t = type.value;
      csvEl.hidden = t !== 'csv';
      txtEl.hidden = t !== 'txt';
      const lines = [];
      if (t === 'csv' || t === 'txt') {
        // (Excel's warnings for a text format: only the active sheet, and the features a text file can't hold.)
        if (sheets > 1) lines.push(`Only the active sheet, ${sheetName}, is saved: a ${t === 'csv' ? 'CSV' : 'text'} file can't hold a workbook of several sheets.`);
        lines.push('Formulas are saved as their values and formatting is lost: open the file again and only the values come back.');
      }
      if (t === 'xlsx' && macros) lines.push("This workbook has macros, which an .xlsx can't keep. Save it as an Excel Macro-Enabled Workbook (.xlsm) to keep them.");
      if (t === 'pdf') lines.push('The sheet shown is saved as it prints.');
      note.textContent = lines.join(' ');
      note.hidden = !lines.length;
    };
    type.addEventListener('change', sync);
    sync();
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
    const ok = () => {
      const base = baseName(name.value.trim()) || name.value.trim();
      // (Excel refuses these characters in a file name.)
      if (!base || /[\\/:*?"<>|]/.test(name.value.replace(/^.*[\\/]/, ''))) { why.textContent = 'The file name is not valid: it can\'t be empty or hold \\ / : * ? " < > |'; return; }
      const t = type.value;
      const enc = t === 'csv' ? csvEnc.value : t === 'txt' ? txtEnc.value : null;
      done({
        fileName: `${base}.${t}`, type: t, sep: t === 'csv' ? sep.value : t === 'txt' ? TAB : null,
        encoding: enc ? enc.replace(/-bom$/, '') : null, bom: enc ? enc.endsWith('-bom') || enc === 'utf-16le' : null,
      });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Save As' }),
      labelled('File name:', name), labelled('Save as type:', type), csvEl, txtEl, note, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('Save', ok), button('Cancel', () => done(null))));
    name.dataset.first = '';
    // (The name selected, ready to type over, as Excel's box has it.)
    queueMicrotask(() => name.select());
  });
}
