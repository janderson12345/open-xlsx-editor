// Text files in and the workbook out: the Text Import box's file read into a new workbook or at the active cell (Data
// > From Text/CSV), Save As (F12) and its box giving the bytes of the type chosen to the host's onSave, or to the
// browser as a download, and the page's locale changed with the grid drawn again. The editor's api (index.js) and its
// ribbon (acts.js) both come here.
import { textImportBox } from './textimportbox.js';
import { saveAsBox } from './saveasbox.js';
import { messageBox } from './dialogs.js';
import { guessSettings, splitText, importCells } from '../ooxml/textimport.js';
import { blankWorkbook, sheetNameOf } from '../ooxml/csv.js';
import { hasMacros } from '../ooxml/filetype.js';
import { currentLocale, localeOf, setLocale } from '../ooxml/locale.js';

const MIME = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  csv: 'text/csv', txt: 'text/plain', pdf: 'application/pdf',
};

/** A file handed to the browser to save (a download), as a site without a host's onSave saves one. */
function download({ fileName, bytes, mime }) {
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const a = Object.assign(document.createElement('a'), { href: url, download: fileName });
  a.hidden = true;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** `env`: { root, state, options, editing, api: () => the editor's api }. */
export function createFiles({ root, state, options, editing, api }) {
  const back = () => state.grid?.el.focus({ preventScroll: true });
  const files = {
    /**
     * A text file's bytes imported through the Text Import box (`ask` false: with `settings` as given, no box): into
     * a new workbook named after the file (`into` 'new', and without a workbook open), or at the active cell of the
     * sheet shown, one step of the undo history ('sheet'). Resolves to the editor or true when done, null if cancelled.
     */
    async importText(bytes, fileName = '', { into = 'new', settings = null, ask = true } = {}) {
      const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      const first = { ...guessSettings(data, fileName, currentLocale()), ...(settings || {}) };
      const got = ask ? await textImportBox(root, { bytes: data, fileName, settings: first }) : first;
      if (!got) { back(); return null; }
      const { rows } = splitText(data, got);
      const opts = { columns: got.columns || [], startRow: got.startRow || 1, locale: localeOf(got.locale || currentLocale()), decimal: got.decimal, group: got.group };
      if (into === 'new' || !state.workbook) {
        const wb = blankWorkbook(sheetNameOf(fileName));
        const { contents, styles } = importCells(wb, 0, 0, 0, rows, opts);
        if (contents.length) wb.apply(0, contents);
        if (styles.length) wb.setStyles(0, styles);
        return api().open(wb.toBytes(), fileName || null);
      }
      if (state.readOnly) return false;
      const s = editing().selection;
      const si = state.sheetIndex;
      const { contents, styles } = importCells(state.workbook, si, s.ar, s.ac, rows, opts);
      const done = editing().run('applyChanges', [{ si, contents, styles }]);
      back();
      return done;
    },
    /** Data > From Text/CSV: a file picked, then its Text Import box, the data put at the active cell. */
    pickText() {
      if (state.readOnly) return false;
      const input = Object.assign(document.createElement('input'), { type: 'file', accept: '.csv,.txt,.tsv,.prn,text/csv,text/plain' });
      input.hidden = true;
      root.append(input);
      return new Promise((resolve) => {
        input.addEventListener('change', async () => {
          const file = input.files?.[0];
          input.remove();
          if (!file) { resolve(null); return; }
          resolve(await files.importText(new Uint8Array(await file.arrayBuffer()), file.name, { into: 'sheet' }));
        });
        input.addEventListener('cancel', () => { input.remove(); resolve(null); });
        input.click();
      });
    },
    /**
     * Save As: the box (or `choice` as it would resolve: { fileName, type, sep, encoding, bom }), then the file's bytes
     * of that type given to `onSave({ fileName, type, bytes, mime })`, or downloaded. The workbook goes by the new
     * name from now on, as in Excel. Resolves to what was saved, or null.
     */
    async saveAs(choice = null) {
      const wb = state.workbook;
      if (!wb) throw new Error('no workbook open');
      const shown = wb.sheets.filter((x, i) => x.state === 'visible' && wb.canShow(i)).length;
      const got = choice || await saveAsBox(root, { fileName: state.fileName || '', macros: hasMacros(wb.pkg), sheets: shown, sheetName: wb.sheets[state.sheetIndex].name,
        pdf: typeof options.exportPdf === 'function', listSep: currentLocale().listSep });
      if (!got) { back(); return null; }
      let bytes;
      try {
        if (got.type === 'pdf') bytes = await options.exportPdf(api());
        else if (got.type === 'csv' || got.type === 'txt') bytes = api().saveCsv({ sep: got.sep, encoding: got.encoding, bom: got.bom, kind: got.type === 'txt' ? 'txt' : 'csv-utf8' });
        else bytes = api().save({ as: got.type });
      } catch (e) {
        await messageBox(root, e.message);
        back();
        return null;
      }
      const out = { fileName: got.fileName, type: got.type, bytes, mime: MIME[got.type] || 'application/octet-stream' };
      if (got.type !== 'pdf') { state.fileName = got.fileName; wb.fileName = got.fileName; }
      if (options.onSave) options.onSave(out);
      else download(out);
      back();
      return out;
    },
    /** The page's locale (ooxml/locale.js) set, and the sheet and formula bar drawn in it. */
    setLocale(spec) {
      const loc = setLocale(spec);
      if (state.grid) { state.grid.refresh(); state.editing?.refreshBar?.(); }
      return loc;
    },
  };
  return files;
}
