// open-xlsx-editor: open .xlsx bytes in a page and show them. Read-only for now: one sheet in a virtualised grid.
import { Workbook } from './ooxml/workbook.js';
import { createGrid } from './editor/grid.js';
import { renderPrint } from './editor/print.js';

export { Workbook };

/**
 * Mount the viewer in `host`.
 * @param {HTMLElement} host
 * @param {{ bytes?: Uint8Array | ArrayBuffer, fileName?: string, theme?: 'light' | 'dark' }} [options]
 */
export function mountXlsxEditor(host, options = {}) {
  const root = document.createElement('div');
  root.className = 'oxe-root';
  if (options.theme) root.dataset.theme = options.theme;
  host.appendChild(root);
  const state = { workbook: null, grid: null, sheetIndex: 0, fileName: null };

  const api = {
    /** Open a workbook and show the sheet it opens on (its active tab, or the first visible worksheet). */
    open(bytes, fileName = null) {
      const wb = Workbook.fromBytes(bytes);
      const i = wb.activeSheet;
      const sheet = wb.sheet(i);
      state.workbook = wb;
      state.sheetIndex = i;
      state.fileName = fileName;
      if (state.grid) state.grid.setSheet(sheet, wb.sheets[i].name);
      else state.grid = createGrid(root, { sheet, label: wb.sheets[i].name });
      return api;
    },
    /** The workbook as .xlsx bytes (unchanged parts byte for byte). */
    save() {
      if (!state.workbook) throw new Error('no workbook open');
      return state.workbook.toBytes();
    },
    /** Build the printed pages of the sheet shown (done before the browser prints, too); returns the page count. */
    preparePrint() {
      if (!state.workbook) return 0;
      return renderPrint(root, state.workbook.sheet(state.sheetIndex));
    },
    get workbook() { return state.workbook; },
    get grid() { return state.grid; },
    get sheetIndex() { return state.sheetIndex; },
    get fileName() { return state.fileName; },
    setTheme(theme) {
      if (theme) root.dataset.theme = theme;
      else delete root.dataset.theme;
    },
    destroy() {
      window.removeEventListener('beforeprint', beforePrint);
      state.grid?.destroy();
      root.remove();
    },
  };
  const beforePrint = () => api.preparePrint();
  window.addEventListener('beforeprint', beforePrint);
  if (options.bytes) api.open(options.bytes, options.fileName || null);
  return api;
}
