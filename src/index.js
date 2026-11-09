// open-xlsx-editor: open .xlsx bytes in a page, show and edit them, save them back. A sheet at a time in a virtualised
// grid, the sheet tabs under it.
import { Workbook } from './ooxml/workbook.js';
import { createGrid } from './editor/grid.js';
import { renderPrint } from './editor/print.js';
import { applyCellStyles } from './editor/cellstyles.js';
import { attachEditing } from './editor/interact.js';
import { createSheetTabs } from './editor/tabs.js';
import { fontFamily } from './ooxml/styles.js';
import { parseRef } from './ooxml/refs.js';

export { Workbook };

/**
 * Mount the editor in `host`.
 * @param {HTMLElement} host
 * @param {{ bytes?: Uint8Array | ArrayBuffer, fileName?: string, theme?: 'light' | 'dark', readOnly?: boolean,
 *   onChange?: () => void }} [options]
 */
export function mountXlsxEditor(host, options = {}) {
  const root = document.createElement('div');
  root.className = 'oxe-root';
  if (options.theme) root.dataset.theme = options.theme;
  host.appendChild(root);
  const state = { workbook: null, grid: null, tabs: null, sheetIndex: 0, fileName: null, looks: {}, readOnly: !!options.readOnly, editing: null };
  const at = (ref) => {
    const p = typeof ref === 'string' ? parseRef(ref) : ref;
    if (!p) throw new Error(`not a cell reference: ${ref}`);
    return p;
  };
  const editing = () => {
    if (!state.editing) throw new Error('no workbook open');
    return state.editing.commands;
  };
  /** Show sheet i in the grid and mark its tab (commands.js showSheet keeps each sheet's selection around this). */
  const showSheet = (i) => {
    state.sheetIndex = i;
    state.grid.setSheet(state.workbook.sheet(i), state.workbook.sheets[i].name, state.looks);
    state.tabs.activate(i);
  };
  /** The tabs a workbook shows: every sheet that isn't hidden, chart sheets as tabs that can't be shown yet. */
  const tabList = (wb) => wb.sheets.map((s, index) => ({ index, name: s.name, state: s.state, color: wb.tabColor(index), disabled: !wb.canShow(index) }))
    .filter((t) => t.state === 'visible');

  const api = {
    /** Open a workbook and show the sheet it opens on (its active tab, or the first visible worksheet). */
    open(bytes, fileName = null) {
      const wb = Workbook.fromBytes(bytes);
      const i = wb.activeSheet;
      const sheet = wb.sheet(i);
      state.workbook = wb;
      state.sheetIndex = i;
      state.fileName = fileName;
      state.looks = applyCellStyles(root, wb);
      if (state.grid) state.grid.setSheet(sheet, wb.sheets[i].name, state.looks);
      else {
        state.grid = createGrid(root, { sheet, label: wb.sheets[i].name, looks: state.looks, readOnly: state.readOnly });
        state.editing = attachEditing(root, {
          workbook: () => state.workbook,
          sheetIndex: () => state.sheetIndex,
          showSheet,
          grid: state.grid,
          looks: () => state.looks,
          readOnly: () => state.readOnly,
          changed: () => options.onChange?.(),
        });
        state.editing.setReadOnly(state.readOnly);
        state.tabs = createSheetTabs(root, { controls: state.grid.el.id, onSelect: (k) => state.editing.commands.run('showSheet', k) });
      }
      state.tabs.set(tabList(wb), i);
      state.editing.reset();
      return api;
    },
    /** The workbook as .xlsx bytes (unchanged parts byte for byte, changed sheets rewritten in place). */
    save() {
      if (!state.workbook) throw new Error('no workbook open');
      return state.workbook.toBytes();
    },
    /** Select a cell ("B2" or { r, c }, 0-based). */
    select(ref) { const p = at(ref); editing().run('select', p.r, p.c); },
    /** Enter text into a cell as if typed there ("12", "TRUE", "50%", "'007", "=SUM(A1:A3)"); false if refused. */
    enter(ref, text) { return editing().run('enter', String(text), at(ref)); },
    /** Set a cell to a number, boolean, string or null; false if refused (read-only, or part of an array formula). */
    setCell(ref, value) { const p = at(ref); return editing().run('setValue', p.r, p.c, value); },
    undo: () => editing().run('undo'),
    redo: () => editing().run('redo'),
    /** Show a sheet (its index, or its name); false if it can't be shown (hidden, or a chart sheet). */
    showSheet(sheet) {
      const n = String(sheet).toLowerCase();
      const i = typeof sheet === 'number' ? sheet : state.workbook ? state.workbook.sheets.findIndex((s) => s.name.toLowerCase() === n) : -1;
      return editing().run('showSheet', i);
    },
    /** Run a named command (commands.js) with its arguments. */
    command: (name, ...args) => editing().run(name, ...args),
    get selection() { return state.editing ? state.editing.commands.selection : null; },
    /** Build the printed pages of the sheet shown (done before the browser prints, too); returns the page count. */
    preparePrint() {
      if (!state.workbook) return 0;
      const wb = state.workbook;
      return renderPrint(root, wb.sheet(state.sheetIndex), state.looks, fontFamily(wb.styles.normal, wb.theme));
    },
    get workbook() { return state.workbook; },
    get grid() { return state.grid; },
    get sheetIndex() { return state.sheetIndex; },
    get fileName() { return state.fileName; },
    get readOnly() { return state.readOnly; },
    setReadOnly(flag) {
      state.readOnly = !!flag;
      state.editing?.setReadOnly(state.readOnly);
    },
    setTheme(theme) {
      if (theme) root.dataset.theme = theme;
      else delete root.dataset.theme;
    },
    destroy() {
      window.removeEventListener('beforeprint', beforePrint);
      state.editing?.destroy();
      state.grid?.destroy();
      root.remove();
    },
  };
  const beforePrint = () => api.preparePrint();
  window.addEventListener('beforeprint', beforePrint);
  if (options.bytes) api.open(options.bytes, options.fileName || null);
  return api;
}
