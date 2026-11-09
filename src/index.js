// open-xlsx-editor: open .xlsx bytes in a page, show and edit them, save them back. A sheet at a time in a virtualised
// grid, the sheet tabs under it (with their menu: insert, delete, rename, move, colour, hide, unhide).
import { Workbook } from './ooxml/workbook.js';
import { createGrid } from './editor/grid.js';
import { renderPrint } from './editor/print.js';
import { applyCellStyles } from './editor/cellstyles.js';
import { attachEditing } from './editor/interact.js';
import { createSheetTabs } from './editor/tabs.js';
import { createMenu, tabMenuEntries } from './editor/sheetmenu.js';
import { messageBox, unhideBox } from './editor/dialogs.js';
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
  /** A sheet by index or name (case ignored), as an index (-1 for none). */
  const sheetAt = (sheet) => {
    if (typeof sheet === 'number') return sheet;
    const n = String(sheet).toLowerCase();
    return state.workbook ? state.workbook.sheets.findIndex((s) => s.name.toLowerCase() === n) : -1;
  };
  /** After sheets changed (commands.js): the tabs again, the sheet shown now at index `now` (-1: it went). */
  const sheetsChanged = (now) => {
    state.sheetIndex = now;
    state.tabs.set(tabList(state.workbook), now);
    if (now >= 0) state.grid.el.setAttribute('aria-label', state.workbook.sheets[now].name);
  };
  /** A refusal Excel shows in a message box; `state.alert` resolves when it is closed. */
  const alert = (msg) => { state.alert = messageBox(root, msg); };
  const menu = createMenu(root);
  const sheetActs = {
    run: (name, ...args) => editing().run(name, ...args),
    rename: (i) => {
      if (state.readOnly) return;
      state.tabs.editName(i, async (name) => {
        // (Renamed, focus goes back to the sheet, as in Excel.)
        if (editing().run('renameSheet', i, name)) { state.grid.el.focus({ preventScroll: true }); return true; }
        await state.alert;
        return false;
      });
    },
    remove: async (i) => {
      const wb = state.workbook;
      // (Excel asks first when the sheet has something in it.)
      const filled = wb.sheets[i].kind !== 'worksheet' || [...wb.sheet(i).cells.values()].some((line) => [...line.values()].some((c) => c.v != null || c.f != null));
      if (filled && await messageBox(root, 'This sheet will be permanently deleted. Do you want to continue?', ['Delete', 'Cancel']) !== 'Delete') return;
      editing().run('deleteSheet', i);
    },
    unhide: async () => {
      const wb = state.workbook;
      const hidden = wb.sheets.map((s, k) => ({ s, k })).filter((x) => x.s.state === 'hidden');
      const pick = await unhideBox(root, hidden.map((x) => x.s.name));
      if (pick != null) editing().run('hideSheet', hidden[pick].k, false);
    },
  };
  const openMenu = (i, point, byKeyboard) => {
    // (As in Excel, the tab right-clicked becomes the sheet shown.)
    if (state.workbook.canShow(i)) editing().run('showSheet', i);
    menu.open(tabMenuEntries(state.workbook, i, { act: sheetActs, color: (c) => state.looks.color(c), readOnly: state.readOnly }), point,
      { label: `${state.workbook.sheets[i].name} sheet`, byKeyboard });
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
          sheetsChanged,
          alert,
          grid: state.grid,
          looks: () => state.looks,
          readOnly: () => state.readOnly,
          changed: () => options.onChange?.(),
        });
        state.editing.setReadOnly(state.readOnly);
        state.tabs = createSheetTabs(root, {
          controls: state.grid.el.id,
          on: {
            select: (k) => editing().run('showSheet', k),
            menu: openMenu,
            rename: sheetActs.rename,
            move: (from, to) => editing().run('moveSheet', from, to),
            add: () => editing().run('insertSheet', { after: true }),
          },
        });
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
    showSheet: (sheet) => editing().run('showSheet', sheetAt(sheet)),
    /** A new sheet before the one shown (`after`: after it), shown; false if read-only. */
    insertSheet: ({ after = false } = {}) => editing().run('insertSheet', { after }),
    /** Rename a sheet (index or name); false if refused (Excel's message shown, as Excel shows it). */
    renameSheet: (sheet, name) => editing().run('renameSheet', sheetAt(sheet), name),
    /** Move a sheet to index `to`. */
    moveSheet: (sheet, to) => editing().run('moveSheet', sheetAt(sheet), to),
    /** Hide a sheet's tab, or show it again (`hidden` false); false if refused (the last sheet shown). */
    hideSheet: (sheet, hidden = true) => editing().run('hideSheet', sheetAt(sheet), hidden),
    /** A sheet's tab colour: { rgb: 'FFRRGGBB' }, { theme, tint }, or null for none. */
    setTabColor: (sheet, color) => editing().run('tabColor', sheetAt(sheet), color),
    /** Delete a sheet, without asking (the tab menu asks first); false if refused (the last sheet shown). */
    deleteSheet: (sheet) => editing().run('deleteSheet', sheetAt(sheet)),
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
      menu.close();
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
