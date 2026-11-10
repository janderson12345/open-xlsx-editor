// open-xlsx-editor: open .xlsx bytes in a page, show and edit them, save them back. The ribbon over a sheet at a time
// in a virtualised grid, the sheet tabs under it (with their menu: insert, delete, rename, move, colour, hide, unhide).
import { Workbook } from './ooxml/workbook.js';
import { createGrid } from './editor/grid.js';
import { renderPrint } from './editor/print.js';
import { applyCellStyles } from './editor/cellstyles.js';
import { attachEditing } from './editor/interact.js';
import { createSheetTabs } from './editor/tabs.js';
import { createMenu, tabMenuEntries } from './editor/sheetmenu.js';
import { messageBox, unhideBox, numberBox, optionsBox } from './editor/dialogs.js';
import { shapeShift, ROWCOL_MESSAGES } from './ooxml/rowcol.js';
import { createRibbon } from './editor/ribbon.js';
import { createFindBar } from './editor/findbar.js';
import { homeTab } from './editor/hometab.js';
import { viewTab } from './editor/viewtab.js';
import { fontFamily } from './ooxml/styles.js';
import { resolveColor } from './ooxml/theme.js';
import { mergeLoss } from './ooxml/mergeops.js';
import { MAX_ROWS, MAX_COLS } from './ooxml/refs.js';
import { parseRef } from './ooxml/refs.js';

export { Workbook };

/** The ribbon group a formatting command belongs to, for `features`: one of a group turned off doesn't run, from the
 *  ribbon or the keyboard. */
const FEATURE_OF = { toggleFont: 'font', stepFont: 'font', borders: 'font', toggleWrap: 'alignment', stepIndent: 'alignment', align: 'alignment', merge: 'alignment',
  resizeLines: 'cells', columnWidth: 'cells', rowHeight: 'cells', fitLines: 'cells', hideLines: 'cells', rowHeightBox: 'cells', columnWidthBox: 'cells',
  freeze: 'window', zoom: 'zoom', zoomBox: 'zoom', zoomSelection: 'zoom', zoomToSelection: 'zoom',
  lines: 'cells', cells: 'cells', insertCells: 'cells', deleteCells: 'cells', insertRows: 'cells', deleteRows: 'cells', insertColumns: 'cells', deleteColumns: 'cells',
  clipCopy: 'clipboard', clipCut: 'clipboard', clipPaste: 'clipboard', copy: 'clipboard', cut: 'clipboard', paste: 'clipboard',
  stepDecimal: 'number', clearFormats: 'editing' };
const groupOf = (name, change) => (name !== 'format' ? FEATURE_OF[name] : change.numFmt != null || change.style ? 'number' : change.align ? 'alignment' : 'font');

/**
 * Mount the editor in `host`.
 * - `readOnly`: show the workbook only (no ribbon, nothing editable); `setReadOnly()` changes it.
 * - `features`: turn parts of the ribbon off: a tab by name ({ home: false }) or a group by its label in lower case
 *   with dashes, on any tab ({ font: false, number: false }); `setFeatures()` changes them. A group turned off also
 *   turns off its keys (Ctrl+B with `font: false`).
 * - `toolbar`: false hides the ribbon but keeps editing.
 * @param {HTMLElement} host
 * @param {{ bytes?: Uint8Array | ArrayBuffer, fileName?: string, theme?: 'light' | 'dark', readOnly?: boolean,
 *   features?: Record<string, boolean>, toolbar?: boolean, onChange?: () => void }} [options]
 */
export function mountXlsxEditor(host, options = {}) {
  const root = document.createElement('div');
  root.className = 'oxe-root';
  if (options.theme) root.dataset.theme = options.theme;
  host.appendChild(root);
  const state = { workbook: null, grid: null, tabs: null, ribbon: null, sheetIndex: 0, fileName: null, looks: {}, readOnly: !!options.readOnly, editing: null,
    features: { ...(options.features || {}) }, toolbar: options.toolbar !== false };
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
    showClip();
  };
  /** The outline of what was copied or cut, on its own sheet only. */
  const showClip = () => {
    const c = state.clip;
    state.grid.setClip(c && c.si === state.sheetIndex && c.wb === state.workbook ? c.rect : null);
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
    // (A copy is let go when sheets change: its sheet may have moved or gone.)
    state.editing?.commands.run('dropClip');
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
  /** Whether a formatting command's ribbon group is on (`features`); every other command always is. */
  const allowed = (name, change) => {
    const group = groupOf(name, change || {});
    const tab = group === 'zoom' || group === 'window' ? 'view' : 'home';
    return !group || (state.features[group] !== false && state.features[tab] !== false);
  };
  /** The ribbon's commands: the editor's own, and the sheet actions that ask first or edit a tab in place. */
  const act = (name, ...args) => {
    if (!state.workbook) return false;
    if (!allowed(name, args[0])) return false;
    const i = state.sheetIndex;
    if (name === 'deleteShownSheet') return sheetActs.remove(i);
    if (name === 'renameShownSheet') return sheetActs.rename(i);
    if (name === 'hideShownSheet') return editing().run('hideSheet', i);
    if (name === 'unhideSheet') return sheetActs.unhide();
    if (name === 'insertRows' || name === 'deleteRows') return editing().run('lines', 'row', name === 'insertRows' ? 1 : -1);
    if (name === 'insertColumns' || name === 'deleteColumns') return editing().run('lines', 'col', name === 'insertColumns' ? 1 : -1);
    if (name === 'insertCells' || name === 'deleteCells') return lines(name === 'insertCells' ? 1 : -1);
    if (name === 'zoomBox') {
      return numberBox(root, { title: 'Zoom', label: 'Custom (%):', value: String(state.grid.view().zoom), min: 10, max: 400, refusal: 'Enter a number between 10 and 400.' })
        .then((n) => n != null && editing().run('zoom', n));
    }
    if (name === 'zoomSelection') return editing().run('zoomToSelection');
    if (name === 'find' || name === 'replace') return state.find?.open(name === 'replace');
    if (name === 'rowHeightBox' || name === 'columnWidthBox') return sizeBox(name === 'rowHeightBox' ? 'row' : 'col');
    if (name === 'clipCopy' || name === 'clipCut') return state.editing.copyOut(name === 'clipCut');
    if (name === 'clipPaste') return state.editing.pasteIn();
    if (name === 'merge') {
      // (Excel asks first when values would be discarded.)
      const sel = editing().selection;
      if (mergeLoss(state.workbook.sheet(i), sel, args[0])) {
        return messageBox(root, 'Merging cells only keeps the upper-left value and discards other values.', ['OK', 'Cancel'])
          .then((b) => b === 'OK' && editing().run('merge', args[0]));
      }
    }
    return editing().run(name, ...args);
  };
  /** Ctrl++ (dir 1) and Ctrl+- (-1), Insert Cells... and Delete Cells...: whole rows or columns selected are inserted
   *  or deleted; otherwise Excel's Insert or Delete box (shift cells right or down, left or up, entire row or column),
   *  the shift Excel would make chosen first. */
  const lines = async (dir) => {
    if (state.readOnly || !allowed('lines')) return false;
    const s = editing().selection;
    if (s.c0 === 0 && s.c1 >= MAX_COLS - 1) return editing().run('lines', 'row', dir);
    if (s.r0 === 0 && s.r1 >= MAX_ROWS - 1) return editing().run('lines', 'col', dir);
    const across = dir > 0 ? 'Shift cells right' : 'Shift cells left';
    const down = dir > 0 ? 'Shift cells down' : 'Shift cells up';
    const got = await optionsBox(root, { title: dir > 0 ? 'Insert' : 'Delete',
      groups: [{ key: 'how', legend: dir > 0 ? 'Insert' : 'Delete', options: [across, down, 'Entire row', 'Entire column'], value: shapeShift(s) === 'col' ? across : down }] });
    const how = got?.choices.how;
    if (how === 'Entire row' || how === 'Entire column') return editing().run('lines', how === 'Entire row' ? 'row' : 'col', dir);
    if (how) {
      const axis = how === down ? 'row' : 'col';
      const done = editing().run('cells', dir, axis);
      if (done !== 'unmerge') return done;
      // (Excel asks before unmerging the merged cells the shift would split.)
      if (await messageBox(root, ROWCOL_MESSAGES.unmerge, ['OK', 'Cancel']) === 'OK') return editing().run('cells', dir, axis, { unmerge: true });
    }
    state.grid.el.focus({ preventScroll: true });
    return false;
  };
  /** The menu of a row or column heading. */
  const headerMenu = (axis, point) => {
    const run = (...a) => () => editing().run(...a);
    const ro = state.readOnly;
    menu.open([{ label: 'Insert', run: run('lines', axis, 1), disabled: ro }, { label: 'Delete', run: run('lines', axis, -1), disabled: ro },
      'separator', { label: 'Clear Contents', run: run('clear'), disabled: ro }, 'separator',
      { label: axis === 'row' ? 'Row Height...' : 'Column Width...', run: () => sizeBox(axis), disabled: ro },
      { label: 'Hide', run: run('hideLines', axis, true), disabled: ro }, { label: 'Unhide', run: run('hideLines', axis, false), disabled: ro }],
    point, { label: axis === 'row' ? 'Row' : 'Column' });
  };
  /** Excel's Row Height or Column Width box for the selection, holding the active cell's row height (points) or column
   *  width (characters). */
  const sizeBox = async (axis) => {
    if (state.readOnly) return false;
    const s = editing().selection;
    const g = state.grid.geometry();
    const rows = axis === 'row';
    const value = rows ? g.rows.size(s.ar) * 0.75 : Math.max(0, g.cols.size(s.ac) - 5) / 7;
    const n = await numberBox(root, rows
      ? { title: 'Row Height', label: 'Row height:', value: String(Math.round(value * 100) / 100), max: 409, refusal: 'Row height must be between 0 and 409.' }
      : { title: 'Column Width', label: 'Column width:', value: String(Math.round(value * 100) / 100), max: 255, refusal: 'Column width must be between 0 and 255.' });
    if (n == null) return false;
    return editing().run(rows ? 'rowHeight' : 'columnWidth', n);
  };
  /** The ribbon showing what the active cell has. */
  const lookChanged = () => {
    if (!state.ribbon || !state.workbook || !state.editing) return;
    const wb = state.workbook;
    const look = state.editing.commands.look();
    state.ribbon.update({ ...look, fontName: fontFamily(look.font, wb.theme), fonts: wb.theme.fonts, canUndo: state.editing.commands.canUndo(), canRedo: state.editing.commands.canRedo(), view: state.grid.view() });
  };
  /** The workbook has new cell formats: their CSS again, and the grid drawn with them. */
  const restyled = () => {
    state.looks = applyCellStyles(root, state.workbook);
    state.grid.setLooks(state.looks);
  };
  const placeRibbon = () => {
    if (!state.ribbon) return;
    state.ribbon.el.hidden = state.readOnly || !state.toolbar;
    state.ribbon.setFeatures(state.features);
  };
  const makeRibbon = () => {
    const el = document.createElement('div');
    root.prepend(el);
    state.ribbon = createRibbon(el, {
      tabs: {
        Home: (kit) => homeTab(kit, { act, color: (c) => resolveColor({ rgb: null, theme: null, indexed: null, auto: false, tint: 0, ...c }, state.workbook?.theme, state.workbook?.styles.palette) }),
        View: (kit) => viewTab(kit, { act }),
      },
      home: () => state.grid?.el.focus({ preventScroll: true }),
    });
    placeRibbon();
    // (Find and Replace: a bar under the ribbon, there with the ribbon hidden too.)
    state.find = createFindBar({
      workbook: () => state.workbook,
      from: () => { const s = editing().selection; return { si: state.sheetIndex, r: s.ar, c: s.ac }; },
      go: (hit) => editing().run('goTo', hit),
      apply: (steps) => editing().run('applyChanges', steps),
      home: () => state.grid?.el.focus({ preventScroll: true }),
      readOnly: () => state.readOnly,
    });
    el.after(state.find.el);
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
          lookChanged,
          restyled,
          clipShown: (c) => { state.clip = c; showClip(); },
          reloaded: () => { state.grid.setModel(state.workbook.sheet(state.sheetIndex)); lookChanged(); },
          lines,
          headerMenu,
          fitWidth: (c) => state.grid.fitWidth(c),
          find: (replace) => state.find?.open(replace),
          shownText: (r, c) => state.grid.shownText(r, c),
          allowed,
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
            keepFocus: () => !!state.editing?.keepsFocus(),
          },
        });
      }
      if (!state.ribbon) makeRibbon();
      state.tabs.set(tabList(wb), i);
      state.editing.reset();
      lookChanged();
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
      placeRibbon();
    },
    /** Turn ribbon features on or off (see mountXlsxEditor's `features`). */
    setFeatures(f) {
      state.features = { ...state.features, ...f };
      placeRibbon();
    },
    get features() { return { ...state.features }; },
    setTheme(theme) {
      if (theme) root.dataset.theme = theme;
      else delete root.dataset.theme;
    },
    /** Open the Find bar (Ctrl+F), or Find and Replace (Ctrl+H). */
    find: (replace = false) => state.find?.open(replace),
    destroy() {
      window.removeEventListener('beforeprint', beforePrint);
      menu.close();
      state.ribbon?.destroy();
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
