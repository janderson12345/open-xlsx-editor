// open-xlsx-editor: open .xlsx bytes in a page, show and edit them, save them back. The ribbon over a sheet at a time
// in a virtualised grid, the sheet tabs under it (with their menu: insert, delete, rename, move, colour, hide, unhide).
import { Workbook } from './ooxml/workbook.js';
import { createGrid } from './editor/grid.js';
import { renderPrint } from './editor/print.js';
import { createActs } from './editor/acts.js';
import { chartPicker } from './editor/chartpick.js';
import { printSetup } from './ooxml/printarea.js';
import { applyCellStyles } from './editor/cellstyles.js';
import { attachEditing } from './editor/interact.js';
import { createSheetTabs } from './editor/tabs.js';
import { createMenu, tabMenuEntries } from './editor/sheetmenu.js';
import { messageBox, unhideBox, passwordBox } from './editor/dialogs.js';
import { PasswordError } from './ooxml/encryption.js';
import { readNotes } from './ooxml/notes.js';
import { conditionalLooks } from './ooxml/condformat.js';
import { errorFlags } from './ooxml/errorcheck.js';
import { readCharts } from './ooxml/charts.js';
import { chartFigure } from './editor/chartview.js';
import { createRibbon } from './editor/ribbon.js';
import { createFindBar } from './editor/findbar.js';
import { homeTab } from './editor/hometab.js';
import { viewTab } from './editor/viewtab.js';
import { dataTab } from './editor/datatab.js';
import { reviewTab } from './editor/reviewtab.js';
import { insertTab } from './editor/inserttab.js';
import { formulasTab } from './editor/formulastab.js';
import { calcMode, needsCalculation, calculateBeforeSave } from './ooxml/calcmode.js';
import { fontFamily } from './ooxml/styles.js';
import { resolveColor } from './ooxml/theme.js';
import { selectionStats } from './ooxml/stats.js';
import { parseRef } from './ooxml/refs.js';
import { AgentWorkbook } from './agent.js';

export { Workbook };
export { openWorkbook } from './agent.js';
export { PasswordError };

/** The ribbon group a formatting command belongs to, for `features`: one of a group turned off doesn't run, from the
 *  ribbon or the keyboard. */
const FEATURE_OF = { toggleFont: 'font', stepFont: 'font', borders: 'font', toggleWrap: 'alignment', stepIndent: 'alignment', align: 'alignment', merge: 'alignment',
  resizeLines: 'cells', columnWidth: 'cells', rowHeight: 'cells', fitLines: 'cells', hideLines: 'cells', rowHeightBox: 'cells', columnWidthBox: 'cells',
  freeze: 'window', zoom: 'zoom', zoomBox: 'zoom', zoomSelection: 'zoom', zoomToSelection: 'zoom',
  lines: 'cells', cells: 'cells', insertCells: 'cells', deleteCells: 'cells', insertRows: 'cells', deleteRows: 'cells', insertColumns: 'cells', deleteColumns: 'cells',
  clipCopy: 'clipboard', clipCut: 'clipboard', clipPaste: 'clipboard', copy: 'clipboard', cut: 'clipboard', paste: 'clipboard',
  pasteSpecial: 'clipboard', pasteSpecialBox: 'clipboard', pasteAs: 'clipboard', formatPainter: 'clipboard', paintFormats: 'clipboard',
  formatCells: 'font', formatCellsBox: 'font',
  stepDecimal: 'number', clearFormats: 'editing', clearAs: 'editing', fill: 'editing', autoFill: 'editing',
  sort: 'editing', sortAsc: 'editing', sortDesc: 'editing', sortBox: 'editing', autoSum: 'editing',
  filterToggle: 'editing', filterBy: 'editing', filterClear: 'editing', filterReapply: 'editing', filterAll: 'editing', filterSort: 'editing' };
const groupOf = (name, change) => (name !== 'format' ? FEATURE_OF[name] : change.numFmt != null || change.style ? 'number' : change.align ? 'alignment' : 'font');

/**
 * Mount the editor in `host`.
 * - `readOnly`: show the workbook only (no ribbon, nothing editable); `setReadOnly()` changes it.
 * - `features`: turn parts of the ribbon off: a tab by name ({ home: false }) or a group by its label in lower case
 *   with dashes, on any tab ({ font: false, number: false }); `setFeatures()` changes them. A group turned off also
 *   turns off its keys (Ctrl+B with `font: false`).
 * - `toolbar`: false hides the ribbon but keeps editing.
 * - `statusBar`: false hides the Average, Count and Sum of the selection beside the sheet tabs.
 * @param {HTMLElement} host
 * @param {{ bytes?: Uint8Array | ArrayBuffer, fileName?: string, theme?: 'light' | 'dark', readOnly?: boolean,
 *   features?: Record<string, boolean>, toolbar?: boolean, statusBar?: boolean, calculateOnOpen?: boolean,
 *   onChange?: () => void }} [options]
 */
export function mountXlsxEditor(host, options = {}) {
  const root = document.createElement('div');
  root.className = 'oxe-root';
  if (options.theme) root.dataset.theme = options.theme;
  host.appendChild(root);
  const state = { onEdit: new Set(), onSelect: new Set(), workbook: null, grid: null, tabs: null, ribbon: null, sheetIndex: 0, fileName: null, looks: {}, readOnly: !!options.readOnly, editing: null,
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
    state.grid.setNotes(readNotes(state.workbook, i));
    showConditional();
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
    moveCopy: (i) => act('moveCopyBox', i),
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
  /** The shown sheet's conditional formats worked out again (after a change, or another sheet shown). */
  const showConditional = () => {
    if (!state.grid || !state.workbook) return;
    state.grid.setConditional(conditionalLooks(state.workbook, state.sheetIndex));
    // (Excel's error indicators, worked out again: errorcheck.js; `errorChecking: false` turns them off.)
    state.grid.setErrors(options.errorChecking === false ? null : errorFlags(state.workbook, state.sheetIndex));
    // (Notes too: written, deleted, or moved with rows and columns.)
    state.grid.setNotes(readNotes(state.workbook, state.sheetIndex));
    // (The sheet's charts, read again with their data; and, for a screen reader, each as a figure with its data.)
    // (Each with its place among the sheet's charts, which the chart edits go by: chartedit.js.)
    const charts = readCharts(state.workbook, state.sheetIndex).map((c, index) => ({ ...c, index })).filter((c) => c.kind);
    state.charts = charts;
    state.grid.setCharts(charts, state.workbook.theme);
    let figures = root.querySelector(':scope > .oxe-chart-figures');
    if (!figures && charts.length) { figures = document.createElement('section'); figures.className = 'oxe-chart-figures'; figures.setAttribute('aria-label', 'Charts on this sheet'); root.append(figures); }
    figures?.replaceChildren(...charts.map(chartFigure));
  };
  /** The ribbon showing what the active cell has. */
  const lookChanged = () => {
    if (!state.ribbon || !state.workbook || !state.editing) return;
    const wb = state.workbook;
    const look = state.editing.commands.look();
    state.ribbon.update({ ...look, fontName: fontFamily(look.font, wb.theme), fonts: wb.theme.fonts, canUndo: state.editing.commands.canUndo(), canRedo: state.editing.commands.canRedo(), view: state.grid.view() });
    showStats();
  };
  /** The status bar: Excel's Average, Count and Sum of the selection (stats.js), or Count alone. */
  const showStats = () => {
    if (!state.tabs || !state.editing) return;
    if (options.statusBar === false) { state.tabs.status('', true); return; }
    const s = selectionStats(state.workbook, state.sheetIndex, state.editing.commands.selection);
    const parts = !s ? [] : s.text ? [`Average: ${s.text.average}`, `Count: ${s.count}`, `Sum: ${s.text.sum}`] : [`Count: ${s.count}`];
    // (Manual calculation with formulas waiting: Excel's Calculate.)
    if (needsCalculation(state.workbook)) parts.unshift('Calculate');
    state.tabs.status(parts.join('    '));
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
        Insert: (kit) => insertTab(kit, { act }),
        Formulas: (kit) => formulasTab(kit, { act, recent: () => state.recentFunctions, calcMode: () => calcMode(state.workbook) }),
        Data: (kit) => dataTab(kit, { act }),
        Review: (kit) => reviewTab(kit, { act }),
        View: (kit) => viewTab(kit, { act }),
      },
      // (Not while a cell is being edited: a function from the Function Library was just typed into it.)
      home: () => { if (!state.editing?.editorOpen) state.grid?.el.focus({ preventScroll: true }); },
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

  // (The ribbon's actions and their boxes and menus, editor/acts.js, given this editor's state and helpers.)
  const { act, lines, askInvalid, showValidation, listMenu, openFilter, filterMenu, cellMenu, chartMenu, headerMenu } = createActs({ root, state, options, at, editing, menu, sheetActs, allowed });
  const api = {
    /** Open a workbook and show the sheet it opens on (its active tab, or the first visible worksheet). One saved
     *  with a password opens with `password`; without one, Excel's Password box asks for it and a promise is
     *  returned (the editor once open, or null: cancelled, or the password wrong, said in Excel's words). */
    open(bytes, fileName = null, { password = null } = {}) {
      let wb;
      try {
        wb = Workbook.fromBytes(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), { password });
      } catch (e) {
        if (!(e instanceof PasswordError) || e.reason !== 'needed') throw e;
        return passwordBox(root, fileName).then((pw) => {
          if (pw == null) return null;
          try { return api.open(bytes, fileName, { password: pw }); } catch (err) {
            if (!(err instanceof PasswordError)) throw err;
            return messageBox(root, err.message).then(() => null);
          }
        });
      }
      // (Formulas saved without a result, as some programs save them, worked out now, as Excel does on opening.)
      if (options.calculateOnOpen !== false) wb.calculateMissing();
      const i = wb.activeSheet;
      const sheet = wb.sheet(i);
      state.workbook = wb;
      state.sheetIndex = i;
      state.fileName = fileName;
      state.looks = applyCellStyles(root, wb);
      if (state.grid) { state.grid.setSheet(sheet, wb.sheets[i].name, state.looks); state.grid.setNotes(readNotes(wb, i)); showConditional(); }
      else {
        state.grid = createGrid(root, { sheet, label: wb.sheets[i].name, looks: state.looks, readOnly: state.readOnly });
        state.grid.onFilterButton((c, button) => openFilter(c, button.getBoundingClientRect()));
        state.grid.onErrorButton((point) => act('errorMenu', point));
        state.grid.onDropdown((button) => listMenu(button.getBoundingClientRect()));
        state.picker = chartPicker({ grid: state.grid, charts: () => state.charts || [], act, menu: (point, k) => chartMenu(point, k), readOnly: () => state.readOnly });
        state.grid.onChartsDrawn(() => state.picker?.redrawn());
        state.grid.setNotes(readNotes(wb, i));
        showConditional();
        state.editing = attachEditing(root, {
          workbook: () => state.workbook,
          sheetIndex: () => state.sheetIndex,
          showSheet,
          sheetsChanged,
          alert,
          grid: state.grid,
          looks: () => state.looks,
          readOnly: () => state.readOnly,
          changed: () => { showConditional(); options.onChange?.(); },
          edited: (e) => { for (const f of [...state.onEdit]) f(e); },
          lookChanged,
          restyled,
          clipShown: (c) => { state.clip = c; showClip(); },
          reloaded: () => { state.grid.setModel(state.workbook.sheet(state.sheetIndex)); lookChanged(); },
          lines,
          headerMenu,
          cellMenu,
          pasteSpecial: () => act('pasteSpecialBox'),
          autoSum: (fn) => act('autoSum', fn),
          filterMenu: () => filterMenu() || listMenu(),
          askInvalid: (info) => askInvalid(info),
          selectionShown: () => { for (const f of [...state.onSelect]) f(); return showValidation(); },
          goToBox: () => act('goToBox'),
          nameManager: () => act('nameManager'),
          insertFunction: () => act('insertFunction'),
          errorMenu: () => { const at = state.grid.errorButtonAt(); return at ? act('errorMenu', at) : false; },
          painting: () => act('painting'),
          commentBox: () => act('commentBox'),
          flashFill: () => act('flashFill'),
          chartKey: (e) => state.picker?.onKey(e),
          formatCells: () => act('formatCellsBox'),
          painted: () => act('paintFormats'),
          toggleNote: () => { const s = editing().selection; return state.grid.toggleNote(s.ar, s.ac); },
          editNote: () => act('editNote'),
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
      // (A manual workbook worked out first, unless Recalculate workbook before saving is off: calcmode.js.)
      if (calculateBeforeSave(state.workbook)) { state.grid.refresh(); showStats(); }
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
    /** The workbook open here for an agent (agent.js): read it, and edit it with each edit one step of the undo
     *  history, drawn as it is made. */
    agent() {
      if (!state.workbook) throw new Error('no workbook open');
      // (Its changes() compare with the workbook as it is now, before this agent's edits.)
      const was = state.workbook.toBytes();
      const password = state.workbook.encryption?.password ?? null;
      return new AgentWorkbook(() => state.workbook, {
        cells: (si, contents, styles) => {
          if (!editing().run('applyChanges', [{ si, contents, styles }])) throw new Error(state.readOnly ? 'The workbook is read-only' : "You can't change part of an array");
        },
        structure: (run, kind) => editing().run('agentChange', run, kind),
      }, { baseline: () => Workbook.fromBytes(was, { password }) });
    },
    /** The password the workbook is saved with (a new one, as Excel's Encrypt with Password), or null to save it
     *  without one. */
    setPassword(password) {
      if (!state.workbook) throw new Error('no workbook open');
      state.workbook.setPassword(password || null);
      options.onChange?.();
    },
    /** Whether the workbook open is saved with a password. */
    get encrypted() { return !!state.workbook?.encrypted; },
    /** Call `fn` after each edit made here (not one from another editor of a shared workbook): { cells: [{ si,
     *  contents }] } for cell contents alone, { book: true } for anything else. Returns a function that stops it. */
    onEdit(fn) { state.onEdit.add(fn); return () => state.onEdit.delete(fn); },
    /** Call `fn` each time the selection here is shown anew (moved, extended, another sheet shown). Returns a function
     *  that stops it. */
    onSelect(fn) { state.onSelect.add(fn); return () => state.onSelect.delete(fn); },
    /** Run a named command (commands.js) with its arguments. */
    command: (name, ...args) => editing().run(name, ...args),
    get selection() { return state.editing ? state.editing.commands.selection : null; },
    /** Build the printed pages of the sheet shown (done before the browser prints, too); returns the page count. */
    preparePrint() {
      if (!state.workbook) return 0;
      const wb = state.workbook;
      return renderPrint(root, wb.sheet(state.sheetIndex), state.looks, fontFamily(wb.styles.normal, wb.theme), { ...printSetup(wb, state.sheetIndex), sheet: wb.sheets[state.sheetIndex].name, file: state.fileName || '', sizePt: wb.styles.normal?.size || 11 });
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
