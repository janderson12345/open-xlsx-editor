// open-xlsx-editor: open .xlsx bytes in a page, show and edit them, save them back. The ribbon over a sheet at a time
// in a virtualised grid, the sheet tabs under it (with their menu: insert, delete, rename, move, colour, hide, unhide).
import { Workbook } from './ooxml/workbook.js';
import { renderPrint } from './editor/print.js';
import { createActs } from './editor/acts.js';
import { printSetup } from './ooxml/printarea.js';
import { headerPictures } from './ooxml/hfpictures.js';
import { applyCellStyles } from './editor/cellstyles.js';
import { fontsChanged } from './editor/fonts.js';
import { useOmf } from './editor/omffonts.js';
import { dropSheet } from './editor/sheets.js';
import { followOf } from './ooxml/links.js';
import { csvWorkbook, isCsv, csvBytes } from './ooxml/csv.js';
import { createParts } from './editor/mountparts.js';
import { groupOf } from './editor/features.js';
import { createStatusBar } from './editor/statusbar.js';
import { createMenu, tabMenuEntries } from './editor/sheetmenu.js';
import { messageBox, unhideBox, passwordBox } from './editor/dialogs.js';
import { PasswordError } from './ooxml/encryption.js';
import { readNotes } from './ooxml/notes.js';
import { breaksRule } from './ooxml/validation.js';
import { printAreaOf } from './ooxml/pagesetupedit.js';
import { sheetProtection, workbookProtection, contentsHidden } from './ooxml/protection.js';
import { conditionalLooks } from './ooxml/condformat.js';
import { sparklineLooks } from './ooxml/sparklines.js';
import { tableLooks } from './ooxml/tablestyle.js';
import { readTables } from './ooxml/tables.js';
import { errorFlags, ERROR_RULES } from './ooxml/errorcheck.js';
import { readCharts } from './ooxml/charts.js';
import { chartFigure } from './editor/chartview.js';
import { createRibbon } from './editor/ribbon.js';
import { createFindBar, pickCell } from './editor/findbar.js';
import { createWatchWindow } from './editor/watchwindow.js';
import { homeTab } from './editor/hometab.js';
import { viewTab } from './editor/viewtab.js';
import { pageLayoutTab } from './editor/pagelayouttab.js';
import { dataTab } from './editor/datatab.js';
import { reviewTab } from './editor/reviewtab.js';
import { insertTab } from './editor/inserttab.js';
import { sparklineTab } from './editor/sparklinetab.js';
import { tableTab } from './editor/tabletab.js';
import { formulasTab } from './editor/formulastab.js';
import { calcMode, calculateBeforeSave } from './ooxml/calcmode.js';
import { fontFamily, formatCodeOf } from './ooxml/styles.js';
import { resolveColor } from './ooxml/theme.js';
import { isDateFormat } from './ooxml/format.js';
import { fillOptionChoices } from './editor/fillopts.js';
import { parseRef } from './ooxml/refs.js';
import { namesOnSheet } from './ooxml/names.js';
import { AgentWorkbook } from './agent.js';
import { createFiles } from './editor/files.js';
import { setLocale, currentLocale, browserLocale } from './ooxml/locale.js';
import { offerBook, openBook } from './editor/pagebooks.js';

export { Workbook };
export { openWorkbook } from './agent.js';
export { PasswordError };
/** The open-metric-fonts release the editor's font faces are for (`fonts: { omf }`; README, Fonts). */
export { OMF_VERSION } from './editor/omf.js';


/**
 * Mount the editor in `host`.
 * - `readOnly`: show the workbook only (no ribbon, nothing editable); `setReadOnly()` changes it.
 * - `features`: turn parts of the ribbon off: a tab by name ({ home: false }) or a group by its label in lower case
 *   with dashes, on any tab ({ font: false, number: false }); `setFeatures()` changes them. A group turned off also
 *   turns off its keys (Ctrl+B with `font: false`).
 * - `toolbar`: false hides the ribbon but keeps editing.
 * - `ribbon`: 'simple' starts with the one-line simplified ribbon, 'classic' with the full one (without it, the one
 *   the user last chose in this browser); the user can still switch with the ribbon's own button. `setRibbon()`
 *   changes it.
 * - `fillLists`: custom lists the fill handle goes round, as Excel's Edit Custom Lists ([['North', 'East', 'South',
 *   'West']]), after its own day and month names.
 * - `statusBar`: false hides the Average, Count and Sum of the selection beside the sheet tabs; an object
 *   ({ minimum: true, maximum: true, numericalCount: true, average, count, sum }) chooses which figures it shows, as
 *   Excel's Customize Status Bar (the status bar right-clicked) does, over the user's own last choice.
 * - `onLink(url)`: a hyperlink clicked that goes outside the workbook (http, https, mailto, ftp); without it, the
 *   browser opens it in a new tab.
 * - `fonts: { omf }`: the URL of an open-metric-fonts web bundle folder the site serves (each package as
 *   <package>/files/, release OMF_VERSION), for the metric twins of the Microsoft fonts the computer lacks (README, Fonts).
 * - `locale`: the country settings numbers and dates are shown and typed in ('de-DE', or { tag, decimal, group,
 *   dateOrder, ... }: ooxml/locale.js, docs/LOCALE.md); the browser's language without it. `setLocale()` changes it.
 * - `onSave({ fileName, type, bytes, mime })`: Save As's file (F12, `saveAs()`); without it, the browser downloads it.
 *   `exportPdf(editor)`: a PDF export (open-xlsx-editor/pdf's exportPdf) for Save As to offer PDF.
 * @param {HTMLElement} host
 * @param {{ bytes?: Uint8Array | ArrayBuffer, fileName?: string, theme?: 'light' | 'dark', readOnly?: boolean,
 *   features?: Record<string, boolean>, toolbar?: boolean, ribbon?: 'simple' | 'classic', statusBar?: boolean | Record<string, boolean>, calculateOnOpen?: boolean,
 *   onChange?: () => void, onLink?: (url: string) => void, fonts?: { omf?: string } }} [options]
 */
export function mountXlsxEditor(host, options = {}) {
  const root = document.createElement('div');
  root.className = 'oxe-root';
  // (The metric twins from the site's own copy of open-metric-fonts, omffonts.js; its folder kept for exportPdf.)
  if (options.fonts?.omf) root.dataset.omf = useOmf(options.fonts.omf);
  if (options.theme) root.dataset.theme = options.theme;
  host.appendChild(root);
  // (The page's country settings, as Excel takes the computer's: ooxml/locale.js.)
  setLocale(options.locale ?? browserLocale());
  const state = { onEdit: new Set(), onSelect: new Set(), circles: new WeakMap(), workbook: null, grid: null, tabs: null, ribbon: null, sheetIndex: 0, fileName: null, looks: {}, readOnly: !!options.readOnly, editing: null,
    features: { ...(options.features || {}) }, toolbar: options.toolbar !== false, ribbonMode: options.ribbon ?? null };
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
  /** Excel's Error Checking Options: `errorChecking` (false: off; { background, rules }) over the user's last choice
   *  in this browser over Excel's defaults (every rule here on). */
  state.errorOptions = (() => {
    let saved = null;
    try { saved = JSON.parse(localStorage.getItem('oxe-error-options') || 'null'); } catch { /* storage blocked */ }
    const given = options.errorChecking === false ? { background: false } : typeof options.errorChecking === 'object' ? options.errorChecking : {};
    return { background: true, rules: [...ERROR_RULES], ...(saved || {}), ...given };
  })();
  /** The error indicators worked out again (errorcheck.js), with the rules turned on; none without background checking. */
  state.refreshErrors = () => {
    if (!state.grid || !state.workbook) return;
    state.workbook.errorRules = state.errorOptions.rules;
    state.grid.setErrors(state.errorOptions.background ? errorFlags(state.workbook, state.sheetIndex) : null);
  };
  const showConditional = () => {
    if (!state.grid || !state.workbook) return;
    state.grid.setConditional(conditionalLooks(state.workbook, state.sheetIndex));
    state.watch?.refresh();
    state.refreshErrors();
    // (Notes too: written, deleted, or moved with rows and columns.)
    state.grid.setNotes(readNotes(state.workbook, state.sheetIndex));
    // (Circle Invalid Data's circles on this sheet, less those whose cells are now put right, as Excel's go.)
    const key = state.workbook.sheets[state.sheetIndex];
    const circled = (state.circles.get(key) || []).filter(({ r, c }) => breaksRule(state.workbook, state.sheetIndex, r, c));
    state.circles.set(key, circled);
    state.grid.setCircles(circled);
    // (The sheet's charts, read again with their data; and, for a screen reader, each as a figure with its data.)
    // (Each with its place among the sheet's charts, which the chart edits go by: chartedit.js.)
    const charts = readCharts(state.workbook, state.sheetIndex).map((c, index) => ({ ...c, index })).filter((c) => c.kind);
    state.charts = charts;
    state.grid.setCharts(charts, state.workbook.theme);
    // (The sheet's sparklines, with their data as it is now.)
    state.grid.setSparklines(sparklineLooks(state.workbook, state.sheetIndex));
    const tables = readTables(state.workbook).filter((t) => t.si === state.sheetIndex);
    state.grid.setTables(tableLooks(state.workbook, tables, state.sheetIndex), tables.filter((t) => t.filter && t.headers).map((t) => ({ name: t.name, filter: t.filter, hidden: t.filter.hidden })));
    let figures = root.querySelector(':scope > .oxe-chart-figures');
    if (!figures && charts.length) { figures = document.createElement('section'); figures.className = 'oxe-chart-figures'; figures.setAttribute('aria-label', 'Charts on this sheet'); root.append(figures); }
    figures?.replaceChildren(...charts.map(chartFigure));
  };
  /** The ribbon showing what the active cell has. */
  const lookChanged = () => {
    if (!state.ribbon || !state.workbook || !state.editing) return;
    const wb = state.workbook;
    const look = state.editing.commands.look();
    state.ribbon.update({ ...look, fontName: fontFamily(look.font, wb.theme), fonts: wb.theme.fonts, canUndo: state.editing.commands.canUndo(), canRedo: state.editing.commands.canRedo(), view: { ...state.grid.view(), formulaBar: !root.querySelector('.oxe-bar')?.hidden }, sheetProtected: !!sheetProtection(wb, state.sheetIndex), bookProtected: !!workbookProtection(wb)?.structure });
    showStats();
  };
  // (The status bar's figures and its Customize Status Bar menu: editor/statusbar.js.)
  const status = createStatusBar({ state, options, menu });
  const showStats = status.show;
  const statusMenu = status.menu;
  /** The workbook's cell formats as the grid draws them (cellstyles.js), and which cells of the sheet shown are
   *  hidden by its protection (ooxml/protection.js contentsHidden). */
  const looksOf = (wb) => ({ ...applyCellStyles(root, wb), hiddenAt: (r, c) => contentsHidden(wb, state.sheetIndex, r, c) });
  /** The workbook has new cell formats: their CSS again, and the grid drawn with them. */
  const restyled = () => {
    state.looks = looksOf(state.workbook);
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
        Home: (kit) => homeTab(kit, { act, workbook: () => state.workbook, color: (c) => resolveColor({ rgb: null, theme: null, indexed: null, auto: false, tint: 0, ...c }, state.workbook?.theme, state.workbook?.styles.palette) }),
        Insert: (kit) => insertTab(kit, { act }),
        'Page Layout': (kit) => pageLayoutTab(kit, { act, hasPrintArea: () => !!(state.workbook && printAreaOf(state.workbook, state.sheetIndex)), page: () => (state.workbook ? state.workbook.sheet(state.sheetIndex).page : null) }),
        Formulas: (kit) => formulasTab(kit, { act, recent: () => state.recentFunctions, calcMode: () => calcMode(state.workbook), names: () => (state.workbook ? namesOnSheet(state.workbook, state.sheetIndex).map((n) => n.name) : []) }),
        Data: (kit) => dataTab(kit, { act }),
        Review: (kit) => reviewTab(kit, { act }),
        View: (kit) => viewTab(kit, { act }),
        'Table Design': (kit) => tableTab(kit, { act, workbook: () => state.workbook }),
        Sparkline: (kit) => sparklineTab(kit, { act, color: (c) => resolveColor({ rgb: null, theme: null, indexed: null, auto: false, tint: 0, ...c }, state.workbook?.theme, state.workbook?.styles.palette) }),
      },
      // (Not while a cell is being edited: a function from the Function Library was just typed into it; nor while part of
      // the formula bar's text is selected, which the Font group has just formatted.)
      home: () => { if (!state.editing?.editorOpen && !state.editing?.textSelected) state.grid?.el.focus({ preventScroll: true }); },
      // (The layout the page asked for, then the user's switch kept when the ribbon is made again.)
      mode: state.ribbonMode,
      onMode: (m) => { state.ribbonMode = m; },
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
      pickCell: () => pickCell(state.grid.el, () => { const sel = editing().selection; return { si: state.sheetIndex, r: sel.ar, c: sel.ac }; }),
      root,
    });
    el.after(state.find.el);
    // (Formulas > Watch Window: docked under the Find bar, kept up to date as the workbook changes.)
    state.watch = createWatchWindow({
      workbook: () => state.workbook,
      bookName: () => state.fileName || '',
      sheetIndex: () => state.sheetIndex,
      selection: () => editing().selection,
      go: (hit) => editing().run('goTo', hit),
      readOnly: () => state.readOnly,
      changed: () => {},
      home: () => state.grid?.el.focus({ preventScroll: true }),
    });
    state.find.el.after(state.watch.el);
  };

  /** Auto Fill Options' menu: the last fill done again another way (fillopts.js): undone, the cells it started from
   *  selected again, and filled with the type chosen, as one more step. Date choices when those cells hold dates. */
  const fillOptionsMenu = (point) => {
    const fill = state.lastFill;
    if (!fill) return;
    const wb = state.workbook;
    const sheet = wb.sheet(state.sheetIndex);
    let dated = false;
    for (let r = fill.from.r0; r <= fill.from.r1; r++) {
      for (let c = fill.from.c0; c <= fill.from.c1; c++) {
        const cell = sheet.cells.get(r)?.get(c);
        if (typeof cell?.v === 'number' && cell.t !== 's' && isDateFormat(formatCodeOf(wb.styles, cell.s || 0))) dated = true;
      }
    }
    menu.open(fillOptionChoices(dated).map(({ type, label }) => ({ label, run: () => {
      editing().run('undo');
      editing().run('select', fill.from.r0, fill.from.c0);
      editing().run('extendTo', fill.from.r1, fill.from.c1);
      editing().run('autoFill', fill.to, { type });
    } })), point, { label: 'Auto Fill Options' });
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
  // (A hyperlink followed: an address opened by the browser, or onLink; a place in the workbook gone to, as the Name
  // Box goes; one a page can't open refused in Excel's words. ooxml/links.js.)
  const followLink = (h) => {
    const to = followOf(h);
    if (to.refused) return alert(to.refused);
    if (to.place) return editing().run('nameBox', to.place);
    if (options.onLink) options.onLink(to.url);
    else window.open(to.url, '_blank', 'noopener,noreferrer');
  };
  // (Text Import, Save As and the locale: editor/files.js.)
  const files = createFiles({ root, state, options, editing, api: () => api });
  const { act, lines, askInvalid, showValidation, listMenu, openFilter, filterMenu, cellMenu, chartMenu, headerMenu } = createActs({ root, state, options, at, editing, menu, sheetActs, allowed, followLink, files });
  const api = {
    /** Open a workbook (or a .csv, by its `fileName`) and show the sheet it opens on (its active tab, or the first
     *  visible worksheet). One saved
     *  with a password opens with `password`; without one, Excel's Password box asks for it and a promise is
     *  returned (the editor once open, or null: cancelled, or the password wrong, said in Excel's words). */
    open(bytes, fileName = null, { password = null } = {}) {
      let wb;
      const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      try {
        // (A .csv: one sheet, each field as if typed, as Excel opens one; ooxml/csv.js.)
        wb = isCsv(data, fileName) ? csvWorkbook(data, fileName) : Workbook.fromBytes(data, { password });
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
      // (Its name, for CELL("address") of another sheet: [name]Sheet!$A$1, as Excel gives it.)
      wb.fileName = fileName;
      // (The other editors' workbooks on the page, by file name, for links to them: editor/pagebooks.js.)
      wb.openBook = openBook;
      // (Formulas saved without a result, as some programs save them, worked out now, as Excel does on opening.)
      if (options.calculateOnOpen !== false) wb.calculateMissing();
      const i = wb.activeSheet;
      const sheet = wb.sheet(i);
      state.workbook = wb;
      // (Custom lists the fill handle goes round, as Excel's Edit Custom Lists: fillops.js.)
      if (options.fillLists) wb.fillLists = options.fillLists;
      state.sheetIndex = i;
      state.fileName = fileName;
      state.looks = looksOf(wb);
      if (state.grid) { state.grid.setSheet(sheet, wb.sheets[i].name, state.looks); state.grid.setNotes(readNotes(wb, i)); showConditional(); }
      else {
        // (The grid, its editing and the sheet tabs, made with the first workbook: editor/mountparts.js.)
        createParts({ root, state, options, act, openFilter, fillOptionsMenu, listMenu, chartMenu, showConditional, showSheet,
          sheetsChanged, alert, followLink, editing, lookChanged, restyled, showClip, lines, headerMenu, cellMenu, filterMenu,
          askInvalid, showValidation, allowed, openMenu, statusMenu, sheetActs }, wb, i, sheet);
      }
      if (!state.ribbon) makeRibbon();
      state.tabs.set(tabList(wb), i);
      state.editing.reset();
      lookChanged();
      return api;
    },
    /** The workbook as .xlsx bytes (unchanged parts byte for byte, changed sheets rewritten in place). */
    /** The workbook as a file's bytes. `as`: 'xlsx', 'xltx', 'xlsm' or 'xltm', made that kind of file from now on, as
     *  Excel's Save As (a template saved as a workbook; ooxml/filetype.js; a workbook with macros only as a macro
     *  kind). */
    save({ as = null } = {}) {
      if (!state.workbook) throw new Error('no workbook open');
      if (as) state.workbook.setFileType(as);
      // (A manual workbook worked out first, unless Recalculate workbook before saving is off: calcmode.js.)
      if (calculateBeforeSave(state.workbook)) { state.grid.refresh(); showStats(); }
      return state.workbook.toBytes();
    },
    /** The sheet shown as a CSV, as Excel's Save As CSV UTF-8 writes one (`utf8: false`: as CSV, code page 1252;
     *  `kind` 'txt' or 'unicode' as Text (Tab delimited) or Unicode Text; or `sep`, `encoding` and `bom`):
     *  ooxml/csv.js csvBytes. */
    saveCsv(how = {}) {
      if (!state.workbook) throw new Error('no workbook open');
      if (calculateBeforeSave(state.workbook)) { state.grid.refresh(); showStats(); }
      return csvBytes(state.workbook, state.sheetIndex, how);
    },
    /** A text file through the Text Import box, into a new workbook (`into: 'new'`) or at the active cell ('sheet'):
     *  editor/files.js importText. */
    importText: (bytes, fileName, how) => files.importText(bytes, fileName, how),
    /** Save As (F12): the box, then the file to `onSave` or downloaded; resolves to { fileName, type, bytes, mime }. */
    saveAs: (choice) => files.saveAs(choice),
    /** The page's country settings (docs/LOCALE.md), the sheet drawn again in them. */
    setLocale: (spec) => files.setLocale(spec),
    get locale() { return currentLocale(); },
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
      return renderPrint(root, wb.sheet(state.sheetIndex), state.looks, fontFamily(wb.styles.normal, wb.theme), { ...printSetup(wb, state.sheetIndex), sheet: wb.sheets[state.sheetIndex].name, file: state.fileName || '', sizePt: wb.styles.normal?.size || 11, pictures: headerPictures(wb, state.sheetIndex), sparks: sparklineLooks(wb, state.sheetIndex), tables: tableLooks(wb, readTables(wb), state.sheetIndex) });
    },
    get workbook() { return state.workbook; },
    get grid() { return state.grid; },
    get sheetIndex() { return state.sheetIndex; },
    get fileName() { return state.fileName; },
    /** The open file's kind: 'xlsx', 'xltx' (a template), 'xlsm' or 'xltm' (null for a .csv). */
    get fileType() { return state.workbook?.fileType ?? null; },
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
    /** The ribbon's layout: 'simple' (one line) or 'classic'. */
    setRibbon(mode) {
      state.ribbonMode = mode === 'simple' ? 'simple' : 'classic';
      state.ribbon?.setMode(state.ribbonMode);
    },
    get ribbon() { return state.ribbon ? state.ribbon.mode : state.ribbonMode; },
    setTheme(theme) {
      if (theme) root.dataset.theme = theme;
      else delete root.dataset.theme;
    },
    /** Open the Find bar (Ctrl+F), or Find and Replace (Ctrl+H). */
    find: (replace = false) => state.find?.open(replace),
    destroy() {
      window.removeEventListener('beforeprint', beforePrint);
      document.fonts?.removeEventListener('loadingdone', fontsLoaded);
      menu.close();
      state.ribbon?.destroy();
      state.editing?.destroy();
      unoffer();
      state.grid?.destroy();
      dropSheet(root);
      root.remove();
    },
  };
  const unoffer = offerBook({ workbook: () => state.workbook, fileName: () => state.fileName });
  const beforePrint = () => api.preparePrint();
  window.addEventListener('beforeprint', beforePrint);
  // (A web font the page serves, a metric twin among them (fonts.js), arrives after the first drawing: the cells are
  // placed and measured again when it does.)
  const fontsLoaded = () => { if (state.workbook && fontsChanged()) restyled(); };
  document.fonts?.addEventListener('loadingdone', fontsLoaded);
  if (options.bytes) api.open(options.bytes, options.fileName || null);
  return api;
}
