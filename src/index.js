// open-xlsx-editor: open .xlsx bytes in a page, show and edit them, save them back. The ribbon over a sheet at a time
// in a virtualised grid, the sheet tabs under it (with their menu: insert, delete, rename, move, colour, hide, unhide).
import { Workbook } from './ooxml/workbook.js';
import { createGrid } from './editor/grid.js';
import { renderPrint } from './editor/print.js';
import { printSetup } from './ooxml/printarea.js';
import { applyCellStyles } from './editor/cellstyles.js';
import { attachEditing } from './editor/interact.js';
import { createSheetTabs } from './editor/tabs.js';
import { createMenu, tabMenuEntries } from './editor/sheetmenu.js';
import { messageBox, unhideBox, numberBox, optionsBox, sortBox, filterBox, pickBox, goToBox, passwordBox, noteBox } from './editor/dialogs.js';
import { setNote, deleteNote } from './ooxml/noteedit.js';
import { PasswordError } from './ooxml/encryption.js';
import { goToNames } from './ooxml/goto.js';
import { filterChoices, shownText } from './ooxml/filterops.js';
import { readNotes } from './ooxml/notes.js';
import { conditionalLooks } from './ooxml/condformat.js';
import { ruleAt, listItems } from './ooxml/validation.js';
import { readCharts } from './ooxml/charts.js';
import { chartFigure } from './editor/chartview.js';
import { dataBeside, sortArea, guessHeader, SORT_MESSAGES } from './ooxml/sortops.js';
import { isDateFormat } from './ooxml/format.js';
import { shapeShift, ROWCOL_MESSAGES } from './ooxml/rowcol.js';
import { createRibbon } from './editor/ribbon.js';
import { createFindBar } from './editor/findbar.js';
import { homeTab } from './editor/hometab.js';
import { viewTab } from './editor/viewtab.js';
import { dataTab } from './editor/datatab.js';
import { reviewTab } from './editor/reviewtab.js';
import { fontFamily, formatCodeOf } from './ooxml/styles.js';
import { resolveColor } from './ooxml/theme.js';
import { mergeLoss } from './ooxml/mergeops.js';
import { selectionStats } from './ooxml/stats.js';
import { MAX_ROWS, MAX_COLS, colName } from './ooxml/refs.js';
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
  pasteSpecial: 'clipboard', pasteSpecialBox: 'clipboard', pasteAs: 'clipboard',
  stepDecimal: 'number', clearFormats: 'editing', clearAs: 'editing', fill: 'editing', autoFill: 'editing',
  sort: 'editing', sortAsc: 'editing', sortDesc: 'editing', sortBox: 'editing', autoSum: 'editing',
  filterToggle: 'editing', filterBy: 'editing', filterClear: 'editing', filterReapply: 'editing', filterAll: 'editing', filterSort: 'editing' };
/** The Paste menu's choices, as Paste Special options. */
const PASTE_AS = { formulas: { paste: 'formulas' }, values: { paste: 'values' }, formats: { paste: 'formats' }, transpose: { paste: 'all', transpose: true }, link: { link: true } };
/** Excel's Paste Special box: its Paste and Operation choices, by label, as clipops.js pasteSpecial's options. */
const PASTE_CHOICES = { All: 'all', Formulas: 'formulas', Values: 'values', Formats: 'formats', 'All except borders': 'noBorders',
  'Column widths': 'widths', 'Formulas and number formats': 'formulasFormats', 'Values and number formats': 'valuesFormats' };
const OPERATIONS = { None: null, Add: 'add', Subtract: 'subtract', Multiply: 'multiply', Divide: 'divide' };
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
 *   features?: Record<string, boolean>, toolbar?: boolean, statusBar?: boolean, onChange?: () => void }} [options]
 */
export function mountXlsxEditor(host, options = {}) {
  const root = document.createElement('div');
  root.className = 'oxe-root';
  if (options.theme) root.dataset.theme = options.theme;
  host.appendChild(root);
  const state = { onEdit: new Set(), workbook: null, grid: null, tabs: null, ribbon: null, sheetIndex: 0, fileName: null, looks: {}, readOnly: !!options.readOnly, editing: null,
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
    if (name === 'pasteSpecialBox') return pasteSpecialBox();
    if (name === 'sortAsc' || name === 'sortDesc') return sortQuick(name === 'sortDesc');
    if (name === 'sortBox') return customSort();
    if (name === 'editNote') {
      // (Excel's New Note and Edit Note: the active cell's note in a box to type in, beside the cell.)
      const s = editing().selection;
      const n = readNotes(state.workbook, i).get(`${s.ar},${s.ac}`);
      if (n?.kind === 'thread') return messageBox(root, "Threaded comments can't be edited here yet.");
      const body = n ? n.runs.filter((x, k) => !(k === 0 && x.bold)).map((x) => x.text).join('').replace(/^\n/, '') : '';
      const author = n ? n.author || (n.runs[0]?.bold ? n.runs[0].text.replace(/:$/, '') : '') : options.author || 'Author';
      const box = (root.querySelector('.oxe-sel-active') || state.grid.el).getBoundingClientRect();
      const ref = `${colName(s.ac)}${s.ar + 1}`;
      return noteBox(root, { title: n ? `Edit Note for ${ref}` : `New Note for ${ref}`, author, text: body, at: { x: box.right + 8, y: box.top } }).then((text) => {
        if (text != null && (n || text !== '')) editing().run('notesChange', (wb, si) => setNote(wb, si, s.ar, s.ac, text, options.author || 'Author'));
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'deleteNote') {
      const s = editing().selection;
      return editing().run('notesChange', (wb, si) => deleteNote(wb, si, s.ar, s.ac));
    }
    if (name === 'showNote') { const s = editing().selection; return state.grid.toggleNote(s.ar, s.ac); }
    if (name === 'goToBox') {
      // (Excel's Go To: a name, a table or a reference; gone to as the Name Box goes.)
      return goToBox(root, goToNames(state.workbook, i)).then((text) => {
        if (text) editing().run('nameBox', text);
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'autoSum') {
      // (Where Excel leaves the formula being edited, the cell editor opens with it, waiting for Enter.)
      const done = editing().run('autoSum', args[0]);
      if (done?.edit) { state.editing.editWith(done.edit); return true; }
      return done;
    }
    if (name === 'pasteAs') return editing().run('pasteSpecial', PASTE_AS[args[0]]);
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
  /** Excel's Paste Special box (Ctrl+Alt+V, Paste > Paste Special...): Paste, Operation, Skip blanks, Transpose, and its
   *  Paste Link button. */
  const pasteSpecialBox = async () => {
    if (state.readOnly || !allowed('pasteSpecial')) return false;
    const got = await optionsBox(root, { title: 'Paste Special',
      groups: [{ key: 'paste', legend: 'Paste', options: Object.keys(PASTE_CHOICES), value: 'All', columns: 2 },
        { key: 'op', legend: 'Operation', options: Object.keys(OPERATIONS), value: 'None', columns: 2 }],
      checks: [{ key: 'skip', label: 'Skip blanks' }, { key: 'transpose', label: 'Transpose' }], buttons: ['Paste Link'] });
    if (!got) { state.grid.el.focus({ preventScroll: true }); return false; }
    if (got.button === 'Paste Link') return editing().run('pasteSpecial', { link: true });
    return editing().run('pasteSpecial', { paste: PASTE_CHOICES[got.choices.paste], op: OPERATIONS[got.choices.op], ...got.checks });
  };
  /** Before a sort of a selection with data beside it, Excel's Sort Warning: true to take that data in too, false to
   *  sort the selection alone, null when cancelled (COM doesn't reach the box; its words and its choice of Expand). */
  const sortWarning = async () => {
    if (!dataBeside(state.workbook.sheet(state.sheetIndex), editing().selection)) return false;
    const got = await optionsBox(root, { title: 'Sort Warning', ok: 'Sort',
      text: 'There is data next to your selection. Since you have not selected this data, it will not be sorted.',
      groups: [{ key: 'what', legend: 'What do you want to do?', options: ['Expand the selection', 'Continue with the current selection'], value: 'Expand the selection' }] });
    if (!got) { state.grid.el.focus({ preventScroll: true }); return null; }
    return got.choices.what === 'Expand the selection';
  };
  /** Sort A to Z and Sort Z to A (`desc`), by the active cell's column (sortops.js quickSort). */
  const sortQuick = async (desc) => {
    if (state.readOnly || !allowed('sort')) return false;
    const expand = await sortWarning();
    return expand != null && editing().run('sort', desc ? 'desc' : 'asc', { expand });
  };
  /** Custom Sort...: Excel's Sort box over what A to Z would sort, the headers guessed as Excel guesses them; its
   *  columns named from the header row (Column B where there is none), each one's orders by what it holds. */
  const customSort = async () => {
    if (state.readOnly || !allowed('sort')) return false;
    const expand = await sortWarning();
    if (expand == null) return false;
    const wb = state.workbook;
    const si = state.sheetIndex;
    const sheet = wb.sheet(si);
    const sel = editing().selection;
    const area = (expand && dataBeside(sheet, sel)) || sortArea(sheet, sel);
    if (!area) return messageBox(root, SORT_MESSAGES.empty).then(() => false);
    const at = (r, c) => sheet.cells.get(r)?.get(c);
    const kindOf = (cells) => {
      const cell = cells.find((x) => x && x.v != null && x.v !== '');
      if (typeof cell?.v !== 'number') return 'text';
      return isDateFormat(formatCodeOf(wb.styles, cell.s || 0)) ? 'date' : 'number';
    };
    const span = (a, b) => Array.from({ length: Math.min(b - a + 1, 500) }, (_, k) => a + k);
    const lines = (leftToRight, header) => (leftToRight
      ? span(area.r0, area.r1).map((r) => ({ at: r, label: `Row ${r + 1}`, kind: kindOf(span(area.c0, area.c1).map((c) => at(r, c))) }))
      : span(area.c0, area.c1).map((c) => {
        const top = at(area.r0, c)?.v;
        const label = header && top != null && top !== '' ? String(top) : `Column ${colName(c)}`;
        return { at: c, label, kind: kindOf(span(area.r0 + (header ? 1 : 0), area.r1).map((r) => at(r, c))) };
      }));
    const start = { keys: [{ at: Math.min(Math.max(sel.ac, area.c0), area.c1), desc: false }], header: guessHeader(wb, si, area), matchCase: false, leftToRight: false };
    const got = await sortBox(root, { lines, start });
    if (!got) { state.grid.el.focus({ preventScroll: true }); return false; }
    return editing().run('sort', { area, ...got });
  };
  /** A typed value the cell's data validation refuses: Excel's box for the rule's style (Stop: Retry or Cancel;
   *  Warning: Continue? Yes, No or Cancel; Information: OK or Cancel), its title first. 'retry', 'force' or 'cancel'. */
  const askInvalid = async ({ style, title, message }) => {
    const text = title ? `${title}

${message}` : message;
    if (style === 'warning') { const b = await messageBox(root, `${text}

Continue?`, ['Yes', 'No', 'Cancel']); return b === 'Yes' ? 'force' : b === 'No' ? 'retry' : 'cancel'; }
    if (style === 'information') return (await messageBox(root, text, ['OK', 'Cancel'])) === 'OK' ? 'force' : 'cancel';
    return (await messageBox(root, text, ['Retry', 'Cancel'])) === 'Retry' ? 'retry' : 'cancel';
  };
  /** The active cell's data validation shown: a list's dropdown button beside it, the rule's input message under it. */
  const showValidation = () => {
    if (!state.grid || !state.workbook || !state.editing) return '';
    const s = editing().selection;
    const rule = ruleAt(state.workbook.sheet(state.sheetIndex), s.ar, s.ac);
    const prompt = rule && rule.showPrompt && (rule.promptTitle || rule.prompt) ? { title: rule.promptTitle, text: rule.prompt } : null;
    state.grid.setValidation(rule ? { r: s.ar, c: s.ac, list: rule.type === 'list' && rule.dropDown && !state.readOnly, prompt } : null);
    return prompt ? `, ${[prompt.title, prompt.text].filter(Boolean).join(': ')}` : '';
  };
  /** A list's choices for the active cell (Alt+Down, or its button): picked, it is entered as if typed. */
  const listMenu = (rect = null) => {
    const s = editing().selection;
    const wb = state.workbook;
    const rule = ruleAt(wb.sheet(state.sheetIndex), s.ar, s.ac);
    if (!rule || rule.type !== 'list' || !rule.dropDown || state.readOnly) return false;
    const { items } = listItems(wb, state.sheetIndex, rule, s.ar, s.ac);
    const at = rect || (state.grid.el.querySelector('.oxe-dv-btn') || state.grid.el).getBoundingClientRect();
    pickBox(root, { title: 'Choose a value', at: { x: at.left, y: at.bottom + 2 }, items: items.map((x) => x.text) }).then((i) => {
      if (i != null) editing().run('enter', items[i].text);
      state.grid.el.focus({ preventScroll: true });
    });
    return true;
  };
  /** The shown sheet's conditional formats worked out again (after a change, or another sheet shown). */
  const showConditional = () => {
    if (!state.grid || !state.workbook) return;
    state.grid.setConditional(conditionalLooks(state.workbook, state.sheetIndex));
    // (Notes too: written, deleted, or moved with rows and columns.)
    state.grid.setNotes(readNotes(state.workbook, state.sheetIndex));
    // (The sheet's charts, read again with their data; and, for a screen reader, each as a figure with its data.)
    const charts = readCharts(state.workbook, state.sheetIndex).filter((c) => c.kind);
    state.grid.setCharts(charts, state.workbook.theme);
    let figures = root.querySelector(':scope > .oxe-chart-figures');
    if (!figures && charts.length) { figures = document.createElement('section'); figures.className = 'oxe-chart-figures'; figures.setAttribute('aria-label', 'Charts on this sheet'); root.append(figures); }
    figures?.replaceChildren(...charts.map(chartFigure));
  };
  /** A filter button's list for column c (dialogs.js filterBox), opened under `rect`: its sort, Clear Filter From, or
   *  the values ticked. */
  const openFilter = async (c, rect) => {
    const wb = state.workbook;
    const si = state.sheetIndex;
    const f = wb.sheet(si).autoFilter;
    if (!f || !allowed('filterBy')) return false;
    const title = shownText(wb, wb.sheet(si).cells.get(f.r0)?.get(c)) || `Column ${colName(c)}`;
    const choices = filterChoices(wb, si, f, c);
    const numbers = choices.values.length > 0 && choices.values.every((v) => v.text !== '' && Number.isFinite(Number(v.text.replace(/,/g, ''))));
    const got = await filterBox(root, { title, at: { x: rect.left, y: rect.bottom + 2 }, choices, numbers, filtered: f.columns.has(c) });
    if (got && !state.readOnly) {
      if (got.do === 'asc' || got.do === 'desc') editing().run('filterSort', c, got.do === 'desc');
      else if (got.do === 'clear') editing().run('filterBy', c, null);
      else editing().run('filterBy', c, got.all ? null : { values: got.values, blank: got.blank });
    }
    state.grid.el.focus({ preventScroll: true });
    return true;
  };
  /** Alt+Down on a heading cell of the filter: its list, as in Excel (false elsewhere, for the key to move on). */
  const filterMenu = () => {
    const f = state.workbook.sheet(state.sheetIndex).autoFilter;
    const s = editing().selection;
    if (!f || s.ar !== f.r0 || s.ac < f.c0 || s.ac > f.c1) return false;
    const btn = state.grid.el.querySelector(`.oxe-filters .oxe-filter-btn[data-col="${s.ac}"]`);
    openFilter(s.ac, (btn || state.grid.el).getBoundingClientRect());
    return true;
  };
  /** The menu of a cell (right-click): Excel's, as far as the editor goes. */
  const cellMenu = (point) => {
    const ro = state.readOnly;
    const run = (name, ...a) => () => act(name, ...a);
    menu.open([{ label: 'Cut', run: run('clipCut'), disabled: ro || !allowed('clipCut') }, { label: 'Copy', run: run('clipCopy'), disabled: !allowed('clipCopy') },
      { label: 'Paste', run: run('clipPaste'), disabled: ro || !allowed('clipPaste') }, { label: 'Paste Special...', run: run('pasteSpecialBox'), disabled: ro || !allowed('pasteSpecialBox') },
      'separator', { label: 'Insert...', run: () => lines(1), disabled: ro || !allowed('lines') }, { label: 'Delete...', run: () => lines(-1), disabled: ro || !allowed('lines') },
      { label: 'Clear Contents', run: () => editing().run('clear'), disabled: ro }, 'separator',
      // (Excel's: New Note on a cell without one; Edit, Delete and Show/Hide Note on one with.)
      ...(() => {
        const s = editing().selection;
        const has = readNotes(state.workbook, state.sheetIndex).has(`${s.ar},${s.ac}`);
        return has ? [{ label: 'Edit Note', run: run('editNote'), disabled: ro }, { label: 'Delete Note', run: run('deleteNote'), disabled: ro }, { label: 'Show/Hide Note', run: run('showNote') }]
          : [{ label: 'New Note', run: run('editNote'), disabled: ro }];
      })()],
    point, { label: 'Cell' });
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
    showStats();
  };
  /** The status bar: Excel's Average, Count and Sum of the selection (stats.js), or Count alone. */
  const showStats = () => {
    if (!state.tabs || !state.editing) return;
    if (options.statusBar === false) { state.tabs.status('', true); return; }
    const s = selectionStats(state.workbook, state.sheetIndex, state.editing.commands.selection);
    const parts = !s ? [] : s.text ? [`Average: ${s.text.average}`, `Count: ${s.count}`, `Sum: ${s.text.sum}`] : [`Count: ${s.count}`];
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
        Data: (kit) => dataTab(kit, { act }),
        Review: (kit) => reviewTab(kit, { act }),
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
        state.grid.onDropdown((button) => listMenu(button.getBoundingClientRect()));
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
          selectionShown: () => showValidation(),
          goToBox: () => act('goToBox'),
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
      return new AgentWorkbook(() => state.workbook, {
        cells: (si, contents, styles) => {
          if (!editing().run('applyChanges', [{ si, contents, styles }])) throw new Error(state.readOnly ? 'The workbook is read-only' : "You can't change part of an array");
        },
        structure: (run, kind) => editing().run('agentChange', run, kind),
      });
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
