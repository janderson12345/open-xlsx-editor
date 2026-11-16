// The ribbon's actions and the boxes and menus they open (index.js shares the editor's state and helpers with them as
// one object): the act dispatcher, Insert and Delete, Paste Special, Sort and its warning, the validation question and
// message, a list's choices, a filter's list, the cell's and the headings' menus, Row Height and Column Width.
import { readTables } from '../ooxml/tables.js';
import { tableAt } from '../ooxml/tableedit.js';
import { messageBox, numberBox, optionsBox, goToBox, noteBox } from './dialogs.js';
import { protectionRefusal } from '../ooxml/protection.js';
import { hyperlinkAt } from '../ooxml/links.js';
import { formulaLink } from '../calc/hyperlinkcell.js';
import { createMoreActs } from './acts-more.js';
import { createSparklineActs } from './acts-sparklines.js';
import { createTableActs } from './acts-tables.js';
import { createOutlineActs, OUTLINE_ACTS } from './acts-outline.js';
import { borderDrawing, DRAW_ACTS } from './drawborders.js';
import { customFilterBox } from './filterbox.js';
import { sortBox, filterBox, pickBox, validationBox } from './datadialogs.js';
import { setNote, deleteNote } from '../ooxml/noteedit.js';
import { goToNames } from '../ooxml/goto.js';
import { specialBox } from './specialbox.js';
import { insertPivot, showFieldList } from './pivotbox.js';
import { pivotAt } from '../ooxml/pivot.js';
import { nameManager, defineNameFor, createNamesFor } from './namesbox.js';
import { areasOf } from './selection.js';
import { filterChoices, shownText } from '../ooxml/filterops.js';
import { readNotes } from '../ooxml/notes.js';
import { ruleAt, listItems } from '../ooxml/validation.js';
import { dataBeside, sortArea, guessHeader, SORT_MESSAGES } from '../ooxml/sortops.js';
import { isDateFormat } from '../ooxml/format.js';
import { shapeShift, ROWCOL_MESSAGES } from '../ooxml/rowcol.js';
import { boxRule } from '../ooxml/dvedit.js';
import { formatCodeOf } from '../ooxml/styles.js';
import { mergeLoss } from '../ooxml/mergeops.js';
import { stepSize } from '../ooxml/formatops.js';
import { MAX_ROWS, MAX_COLS, colName } from '../ooxml/refs.js';

/** The Paste menu's choices, as Paste Special options. */
const PASTE_AS = { formulas: { paste: 'formulas' }, values: { paste: 'values' }, formats: { paste: 'formats' }, transpose: { paste: 'all', transpose: true }, link: { link: true } };
/** Excel's Paste Special box: its Paste and Operation choices, by label, as clipops.js pasteSpecial's options. */
const PASTE_CHOICES = { All: 'all', Formulas: 'formulas', Values: 'values', Formats: 'formats', 'Comments and Notes': 'comments', Validation: 'validation',
  'All using Source theme': 'theme', 'All except borders': 'noBorders', 'Column widths': 'widths', 'Formulas and number formats': 'formulasFormats', 'Values and number formats': 'valuesFormats',
  'All merging conditional formats': 'merging' };
const OPERATIONS = { None: null, Add: 'add', Subtract: 'subtract', Multiply: 'multiply', Divide: 'divide' };

/** The actions: `env` is mountXlsxEditor's shared state and helpers (index.js). Returns each by name. */
/** Whether (r, c) of sheet si is a cell of a table's total row. */
const totalCellAt = (wb, si, r, c) => { const t = tableAt(wb, si, r, c); return !!t && !!t.totals && r === t.r1; };

export function createActs(env) {
  const { root, state, options, editing, menu, sheetActs, allowed, followLink } = env;
  /** Format Painter: the selection's formats copied (outlined as a copy is), then given to the next cells selected by
   *  Paste Special > Formats (clipops.js, as measured for paste-excel.json), once; pressed twice quickly, kept for
   *  every selection until it is pressed again or Esc lets the copy go. */
  const painter = {
    now: null,
    active() { if (this.now && state.clip !== this.now.clip) this.stop(false); return !!this.now; },
    stop(drop = true) {
      const was = this.now;
      this.now = null;
      root.classList.remove('oxe-painting');
      if (drop && was && state.clip === was.clip) editing().run('dropClip');
      state.ribbon?.painting?.(false);
    },
    toggle() {
      if (this.active()) {
        // (A second press soon after the first, nothing painted yet: kept on, as Excel's double click keeps it.)
        if (!this.now.lock && !this.now.used && Date.now() - this.now.at < 500) { this.now.lock = true; return true; }
        this.stop();
        return true;
      }
      if (!state.editing.pickFormats()) return false;
      this.now = { clip: state.clip, lock: false, used: false, at: Date.now() };
      root.classList.add('oxe-painting');
      state.ribbon?.painting?.(true);
      return true;
    },
    paint() {
      if (!this.active()) return false;
      this.now.used = true;
      const done = editing().run('pasteSpecial', { paste: 'formats' });
      if (!this.now.lock) this.stop();
      return done;
    },
  };
  const drawing = borderDrawing(env);
  let more = null;
  let sparks = null;
  let tables = null;
  let outline = null;
  const moreAct = (name, args, i) => (more ??= createMoreActs({ ...env, act }))(name, args, i);
  /** Bold, Italic, Underline, Strikethrough, the font, its size, Grow and Shrink Font and the font colour on the part
   *  of a cell's text selected while it is edited (richmarks.js), as Excel's Font group does then; false when no text
   *  is selected (the cells are formatted). Automatic is theme colour 1, as a cell's own Automatic is written. */
  const partOfText = (name, args) => {
    const ed = state.editing;
    if (!ed?.textSelected) return false;
    if (name === 'toggleFont') return ed.markEditor(args[0]);
    if (name === 'format' && args[0]?.font && Object.keys(args[0]).length === 1) {
      const f = { ...args[0].font };
      if ('color' in f && !f.color) f.color = { theme: 1 };
      return ed.markEditorFont(f);
    }
    if (name === 'stepFont') {
      const s = ed.commands.selection;
      const cell = state.workbook.sheet(state.sheetIndex).cells.get(s.ar)?.get(s.ac);
      const { styles } = state.workbook;
      const base = styles.fonts[(styles.cellXfs[cell?.s || 0] || styles.cellXfs[0]).fontId]?.size ?? 11;
      return ed.markEditorFont((size) => ({ size: stepSize(size, args[0]) }), base);
    }
    return false;
  };
  /** The ribbon's commands: the editor's own, and the sheet actions that ask first or edit a tab in place. */
  const act = (name, ...args) => {
    if (name === 'painting') return painter.active();
    if (DRAW_ACTS.has(name)) return drawing.act(name, ...args);
    if (!state.workbook) return false;
    // (The Font group while part of a cell's text is selected in the cell editor or the formula bar: that part.)
    if (partOfText(name, args)) return true;
    if (!allowed(name, args[0])) return false;
    const i = state.sheetIndex;
    // (Insert > Sparklines and the Sparkline tab: acts-sparklines.js.)
    if (name.startsWith('sparkline')) return (sparks ??= createSparklineActs({ ...env, act }))(name, args);
    // (Insert > Table, Format as Table and the Table Design tab: acts-tables.js.)
    if (name.startsWith('table') || name === 'formatAsTable') return (tables ??= createTableActs({ ...env, act }))(name, args);
    // (Data > Outline's boxes: acts-outline.js.)
    if (OUTLINE_ACTS.has(name)) return (outline ??= createOutlineActs({ ...env, act }))(name, args, i);
    // (Charts, comments, Format Cells and conditional formatting: acts-more.js.)
    const more = moreAct(name, args, i);
    if (more !== undefined) return more;
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
    if (name === 'defaultWidthBox') {
      // (Excel's Standard Width box, holding the sheet's default width in characters.)
      if (state.readOnly) return false;
      const value = Math.max(0, state.grid.geometry().cols.base - 5) / 7;
      return numberBox(root, { title: 'Standard Width', label: 'Standard column width:', value: String(Math.round(value * 100) / 100), max: 255, refusal: 'Column width must be between 0 and 255.' })
        .then((n) => (n == null ? false : editing().run('defaultWidth', n)));
    }
    if (name === 'clipCopy' || name === 'clipCut') return state.editing.copyOut(name === 'clipCut');
    if (name === 'clipPaste') return state.editing.pasteIn();
    if (name === 'formatPainter') return painter.toggle();
    if (name === 'paintFormats') return painter.paint();
    if (name === 'pasteSpecialBox') return pasteSpecialBox();
    if (name === 'sortAsc' || name === 'sortDesc') return sortQuick(name === 'sortDesc');
    if (name === 'sortBox') return customSort();
    if (name === 'editNote') {
      // (Excel's New Note and Edit Note: the active cell's note in a box to type in, beside the cell.)
      const s = editing().selection;
      const n = readNotes(state.workbook, i).get(`${s.ar},${s.ac}`);
      if (n?.kind === 'thread') return act('commentBox');
      // (A new note on a protected sheet without Edit objects: refused before the box, as Excel's measured.)
      const barred = !n && protectionRefusal(state.workbook, i, 'objects');
      if (barred) return messageBox(root, barred).then(() => false);
      const body = n ? n.runs.filter((x, k) => !(k === 0 && x.bold)).map((x) => x.text).join('').replace(/^\n/, '') : '';
      const author = n ? n.author || (n.runs[0]?.bold ? n.runs[0].text.replace(/:$/, '') : '') : options.author || 'Author';
      const box = (root.querySelector('.oxe-sel-active') || state.grid.el).getBoundingClientRect();
      const ref = `${colName(s.ac)}${s.ar + 1}`;
      return noteBox(root, { title: n ? `Edit Note for ${ref}` : `New Note for ${ref}`, author, text: body, at: { x: box.right + 8, y: box.top } }).then((text) => {
        // (A new note on a protected sheet: Edit objects allowed; one there already changes, as Excel's did.)
        if (text != null && (n || text !== '')) editing().run('sheetChange', (wb, si) => setNote(wb, si, s.ar, s.ac, text, options.author || 'Author'), n ? null : 'objects');
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'deleteNote') {
      const s = editing().selection;
      return editing().run('sheetChange', (wb, si) => deleteNote(wb, si, s.ar, s.ac));
    }
    if (name === 'showNote') { const s = editing().selection; return state.grid.toggleNote(s.ar, s.ac); }
    if (name === 'showAllNotes') return state.grid.showAllNotes();
    if (name === 'validationBox') {
      // (Excel's Data Validation box, holding the active cell's rule; its answer made for the selection.)
      const s = editing().selection;
      return validationBox(root, boxRule(ruleAt(state.workbook.sheet(i), s.ar, s.ac), state.workbook.date1904)).then((rule) => {
        if (rule) editing().run('sheetChange', (wb, si) => wb.setValidation(si, areasOf(s), rule === 'clear' ? null : rule), 'always');
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'goToBox') {
      // (Excel's Go To: a name, a table or a reference; gone to as the Name Box goes.)
      return goToBox(root, goToNames(state.workbook, i)).then((text) => {
        if (text === goToBox.SPECIAL) return act('specialBox');
        if (text) editing().run('nameBox', text);
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'nameManager' || name === 'defineName' || name === 'createNames') {
      // (Formulas > Defined Names: each change one step of the history, Excel's words when it refuses.)
      const run = (fn) => {
        let refused = null;
        editing().run('sheetChange', (wb) => { try { fn(wb); } catch (e) { refused = e.message; return false; } return true; });
        return refused;
      };
      const s = editing().selection;
      const newName = () => defineNameFor(root, state.workbook, i, s, run);
      const opened = name === 'nameManager' ? nameManager(root, { workbook: () => state.workbook, run, newName }) : name === 'createNames' ? createNamesFor(root, state.workbook, i, s, run) : newName();
      return opened.then((x) => { state.grid.el.focus({ preventScroll: true }); return x; });
    }
    if (name === 'pivotFieldList') {
      // (Show Field List: the PivotTable at the active cell's fields, moved.)
      const s = editing().selection;
      return showFieldList(root, state.workbook, i, { r: s.ar, c: s.ac }, (...a) => editing().run(...a)).then((done) => { state.grid.el.focus({ preventScroll: true }); return done; });
    }
    if (name === 'pivotBox') {
      // (Insert > PivotTable: its boxes, then the PivotTable made.)
      const s = editing().selection;
      return insertPivot(root, state.workbook, i, { r: s.ar, c: s.ac }, (...a) => editing().run(...a)).then((done) => { state.grid.el.focus({ preventScroll: true }); return done; });
    }
    if (name === 'specialBox') {
      // (Excel's Go To Special box; its choice found and selected.)
      return specialBox(root).then((choice) => {
        if (choice) editing().run('goToSpecial', choice);
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
    // (The cell and font colours a column's cells have, for Sort On Cell Color and Font Color.)
    const coloursOf = (cells) => {
      const out = { cellColor: new Set(), fontColor: new Set() };
      for (const cell of cells) {
        const xf = wb.styles.cellXfs[cell?.s || 0];
        if (!xf) continue;
        const fill = wb.styles.fills[xf.fillId];
        if (fill?.pattern === 'solid' && fill.fg?.rgb) out.cellColor.add(fill.fg.rgb.toUpperCase());
        const ink = wb.styles.fonts[xf.fontId]?.color?.rgb;
        if (ink) out.fontColor.add(ink.toUpperCase());
      }
      return { cellColor: [...out.cellColor], fontColor: [...out.fontColor] };
    };
    const lines = (leftToRight, header) => (leftToRight
      ? span(area.r0, area.r1).map((r) => ({ at: r, label: `Row ${r + 1}`, kind: kindOf(span(area.c0, area.c1).map((c) => at(r, c))), colors: coloursOf(span(area.c0, area.c1).map((c) => at(r, c))) }))
      : span(area.c0, area.c1).map((c) => {
        const top = at(area.r0, c)?.v;
        const label = header && top != null && top !== '' ? String(top) : `Column ${colName(c)}`;
        const cells = span(area.r0 + (header ? 1 : 0), area.r1).map((r) => at(r, c));
        return { at: c, label, kind: kindOf(cells), colors: coloursOf(cells) };
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
    // (A table's total cell: its drop-down of functions, acts-tables.js tableTotalMenu.)
    const total = !rule && !state.readOnly && totalCellAt(state.workbook, state.sheetIndex, s.ar, s.ac);
    state.grid.setValidation(rule ? { r: s.ar, c: s.ac, list: rule.type === 'list' && rule.dropDown && !state.readOnly, prompt } : total ? { r: s.ar, c: s.ac, list: true } : null);
    return prompt ? `, ${[prompt.title, prompt.text].filter(Boolean).join(': ')}` : '';
  };
  /** A list's choices for the active cell (Alt+Down, or its button): picked, it is entered as if typed. */
  const listMenu = (rect = null) => {
    const s = editing().selection;
    const wb = state.workbook;
    const rule = ruleAt(wb.sheet(state.sheetIndex), s.ar, s.ac);
    if (!rule && totalCellAt(wb, state.sheetIndex, s.ar, s.ac)) return act('tableTotalMenu', rect);
    if (!rule || rule.type !== 'list' || !rule.dropDown || state.readOnly) return false;
    const { items } = listItems(wb, state.sheetIndex, rule, s.ar, s.ac);
    const at = rect || (state.grid.el.querySelector('.oxe-dv-btn') || state.grid.el).getBoundingClientRect();
    pickBox(root, { title: 'Choose a value', at: { x: at.left, y: at.bottom + 2 }, items: items.map((x) => x.text) }).then((i) => {
      if (i != null) editing().run('enter', items[i].text);
      state.grid.el.focus({ preventScroll: true });
    });
    return true;
  };
  /** A filter button's list for column c (dialogs.js filterBox), opened under `rect`: its sort, Clear Filter From, or
   *  the values ticked. */
  const openFilter = async (c, rect, table = null) => {
    const wb = state.workbook;
    const si = state.sheetIndex;
    // (A table's own button: its filter, kept in its part, tablefilter.js.)
    const f = table ? readTables(wb).find((x) => x.name === table)?.filter : wb.sheet(si).autoFilter;
    if (!f || !allowed('filterBy')) return false;
    const title = shownText(wb, wb.sheet(si).cells.get(f.r0)?.get(c)) || `Column ${colName(c)}`;
    const choices = filterChoices(wb, si, f, c);
    const numbers = choices.values.length > 0 && choices.values.every((v) => v.text !== '' && Number.isFinite(Number(v.text.replace(/,/g, ''))));
    // (A column of dates: Date Filters; its fills and font colours offered to filter by.)
    const sheet = wb.sheet(si);
    const dates = [];
    const fills = new Set();
    const fonts = new Set();
    for (let r = f.r0 + 1; r <= f.r1; r++) {
      const cell = sheet.cells.get(r)?.get(c);
      if (!cell) continue;
      if (typeof cell.v === 'number' && isDateFormat(formatCodeOf(wb.styles, cell.s || 0))) dates.push(r);
      const xf = wb.styles.cellXfs[cell.s || 0] || wb.styles.cellXfs[0];
      const fill = wb.styles.fills[xf.fillId];
      if (fill?.pattern === 'solid' && state.looks.color(fill.fg)) fills.add(state.looks.color(fill.fg).toUpperCase());
      const fc = wb.styles.fonts[xf.fontId]?.color;
      if (fc && !fc.auto && !(fc.theme === 1 && !fc.tint) && state.looks.color(fc)) fonts.add(state.looks.color(fc).toUpperCase());
    }
    const kind = dates.length && dates.length === choices.values.length ? 'Date' : numbers ? 'Number' : 'Text';
    const got = await filterBox(root, { title, at: { x: rect.left, y: rect.bottom + 2 }, choices, numbers, filtered: f.columns.has(c), kind });
    if (got?.do === 'custom') {
      const col = await customFilterBox(root, { title, kind: kind.toLowerCase(), colors: { fill: [...fills], font: [...fonts] } });
      if (col && !state.readOnly) editing().run('filterBy', c, col, table);
      state.grid.el.focus({ preventScroll: true });
      return true;
    }
    if (got && !state.readOnly) {
      if (got.do === 'asc' || got.do === 'desc') editing().run('filterSort', c, got.do === 'desc', table);
      else if (got.do === 'clear') editing().run('filterBy', c, null, table);
      else editing().run('filterBy', c, got.all ? null : { values: got.values, blank: got.blank, ...(got.dates ? { dates: got.dates } : {}) }, table);
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
  /** A chart's menu (right click, Shift+F10): Edit Chart..., Select Data..., Format Chart..., Delete. */
  const chartMenu = (point, k) => {
    const ro = state.readOnly;
    menu.open([{ label: 'Edit Chart...', run: () => act('chartBox', k), disabled: ro }, { label: 'Select Data...', run: () => act('chartDataBox', k), disabled: ro }, { label: 'Format Chart...', run: () => act('chartFormatBox', k), disabled: ro }, 'separator', { label: 'Delete', run: () => act('chartDelete', k), disabled: ro }], point,
      { back: () => state.grid.el.focus({ preventScroll: true }) });
  };
  const cellMenu = (point) => {
    const ro = state.readOnly;
    const run = (name, ...a) => () => act(name, ...a);
    menu.open([{ label: 'Cut', run: run('clipCut'), disabled: ro || !allowed('clipCut') }, { label: 'Copy', run: run('clipCopy'), disabled: !allowed('clipCopy') },
      { label: 'Paste', run: run('clipPaste'), disabled: ro || !allowed('clipPaste') }, { label: 'Paste Special...', run: run('pasteSpecialBox'), disabled: ro || !allowed('pasteSpecialBox') },
      'separator', { label: 'Insert...', run: () => lines(1), disabled: ro || !allowed('lines') }, { label: 'Delete...', run: () => lines(-1), disabled: ro || !allowed('lines') },
      { label: 'Clear Contents', run: () => editing().run('clear'), disabled: ro }, 'separator',
      { label: 'Format Cells...', run: run('formatCellsBox'), disabled: ro || !allowed('formatCellsBox') }, 'separator',
      // (On a PivotTable's cell, Excel's Refresh and Show Field List.)
      ...(() => {
        const s = editing().selection;
        return pivotAt(state.workbook, state.sheetIndex, s.ar, s.ac) ? [{ label: 'Refresh', run: () => editing().run('refreshPivots', false), disabled: ro }, { label: 'Show Field List', run: run('pivotFieldList'), disabled: ro }, 'separator'] : [];
      })(),
      // (A hyperlink's cell: Excel's Edit Hyperlink..., Open Hyperlink and Remove Hyperlink; another, Link.)
      ...(() => {
        const s = editing().selection;
        const h = hyperlinkAt(state.workbook, state.sheetIndex, s.ar, s.ac);
        // (A HYPERLINK formula's cell: Open Hyperlink; it has no hyperlink of its own to edit or remove.)
        const f = !h && formulaLink(state.workbook, state.sheetIndex, s.ar, s.ac);
        if (f) return [{ label: 'Open Hyperlink', run: () => followLink(f) }, 'separator'];
        return h ? [{ label: 'Edit Hyperlink...', run: run('linkBox'), disabled: ro }, { label: 'Open Hyperlink', run: () => followLink(h) }, { label: 'Remove Hyperlink', run: () => editing().run('clearAs', 'removeHyperlinks'), disabled: ro }, 'separator']
          : [{ label: 'Link', run: run('linkBox'), disabled: ro }, 'separator'];
      })(),
      // (Excel's: New Comment and New Note on a cell with neither; a thread's Reply to Comment and Delete Comment; a
      // note's Edit, Delete and Show/Hide Note.)
      ...(() => {
        const s = editing().selection;
        const n = readNotes(state.workbook, state.sheetIndex).get(`${s.ar},${s.ac}`);
        if (n?.kind === 'thread') return [{ label: 'Reply to Comment', run: run('commentBox'), disabled: ro }, { label: 'Delete Comment', run: run('deleteNote'), disabled: ro }, { label: 'Show/Hide Comment', run: run('showNote') }];
        return n ? [{ label: 'Edit Note', run: run('editNote'), disabled: ro }, { label: 'Delete Note', run: run('deleteNote'), disabled: ro }, { label: 'Show/Hide Note', run: run('showNote') }]
          : [{ label: 'New Comment', run: run('commentBox'), disabled: ro }, { label: 'New Note', run: run('editNote'), disabled: ro }];
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
  return { act, lines, pasteSpecialBox, sortWarning, sortQuick, customSort, askInvalid, showValidation, listMenu, openFilter, filterMenu, cellMenu, chartMenu, headerMenu, sizeBox };
}
