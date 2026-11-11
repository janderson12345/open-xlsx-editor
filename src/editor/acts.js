// The ribbon's actions and the boxes and menus they open (index.js shares the editor's state and helpers with them as
// one object): the act dispatcher, Insert and Delete, Paste Special, Sort and its warning, the validation question and
// message, a list's choices, a filter's list, the cell's and the headings' menus, Row Height and Column Width.
import { messageBox, numberBox, optionsBox, goToBox, noteBox } from './dialogs.js';
import { sortBox, filterBox, pickBox, validationBox, cfRuleBox } from './datadialogs.js';
import { CF_PRESETS } from '../ooxml/cfedit.js';
import { setNote, deleteNote } from '../ooxml/noteedit.js';
import { goToNames } from '../ooxml/goto.js';
import { specialBox } from './specialbox.js';
import { insertPivot, showFieldList } from './pivotbox.js';
import { pivotAt } from '../ooxml/pivot.js';
import { nameManager, defineNameFor } from './namesbox.js';
import { formatCellsBox } from './formatbox.js';
import { threadBox } from './threadbox.js';
import { addThread, replyThread, editThread } from '../ooxml/threadedit.js';
import { insertChart } from '../ooxml/chartnew.js';
import { setChartTitle, setChartLegend, setChartLabels, setChartKind, setChartAnchor, deleteChart } from '../ooxml/chartedit.js';
import { chartBox } from './chartbox.js';
import { currentRegion } from '../ooxml/special.js';
import { xfOf, fontFamily } from '../ooxml/styles.js';
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
import { MAX_ROWS, MAX_COLS, colName } from '../ooxml/refs.js';

/** The Paste menu's choices, as Paste Special options. */
/** The Font tab's list: the theme's fonts first, then these. */
const FONT_NAMES = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];
const PASTE_AS = { formulas: { paste: 'formulas' }, values: { paste: 'values' }, formats: { paste: 'formats' }, transpose: { paste: 'all', transpose: true }, link: { link: true } };
/** Excel's Paste Special box: its Paste and Operation choices, by label, as clipops.js pasteSpecial's options. */
const PASTE_CHOICES = { All: 'all', Formulas: 'formulas', Values: 'values', Formats: 'formats', 'All except borders': 'noBorders',
  'Column widths': 'widths', 'Formulas and number formats': 'formulasFormats', 'Values and number formats': 'valuesFormats' };
const OPERATIONS = { None: null, Add: 'add', Subtract: 'subtract', Multiply: 'multiply', Divide: 'divide' };

/** The actions: `env` is mountXlsxEditor's shared state and helpers (index.js). Returns each by name. */
export function createActs(env) {
  const { root, state, options, editing, menu, sheetActs, allowed } = env;
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
  /** The ribbon's commands: the editor's own, and the sheet actions that ask first or edit a tab in place. */
  const act = (name, ...args) => {
    if (name === 'painting') return painter.active();
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
      const body = n ? n.runs.filter((x, k) => !(k === 0 && x.bold)).map((x) => x.text).join('').replace(/^\n/, '') : '';
      const author = n ? n.author || (n.runs[0]?.bold ? n.runs[0].text.replace(/:$/, '') : '') : options.author || 'Author';
      const box = (root.querySelector('.oxe-sel-active') || state.grid.el).getBoundingClientRect();
      const ref = `${colName(s.ac)}${s.ar + 1}`;
      return noteBox(root, { title: n ? `Edit Note for ${ref}` : `New Note for ${ref}`, author, text: body, at: { x: box.right + 8, y: box.top } }).then((text) => {
        if (text != null && (n || text !== '')) editing().run('sheetChange', (wb, si) => setNote(wb, si, s.ar, s.ac, text, options.author || 'Author'));
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'commentBox') {
      // (Excel's comment thread on the active cell: started, replied to, a comment edited or a reply deleted, the whole
      // thread deleted; each one step of Undo, the box opened again on the thread after each.)
      const s = editing().selection;
      const ref = `${colName(s.ac)}${s.ar + 1}`;
      const n = readNotes(state.workbook, i).get(`${s.ar},${s.ac}`);
      if (n?.kind === 'note') return messageBox(root, `${ref} has a note: a cell has a note or a comment, not both.`);
      const box = (root.querySelector('.oxe-sel-active') || state.grid.el).getBoundingClientRect();
      return threadBox(root, { ref, items: n?.items || [], at: { x: box.right + 8, y: box.top }, readOnly: state.readOnly }).then((answer) => {
        state.grid.el.focus({ preventScroll: true });
        if (!answer) return false;
        const who = options.author || 'Author';
        let refused = null;
        editing().run('sheetChange', (wb, si) => {
          try {
            if (answer.removeAll) return deleteNote(wb, si, s.ar, s.ac);
            if (answer.post != null) { if (n) replyThread(wb, si, s.ar, s.ac, answer.post, who); else addThread(wb, si, s.ar, s.ac, answer.post, who); }
            else if (answer.remove != null) editThread(wb, si, s.ar, s.ac, answer.remove, null);
            else if (answer.edit != null) editThread(wb, si, s.ar, s.ac, answer.edit, answer.text);
          } catch (e) { refused = e.message; return false; }
          return true;
        });
        if (refused) return messageBox(root, refused).then(() => false);
        return answer.removeAll ? true : act('commentBox');
      });
    }
    if (name === 'deleteNote') {
      const s = editing().selection;
      return editing().run('sheetChange', (wb, si) => deleteNote(wb, si, s.ar, s.ac));
    }
    if (name === 'showNote') { const s = editing().selection; return state.grid.toggleNote(s.ar, s.ac); }
    if (name === 'cfAdd') {
      // (Excel's Conditional Formatting menu: the rule's box, then the rule for the selection.)
      const s = editing().selection;
      const sel = [{ r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }];
      const add = (rule) => editing().run('sheetChange', (wb, si) => wb.setConditional(si, sel, rule));
      const presets = Object.entries(CF_PRESETS).map(([value, p]) => ({ value, label: p.label }));
      const kind = args[0];
      if (kind === 'scale3' || kind === 'scale2') return add({ kind: 'colorScale', points: kind === 'scale2' ? 2 : 3 });
      const WORDS = { greaterThan: 'GREATER THAN', lessThan: 'LESS THAN', between: 'BETWEEN', equal: 'EQUAL TO', containsText: 'CONTAIN THE TEXT' };
      const TITLES = { greaterThan: 'Greater Than', lessThan: 'Less Than', between: 'Between', equal: 'Equal To', containsText: 'Text That Contains', duplicate: 'Duplicate Values', top: 'Top 10 Items', topPercent: 'Top 10%', bottom: 'Bottom 10 Items', bottomPercent: 'Bottom 10%', above: 'Above Average', below: 'Below Average' };
      const ranked = ['top', 'topPercent', 'bottom', 'bottomPercent'].includes(kind);
      const text = WORDS[kind] ? `Format cells that are ${WORDS[kind]}:` : ranked ? `Format cells that rank in the ${kind.startsWith('top') ? 'TOP' : 'BOTTOM'}:` : kind === 'duplicate' ? '' : `Format cells that are ${kind === 'above' ? 'ABOVE' : 'BELOW'} AVERAGE for the selected range`;
      const inputs = kind === 'between' ? [{ label: 'From', value: '' }, { label: 'and', value: '' }] : WORDS[kind] ? [{ label: 'Value', value: '' }] : ranked ? [{ label: kind.endsWith('Percent') ? 'Percent' : 'Items', value: '10' }] : [];
      return cfRuleBox(root, { title: TITLES[kind], text, inputs, choice: kind === 'duplicate' ? [{ value: 'duplicate', label: 'Duplicate' }, { value: 'unique', label: 'Unique' }] : null, presets, preset: kind === 'below' ? 'yellow' : 'lightRed' }).then((a) => {
        if (a) {
          if (WORDS[kind] && kind !== 'containsText') add({ kind: 'cellIs', operator: kind, f1: a.values[0], f2: a.values[1], preset: a.preset });
          else if (kind === 'containsText') add({ kind: 'containsText', text: a.values[0], preset: a.preset });
          else if (kind === 'duplicate') add({ kind: 'duplicate', unique: a.choice === 'unique', preset: a.preset });
          else if (ranked) add({ kind: 'top10', rank: Math.max(1, Math.round(Number(a.values[0]) || 10)), percent: kind.endsWith('Percent'), bottom: kind.startsWith('bottom'), preset: a.preset });
          else add({ kind: 'average', below: kind === 'below', preset: a.preset });
        }
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'cfClear') {
      const s = editing().selection;
      return editing().run('sheetChange', (wb, si) => wb.setConditional(si, args[0] === 'sheet' ? null : [{ r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }], null));
    }
    if (name === 'validationBox') {
      // (Excel's Data Validation box, holding the active cell's rule; its answer made for the selection.)
      const s = editing().selection;
      return validationBox(root, boxRule(ruleAt(state.workbook.sheet(i), s.ar, s.ac), state.workbook.date1904)).then((rule) => {
        if (rule) editing().run('sheetChange', (wb, si) => wb.setValidation(si, areasOf(s), rule === 'clear' ? null : rule));
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
    if (name === 'nameManager' || name === 'defineName') {
      // (Formulas > Defined Names: each change one step of the history, Excel's words when it refuses.)
      const run = (fn) => {
        let refused = null;
        editing().run('sheetChange', (wb) => { try { fn(wb); } catch (e) { refused = e.message; return false; } return true; });
        return refused;
      };
      const s = editing().selection;
      const newName = () => defineNameFor(root, state.workbook, i, s, run);
      const opened = name === 'nameManager' ? nameManager(root, { workbook: () => state.workbook, run, newName }) : newName();
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
    if (name === 'insertChart') {
      // (Insert > Charts: the selection charted; one cell selected, the table round it, as Excel takes it.)
      const s = editing().selection;
      const sel = s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(state.workbook.sheet(i), s.r0, s.c0) : s;
      let made = null;
      const done = editing().run('sheetChange', (wb, si) => { made = insertChart(wb, si, sel, args[0]); return !!made; });
      if (!done && !made && !state.readOnly) return messageBox(root, 'Select the cells to chart first: a table with numbers in it.').then(() => false);
      return done;
    }
    if (name === 'chartAnchor') return editing().run('sheetChange', (wb, si) => { setChartAnchor(wb, si, args[0], args[1], args[2]); });
    if (name === 'chartDelete') return editing().run('sheetChange', (wb, si) => { deleteChart(wb, si, args[0]); });
    if (name === 'chartEdit') {
      // (Edit Chart's answer: each change as Excel writes it, chartedit.js; refused for a chart written otherwise.)
      const [k, change] = args;
      let refused = null;
      const done = editing().run('sheetChange', (wb, si) => {
        try {
          if (change.kind) setChartKind(wb, si, k, change.kind);
          if ('title' in change) setChartTitle(wb, si, k, change.title);
          if ('legend' in change) setChartLegend(wb, si, k, change.legend);
          if ('labels' in change) setChartLabels(wb, si, k, change.labels);
        } catch (e) { refused = e.message; return false; }
        return true;
      });
      if (refused) return messageBox(root, refused).then(() => false);
      state.picker?.pick(k, false);
      return done;
    }
    if (name === 'chartBox') {
      const chart = (state.charts || []).find((c) => c.index === args[0]);
      if (!chart) return false;
      return chartBox(root, chart).then((change) => {
        state.grid.el.focus({ preventScroll: true });
        return change ? act('chartEdit', args[0], change) : false;
      });
    }
    if (name === 'formatCellsBox') {
      // (Ctrl+1: the active cell's format in Excel's box; what was changed there given to the selection.)
      const wb = state.workbook;
      const s = editing().selection;
      const cell = wb.sheet(i).cells.get(s.ar)?.get(s.ac) || null;
      const xf = xfOf(wb.styles, cell?.s || 0);
      const font = wb.styles.fonts[xf.fontId];
      const look = { xf, font, fontName: fontFamily(font, wb.theme), fontNames: [...new Set([wb.theme.fonts.minor, wb.theme.fonts.major, ...FONT_NAMES])],
        fill: wb.styles.fills[xf.fillId], border: wb.styles.borders[xf.borderId], code: editing().look().code, cell, date1904: wb.date1904,
        many: areasOf(s).some((a) => a.r1 > a.r0 || a.c1 > a.c0), color: (c) => state.looks?.color(c) };
      return formatCellsBox(root, look, args[0]).then((change) => {
        if (change) editing().run('formatCells', change);
        state.grid.el.focus({ preventScroll: true });
        return !!change;
      });
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
  /** A chart's menu (right click, Shift+F10): Edit Chart..., Delete. */
  const chartMenu = (point, k) => {
    const ro = state.readOnly;
    menu.open([{ label: 'Edit Chart...', run: () => act('chartBox', k), disabled: ro }, 'separator', { label: 'Delete', run: () => act('chartDelete', k), disabled: ro }], point,
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
