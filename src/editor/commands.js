// Every action is a named command, shared by the keyboard, the mouse, the formula bar and the API: moving and
// extending the selection, entering and clearing values and formulas, undo and redo, changing sheet. Changes go
// through the workbook (only the sheet's XML is touched, as everywhere; the formulas that read a changed cell are
// worked out again) and into an undo history of cell contents, one for the workbook, as in Excel; the selection, kept
// for each sheet, is view state and never reaches the history or the file.
import { cellSelection, extendSelection } from './selection.js';
import { step, jump, lastCell } from './navigate.js';
import { parseInput } from '../ooxml/input.js';
import { refName } from '../ooxml/refs.js';

/** The commands for one editor. `ctx` = { workbook(), sheetIndex(), showSheet(i), grid, readOnly(), announce(msg),
 *  changed(), selected(sel, note) }. */
export function createCommands(ctx) {
  // (`views`: each sheet's selection and scroll position while another is shown.)
  const state = { sel: cellSelection(0, 0), focus: { r: 0, c: 0 }, undo: [], redo: [], views: new Map() };
  const sheet = () => ctx.workbook().sheet(ctx.sheetIndex());
  const limits = () => {
    const g = ctx.grid.geometry();
    return { rows: g.rows.count - 1, cols: g.cols.count - 1 };
  };
  const show = (sel, focus, note = '') => {
    state.sel = sel;
    state.focus = focus || { r: sel.ar, c: sel.ac };
    ctx.grid.setSelection(sel);
    ctx.grid.scrollIntoView(state.focus.r, state.focus.c);
    ctx.selected?.(sel, note);
  };
  /** What a cell holds, as a change that would put it back: { r, c, formula } or { r, c, value }. */
  const contents = (r, c) => {
    const cell = sheet().cells.get(r)?.get(c);
    if (cell && cell.f != null && cell.f !== '') return { r, c, formula: cell.f };
    return { r, c, value: cell && cell.f == null ? cell.v ?? null : null };
  };
  const inArray = (r, c) => sheet().arrays.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1);

  /** Apply changes [{ r, c, value } or { r, c, formula }] as one step of the history; refused, whole, if a cell is part
   *  of an array formula. */
  function apply(changes, record = true) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    if (changes.some((ch) => inArray(ch.r, ch.c))) { ctx.announce("You can't change part of an array"); return false; }
    const before = changes.map((ch) => contents(ch.r, ch.c));
    const { circular } = ctx.workbook().apply(ctx.sheetIndex(), changes);
    if (record) { state.undo.push({ sheet: ctx.sheetIndex(), before, after: changes }); state.redo = []; }
    ctx.grid.refresh();
    ctx.changed?.();
    if (circular.length) ctx.announce(`Circular reference: ${circular.map((e) => refName(e.r, e.c)).join(', ')} read themselves and weren't worked out`);
    return true;
  }

  const commands = {
    /** Select one cell. */
    select(r, c) { show(cellSelection(r, c)); },
    /** Extend the selection from the active cell to (r, c). */
    extendTo(r, c) { show(extendSelection(state.sel, r, c), { r, c }); },
    /** Arrow keys: a cell (or, with `jump`, to the edge of the data) in direction (dr, dc); `extend` moves the far
     *  corner of the selection instead of the active cell. */
    move(dr, dc, { extend = false, jump: far = false } = {}) {
      const from = extend ? state.focus : { r: state.sel.ar, c: state.sel.ac };
      const to = (far ? jump : step)(sheet(), from.r, from.c, dr, dc, limits());
      if (extend) commands.extendTo(to.r, to.c);
      else commands.select(to.r, to.c);
    },
    /** Page Up and Page Down: a screenful of rows. */
    page(dir, { extend = false } = {}) { commands.move(dir * ctx.grid.rowsInView(), 0, { extend }); },
    /** Home (the row's first cell), Ctrl+Home (A1), Ctrl+End (the last used cell). */
    home({ sheetStart = false } = {}) { commands.select(sheetStart ? 0 : state.sel.ar, 0); },
    end() { const { r, c } = lastCell(sheet()); commands.select(r, c); },
    /** Enter what was typed into the active cell (or `at`): parsed as Excel reads an entry (input.js), a formula if
     *  it starts with =. */
    enter(text, at = null) {
      const { r, c } = at || { r: state.sel.ar, c: state.sel.ac };
      const parsed = parseInput(text);
      if (parsed.error) { ctx.announce(parsed.error); return false; }
      return apply([parsed.formula != null ? { r, c, formula: parsed.formula } : { r, c, value: parsed.value }]);
    },
    /** Set a cell to a value as given (a number, boolean, string or null), as the API does. */
    setValue(r, c, value) { return apply([{ r, c, value }]); },
    /** Delete: empty the selected cells that have something in them (their formats stay). */
    clear() {
      const s = state.sel;
      const changes = [];
      for (let r = s.r0; r <= s.r1; r++) {
        for (let c = s.c0; c <= s.c1; c++) {
          const cell = sheet().cells.get(r)?.get(c);
          if (cell && (cell.v != null || cell.f != null)) changes.push({ r, c, value: null });
        }
      }
      return changes.length ? apply(changes) : false;
    },
    undo() { return replay(state.undo, state.redo, 'before', 'Undone'); },
    redo() { return replay(state.redo, state.undo, 'after', 'Redone'); },
    /** Show sheet i, with the selection and scroll position it had (A1 the first time); false if it can't be shown
     *  (hidden, or a chart sheet). */
    showSheet(i) {
      const wb = ctx.workbook();
      if (!wb.canShow(i)) {
        if (wb.sheets[i]?.kind === 'chartsheet') ctx.announce(`${wb.sheets[i].name} is a chart sheet, which can't be shown yet`);
        return false;
      }
      const from = ctx.sheetIndex();
      if (i === from) return true;
      const view = ctx.grid.el;
      state.views.set(from, { sel: state.sel, focus: state.focus, top: view.scrollTop, left: view.scrollLeft });
      ctx.showSheet(i);
      const back = state.views.get(i);
      if (back) { view.scrollTop = back.top; view.scrollLeft = back.left; }
      show(back ? back.sel : cellSelection(0, 0), back?.focus, `${wb.sheets[i].name}, `);
      return true;
    },
    /** Ctrl+Page Down and Ctrl+Page Up: the next or previous sheet that can be shown (hidden ones and chart sheets
     *  passed over, no wrapping round, as in Excel). */
    nextSheet(dir) {
      const wb = ctx.workbook();
      for (let i = ctx.sheetIndex() + dir; i >= 0 && i < wb.sheets.length; i += dir) if (wb.canShow(i)) return commands.showSheet(i);
      return false;
    },
  };

  function replay(from, to, side, word) {
    const entry = from.pop();
    if (!entry) return false;
    // (The history is the workbook's: undoing a change on another sheet shows that sheet first, as Excel does.)
    if (entry.sheet !== ctx.sheetIndex() && !commands.showSheet(entry.sheet)) { from.push(entry); return false; }
    if (!apply(entry[side], false)) { from.push(entry); return false; }
    to.push(entry);
    const first = entry[side][0];
    commands.select(first.r, first.c);
    ctx.announce(`${word}: ${refName(first.r, first.c)}`);
    return true;
  }

  return {
    commands,
    /** Run a command by name. */
    run(name, ...args) {
      if (!commands[name]) throw new Error(`no command ${name}`);
      return commands[name](...args);
    },
    get selection() { return { ...state.sel }; },
    canUndo: () => state.undo.length > 0,
    canRedo: () => state.redo.length > 0,
    /** Start again (a new workbook): A1 selected, no history, no sheet's view kept. */
    reset() { state.undo = []; state.redo = []; state.views.clear(); show(cellSelection(0, 0)); },
  };
}
