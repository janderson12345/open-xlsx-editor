// Every action is a named command, shared by the keyboard, the mouse, the formula bar and the API: moving and
// extending the selection, entering and clearing values and formulas, undo and redo, changing sheet, and adding,
// renaming, moving, hiding, colouring and deleting sheets. Cell changes go through the workbook (only the sheet's XML
// is touched, as everywhere; the formulas that read a changed cell are worked out again) and into an undo history of
// cell contents, one for the workbook, as in Excel; the selection, kept for each sheet, is view state and never reaches
// the history or the file. Sheets are known by their sheetId in both, as indexes change when sheets move.
import { cellSelection, extendSelection } from './selection.js';
import { step, jump, lastCell } from './navigate.js';
import { parseInput } from '../ooxml/input.js';
import { refName } from '../ooxml/refs.js';
import { SheetError, nearest } from '../ooxml/sheetops.js';
import * as fmt from '../ooxml/formatops.js';
import { mergeChange } from '../ooxml/mergeops.js';
import * as clip from '../ooxml/clipops.js';

/** The commands for one editor. `ctx` = { workbook(), sheetIndex(), showSheet(i), sheetsChanged(), grid, readOnly(),
 *  announce(msg), alert(msg) (a refusal Excel shows in a message box), changed(), selected(sel, note), restyled() (the
 *  workbook has new cell formats), shownText(r, c) (a cell's text as the grid shows it) }. */
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
  const idOf = (i) => ctx.workbook().sheets[i]?.sheetId;
  const indexOf = (id) => ctx.workbook().sheets.findIndex((s) => s.sheetId === id);

  /** Apply changes [{ r, c, value } or { r, c, formula }] as one step of the history; refused, whole, if a cell is part
   *  of an array formula. */
  function apply(changes, record = true) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    if (changes.some((ch) => inArray(ch.r, ch.c))) { ctx.announce("You can't change part of an array"); return false; }
    const before = changes.map((ch) => contents(ch.r, ch.c));
    const { circular } = ctx.workbook().apply(ctx.sheetIndex(), changes);
    if (record) { state.undo.push({ sheet: idOf(ctx.sheetIndex()), before, after: changes }); state.redo = []; }
    ctx.grid.refresh();
    ctx.changed?.();
    if (circular.length) ctx.announce(`Circular reference: ${circular.map((e) => refName(e.r, e.c)).join(', ')} read themselves and weren't worked out`);
    return true;
  }

  /** Give cells new formats [{ r, c, s }] as one step of the history (formatops.js works them out). */
  function restyle(cells, record = true) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    const sh = sheet();
    const now = cells.filter(({ r, c, s }) => (sh.cells.get(r)?.get(c)?.s || 0) !== s);
    if (!now.length) return false;
    const before = now.map(({ r, c }) => ({ r, c, s: sh.cells.get(r)?.get(c)?.s || 0 }));
    ctx.workbook().setStyles(ctx.sheetIndex(), now);
    if (record) { state.undo.push({ sheet: idOf(ctx.sheetIndex()), styles: true, before, after: now }); state.redo = []; }
    ctx.restyled?.();
    ctx.changed?.();
    return true;
  }
  /** What a cell of any sheet holds, as a change that would put it back. */
  const contentsOn = (sh, r, c) => {
    const cell = sh.cells.get(r)?.get(c);
    if (cell && cell.f != null && cell.f !== '') return { r, c, formula: cell.f };
    return { r, c, value: cell && cell.f == null ? cell.v ?? null : null };
  };
  /** Values, formats and merged ranges changed together, sheet by sheet ([{ sheet (its sheetId), contents, styles,
   *  merges }]: a merge, a paste, or their undo). Refused, whole, if a cell is part of an array formula. */
  function applySteps(list) {
    const wb = ctx.workbook();
    for (const st of list) {
      const sh = wb.sheet(indexOf(st.sheet));
      if (st.contents.some(({ r, c }) => sh.arrays.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1))) { ctx.announce("You can't change part of an array"); return false; }
    }
    const circular = [];
    for (const st of list) {
      const k = indexOf(st.sheet);
      if (st.contents.length) circular.push(...wb.apply(k, st.contents).circular);
      if (st.merges) wb.setMerges(k, st.merges);
      if (st.styles.length) wb.setStyles(k, st.styles);
    }
    ctx.restyled?.();
    ctx.changed?.();
    if (circular.length) ctx.announce(`Circular reference: ${circular.map((e) => refName(e.r, e.c)).join(', ')} read themselves and weren't worked out`);
    return true;
  }
  /** Changes of several sheets ([{ si, contents, styles, merges }]) made, as one step of the history. */
  function runSteps(raw) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    const wb = ctx.workbook();
    const list = raw.map((x) => ({ sheet: idOf(x.si), contents: x.contents, styles: x.styles, merges: x.merges }));
    const before = list.map((x) => {
      const sh = wb.sheet(indexOf(x.sheet));
      return { sheet: x.sheet, contents: x.contents.map(({ r, c }) => contentsOn(sh, r, c)), styles: x.styles.map(({ r, c }) => ({ r, c, s: sh.cells.get(r)?.get(c)?.s || 0 })), merges: sh.merges.map((m) => ({ ...m })) };
    });
    if (!applySteps(list)) return false;
    state.undo.push({ sheet: idOf(ctx.sheetIndex()), steps: true, before, after: list });
    state.redo = [];
    return true;
  }
  /** A format command: formatops.js's function over the selection on the sheet shown. */
  const formatting = (fn, ...args) => (ctx.readOnly() ? (ctx.announce('The workbook is read-only'), false) : restyle(fn(ctx.workbook(), ctx.sheetIndex(), state.sel, ...args)));

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
    /** Every selected cell with the same format change (stylesedit.js: font, fill, border, align, numFmt, style). */
    format(change) { return formatting(fmt.formatEach, change); },
    /** Bold (b), Italic (i), Underline (u) or Strikethrough (strike), following the active cell. */
    toggleFont(key) { return formatting(fmt.toggleFont, key); },
    /** Grow (1) or Shrink (-1) Font. */
    stepFont(dir) { return formatting(fmt.stepFont, dir); },
    toggleWrap() { return formatting(fmt.toggleWrap); },
    /** Increase (1) or Decrease (-1) Indent. */
    stepIndent(dir) { return formatting(fmt.stepIndent, dir); },
    /** Horizontal ('left', 'center', 'right') or vertical ('top', 'center', 'bottom') alignment. */
    align(key, value) { return formatting(fmt.setAlign, key, value); },
    /** A preset of the Borders menu (formatops.js BORDER_PRESETS). */
    borders(preset) { return formatting(fmt.setBorders, preset); },
    /** Increase (1) or Decrease (-1) Decimal. */
    stepDecimal(dir) { return formatting(fmt.stepDecimal, dir, ctx.shownText?.(state.sel.ar, state.sel.ac) ?? null); },
    /** Merge & Center ('center'), Merge Across ('across'), Merge Cells ('cells') or Unmerge Cells ('unmerge') on the
     *  selection (mergeops.js), as one step of the history. */
    merge(kind) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const change = mergeChange(ctx.workbook(), ctx.sheetIndex(), state.sel, kind);
      if (!change || !runSteps([{ si: ctx.sheetIndex(), ...change }])) return false;
      // (The selection becomes the merged area it is in, as the grid shows it.)
      show(state.sel, state.focus);
      return true;
    },
    /** Copy (Ctrl+C) or Cut (Ctrl+X) the selection: kept to be pasted, outlined until then (Esc lets it go). Returns
     *  the snapshot (clipops.js), whose `text` the caller sets to what went on the system clipboard. */
    copy(cut = false) {
      if (cut && ctx.readOnly()) { ctx.announce('The workbook is read-only'); return null; }
      state.clip = clip.snapshot(ctx.workbook(), ctx.sheetIndex(), state.sel, cut);
      ctx.clipShown?.(state.clip);
      return state.clip;
    },
    cut() { return commands.copy(true); },
    /** Let the copy or cut go (Esc): its outline goes. */
    dropClip() {
      if (!state.clip) return false;
      state.clip = null;
      ctx.clipShown?.(null);
      return true;
    },
    /** Paste (Ctrl+V) at the selection: the cells copied or cut here, or, when `text` from the system clipboard isn't
     *  what was copied here, that text read cell by cell as if typed. A paste of a cut moves the cells and lets the cut
     *  go. The pasted cells are selected. */
    paste(text = null) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const mine = state.clip && (text == null || text === state.clip.text || !text) && (!state.clip.cut || state.clip.wb === wb);
      const steps = mine ? (state.clip.cut ? clip.pasteCut : clip.pasteCopy)(wb, si, state.sel, state.clip) : text ? clip.pasteText(wb, si, state.sel, text) : null;
      if (!steps || !runSteps(steps)) return false;
      if (mine && state.clip.cut) commands.dropClip();
      const area = steps.find((x) => x.si === si)?.area;
      if (area) show({ ar: area.r0, ac: area.c0, ...area }, null, 'Pasted, ');
      return true;
    },
    /** Clear Formats: the selected cells back to the Normal style's format. */
    clearFormats() { return formatting((wb, si, sel) => fmt.cellsOf(wb.sheet(si), sel).map(({ r, c }) => ({ r, c, s: 0 }))); },
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
      if (from >= 0) state.views.set(idOf(from), { sel: state.sel, focus: state.focus, top: view.scrollTop, left: view.scrollLeft });
      ctx.showSheet(i);
      const back = state.views.get(idOf(i));
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
    /** A new sheet before the one shown (Insert, Shift+F11), or after it (`after`: the + button), shown. */
    insertSheet({ after = false } = {}) {
      let at = -1;
      return sheetOp((wb) => { at = wb.addSheet(ctx.sheetIndex() + (after ? 1 : 0)); }, () => at);
    },
    /** Rename sheet i (Excel's refusals said in its words). */
    renameSheet(i, name) { return sheetOp((wb) => wb.renameSheet(i, name), null, true); },
    /** Move sheet i to index `to`. */
    moveSheet(i, to) {
      if (to < 0 || to >= ctx.workbook().sheets.length) return false;
      return sheetOp((wb) => wb.moveSheet(i, to));
    },
    /** Hide sheet i's tab (showing the next sheet if it was the one shown), or show it again (`hidden` false). */
    hideSheet(i, hidden = true) {
      const id = idOf(i);
      return sheetOp((wb) => wb.setSheetHidden(i, hidden), hidden ? null : () => indexOf(id));
    },
    /** Sheet i's tab colour ({ rgb } or { theme, tint }; null for none). */
    tabColor(i, color) { return sheetOp((wb) => wb.setTabColor(i, color)); },
    /** Delete sheet i (no question asked here: the tab menu asks first, as Excel does). */
    deleteSheet(i) { return sheetOp((wb) => wb.deleteSheet(i), null, true); },
  };

  /** Run a change to the workbook's sheets. Afterwards the sheet shown is `show()` if given, else the same sheet
   *  wherever it went, else (gone or hidden) its nearest neighbour. A rename or delete clears the undo history (a
   *  change in it could name the sheet as it was). */
  function sheetOp(run, show = null, clearsHistory = false) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    const wb = ctx.workbook();
    const was = ctx.sheetIndex();
    const shownId = idOf(was);
    try {
      run(wb);
    } catch (e) {
      if (!(e instanceof SheetError)) throw e;
      ctx.alert(e.message);
      return false;
    }
    if (clearsHistory) { state.undo = []; state.redo = []; }
    for (const id of state.views.keys()) if (indexOf(id) < 0) state.views.delete(id);
    const now = indexOf(shownId);
    ctx.sheetsChanged(now); // (the tabs again, and the sheet shown at its new index, or -1 if it went)
    let target = show ? show() : now;
    if (target < 0 || !wb.canShow(target)) target = nearest(wb, target < 0 ? Math.min(was, wb.sheets.length - 1) : target);
    if (target !== now || now < 0) commands.showSheet(target);
    ctx.changed?.();
    return true;
  }

  function replay(from, to, side, word) {
    const entry = from.pop();
    if (!entry) return false;
    // (The history is the workbook's: undoing a change on another sheet shows that sheet first, as Excel does.)
    const at = indexOf(entry.sheet);
    if (at !== ctx.sheetIndex() && !commands.showSheet(at)) { from.push(entry); return false; }
    const done = entry.steps ? applySteps(entry[side]) : entry.styles ? restyle(entry[side], false) || true : apply(entry[side], false);
    if (!done) { from.push(entry); return false; }
    to.push(entry);
    const first = entry.steps ? { r: state.sel.ar, c: state.sel.ac } : entry[side][0];
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
    /** The active cell's look (formatops.js activeLook) for the ribbon. */
    look: () => fmt.activeLook(ctx.workbook(), ctx.sheetIndex(), state.sel),
    canUndo: () => state.undo.length > 0,
    canRedo: () => state.redo.length > 0,
    /** Start again (a new workbook): A1 selected, no history, no sheet's view kept. */
    reset() { state.undo = []; state.redo = []; state.views.clear(); commands.dropClip(); show(cellSelection(0, 0)); },
  };
}
