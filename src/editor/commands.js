// Every action is a named command, shared by the keyboard, the mouse, the formula bar and the API: moving and
// extending the selection, entering and clearing values and formulas, undo and redo, changing sheet, and adding,
// renaming, moving, hiding, colouring and deleting sheets. Cell changes go through the workbook (only the sheet's XML
// is touched, as everywhere; the formulas that read a changed cell are worked out again) and into an undo history of
// cell contents, one for the workbook, as in Excel; the selection, kept for each sheet, is view state and never reaches
// the history or the file. Sheets are known by their sheetId in both, as indexes change when sheets move.
import { cellSelection, extendSelection } from './selection.js';
import { readPivots, refreshPivot, pivotAt } from '../ooxml/pivot.js';
import { step, jump, lastCell } from './navigate.js';
import { entryChanges } from '../ooxml/entry.js';
import { refName, colName, MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { SheetError, nearest } from '../ooxml/sheetops.js';
import * as fmt from '../ooxml/formatops.js';
import { mergeChange } from '../ooxml/mergeops.js';
import * as clip from '../ooxml/clipops.js';
import { typedWidth, pxWidth } from '../ooxml/sizes.js';
import { shiftCells } from '../ooxml/rowcol.js';
import { fillCopy, autoFill } from '../ooxml/fillops.js';
import { clearChanges } from '../ooxml/clearops.js';
import { quickSort, sortChanges } from '../ooxml/sortops.js';
import { autoSum as autoSumOf } from '../ooxml/autosumops.js';
import { goToTarget } from '../ooxml/goto.js';
import { filterArea } from '../ooxml/filterops.js';
import { ruleAt, valid, VALIDATION_MESSAGE } from '../ooxml/validation.js';

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
    ctx.edited?.({ cells: [{ si: ctx.sheetIndex(), contents: changes }] });
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
    ctx.edited?.({ book: true });
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
    const mergesOf = (st) => JSON.stringify(wb.sheet(indexOf(st.sheet)).merges);
    const mergedBefore = list.map(mergesOf);
    for (const st of list) {
      const k = indexOf(st.sheet);
      if (st.contents.length) circular.push(...wb.apply(k, st.contents).circular);
      if (st.merges) wb.setMerges(k, st.merges);
      if (st.styles.length) wb.setStyles(k, st.styles);
    }
    // (For a shared workbook, collab.js: contents alone go cell by cell, formats and merges with the whole workbook.)
    const formats = list.some((st, k) => st.styles.length || mergesOf(st) !== mergedBefore[k]);
    ctx.edited?.(formats ? { book: true } : { cells: list.map((st) => ({ si: indexOf(st.sheet), contents: st.contents })) });
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
  /** The selection's rows or columns, [first, last] (past the used range cut to it, as formatting is), or just line
   *  `index` when it isn't one of the selected whole lines. */
  function linesOf(axis, index = null) {
    const s = state.sel;
    const sh = sheet();
    const rows = axis === 'row';
    const whole = rows ? s.c0 === 0 && s.c1 >= MAX_COLS - 1 : s.r0 === 0 && s.r1 >= MAX_ROWS - 1;
    const [a, b] = rows ? [s.r0, s.r1] : [s.c0, s.c1];
    if (index != null && !(whole && index >= a && index <= b)) return [index, index];
    const last = rows ? (b - a > 10000 ? Math.max(a, Math.min(b, sh.maxRow)) : b) : (b - a > 1000 ? Math.max(a, Math.min(b, sh.maxCol)) : b);
    return [a, last];
  }
  /** Sizes changed by `apply(wb, si)`, as one step of the history (a snapshot of the workbook before and after). */
  function sizeStep(apply, note) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    const wb = ctx.workbook();
    const before = wb.snapshot();
    apply(wb, ctx.sheetIndex());
    state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
    ctx.edited?.({ book: true });
    state.redo = [];
    ctx.reloaded?.();
    ctx.changed?.();
    ctx.announce(note);
    return true;
  }
  /** Rows, columns or cells moved by `change(wb, si)`, as one step of the history (a snapshot before and after).
   *  Excel's refusals said in its words (false); 'unmerge' for its question about merged cells. */
  function moved(change) {
    const wb = ctx.workbook();
    const before = wb.snapshot();
    try {
      change(wb, ctx.sheetIndex());
    } catch (e) {
      if (!(e instanceof SheetError)) throw e;
      if (e.ask) return 'unmerge';
      ctx.alert(e.message);
      return false;
    }
    state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
    ctx.edited?.({ book: true });
    state.redo = [];
    commands.dropClip();
    ctx.reloaded?.();
    ctx.changed?.();
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
    enter(text, at = null, { force = false } = {}) {
      const { r, c } = at || { r: state.sel.ar, c: state.sel.ac };
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      if (inArray(r, c)) { ctx.announce("You can't change part of an array"); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const ch = entryChanges(wb, si, r, c, text);
      if (ch.error) { ctx.announce(ch.error); return false; }
      // (A cell's data validation judges what is typed, a formula by its result (validation.js): refused, Excel's
      // question comes back as { invalid } for the caller to ask; `force` enters it anyway, as Yes and OK do.)
      const rule = !force && ruleAt(sheet(), r, c);
      if (rule && rule.showError) {
        const v = ch.content.formula != null ? wb.calc.formulaAt(si, r, c, ch.content.formula) : ch.content.value;
        if (!valid(wb, si, rule, r, c, v)) return { invalid: { style: rule.errorStyle, title: rule.errorTitle, message: rule.error || VALIDATION_MESSAGE } };
      }
      // (A date, a time, currency or a percentage typed into a General cell gives it Excel's format: one step with the value.)
      if (!ch.style) return apply([ch.content]);
      return runSteps([{ si: ctx.sheetIndex(), contents: [ch.content], styles: [ch.style] }]);
    },
    /** Set a cell to a value as given (a number, boolean, string or null), as the API does. */
    setValue(r, c, value) { return apply([{ r, c, value }]); },
    /** Delete and Clear Contents: empty the selected cells that have something in them (their formats stay). Only
     *  the cells there are are looked at, so a whole column is quick. Refused, as Excel refuses, over part of a merged
     *  range. */
    clear() { return commands.clearAs('contents'); },
    /** Home > Clear: Clear All ('all'), Clear Contents ('contents') or Clear Formats ('formats') of the selection
     *  (clearops.js), as one step of the history. */
    clearAs(what) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const step = clearChanges(ctx.workbook(), ctx.sheetIndex(), state.sel, what);
      if (step.error) { ctx.alert(step.error); return false; }
      if (what === 'contents' && !step.contents.length) return false;
      if (what === 'contents') return apply(step.contents);
      return runSteps([step]);
    },
    /** Whole rows ('row') or columns ('col') selected: line i, or from the anchor's line to i (`extend`). The active
     *  cell is the first one of the line, as in Excel. */
    selectLines(axis, i, extend = false) {
      const rows = axis === 'row';
      const anchor = extend ? (rows ? state.sel.ar : state.sel.ac) : i;
      const a = Math.min(anchor, i);
      const b = Math.max(anchor, i);
      const sel = rows ? { ar: anchor, ac: 0, r0: a, r1: b, c0: 0, c1: MAX_COLS - 1 } : { ar: 0, ac: anchor, r0: 0, r1: MAX_ROWS - 1, c0: a, c1: b };
      if (extend) { sel.ar = state.sel.ar; sel.ac = state.sel.ac; }
      const what = rows ? (a === b ? `Row ${a + 1}` : `Rows ${a + 1} to ${b + 1}`) : (a === b ? `Column ${colName(a)}` : `Columns ${colName(a)} to ${colName(b)}`);
      show(sel, { r: sel.ar, c: sel.ac }, `${what}, `);
    },
    /** Shift+Space (the selection's rows) and Ctrl+Space (its columns). */
    selectRows() { const s = state.sel; show({ ...s, c0: 0, c1: MAX_COLS - 1 }, state.focus); },
    selectColumns() { const s = state.sel; show({ ...s, r0: 0, r1: MAX_ROWS - 1 }, state.focus); },
    /** Ctrl+A: the whole sheet. */
    selectAll() { show({ ...state.sel, r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }, state.focus); },
    /** Insert (dir 1) or delete (-1) as many whole rows ('row') or columns ('col') as the selection spans, at its
     *  first one (rowcol.js), as one step of the history; Excel's refusals said in its words. */
    lines(axis, dir) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const s = state.sel;
      const rows = axis === 'row';
      const count = rows ? s.r1 - s.r0 + 1 : s.c1 - s.c0 + 1;
      if (count >= (rows ? MAX_ROWS : MAX_COLS)) return false;
      const done = moved((wb, si) => wb.changeLines(si, axis, rows ? s.r0 : s.c0, dir * count));
      if (done === true) show(state.sel, state.focus, `${dir > 0 ? 'Inserted' : 'Deleted'} ${count} ${rows ? 'row' : 'column'}${count > 1 ? 's' : ''}, `);
      return done;
    },
    /** Insert (dir 1) or delete (-1) the selected cells, the rest of their columns shifting down or up (axis 'row') or
     *  of their rows right or left ('col'), as one step of the history (rowcol.js shiftCells). Returns 'unmerge' when
     *  merged cells would be split: asked first, as Excel asks, and done again with `unmerge`. */
    cells(dir, axis, { unmerge = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const s = state.sel;
      const done = moved((wb, si) => shiftCells(wb, si, s, dir, axis, { unmerge }));
      if (done !== true) return done;
      // (Excel keeps the inserted cells selected; after a delete, the cells that took their place.)
      const way = dir > 0 ? (axis === 'row' ? 'shifted down' : 'shifted right') : (axis === 'row' ? 'shifted up' : 'shifted left');
      show(state.sel, state.focus, `${dir > 0 ? 'Inserted' : 'Deleted'} cells, the rest ${way}, `);
      return true;
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
    /** Freeze Panes ('panes': above and left of the active cell; again when frozen, Unfreeze Panes, as Excel's menu
     *  turns), Freeze Top Row ('topRow'), Freeze First Column ('firstColumn') or Unfreeze ('none'). A view change, kept
     *  in the file but not in the undo history, as in Excel. */
    freeze(kind) {
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const now = wb.sheet(si).freeze || { rows: 0, cols: 0 };
      const s = state.sel;
      let freeze = { rows: 0, cols: 0 };
      if (kind === 'topRow') freeze = { rows: 1, cols: 0 };
      else if (kind === 'firstColumn') freeze = { rows: 0, cols: 1 };
      else if (kind === 'panes' && !(now.rows || now.cols)) {
        // (Excel freezes the middle of the window when A1 is active; here that needs another cell.)
        if (!s.ar && !s.ac) { ctx.announce('Select the cell below and to the right of what should stay in place'); return false; }
        freeze = { rows: s.ar, cols: s.ac };
      }
      const topLeft = ctx.grid.view().topLeft;
      wb.setView(si, { freeze, active: { r: s.ar, c: s.ac }, topLeft });
      ctx.grid.setView({ freeze });
      ctx.changed?.();
      ctx.announce(freeze.rows || freeze.cols ? `Frozen: ${freeze.rows} row${freeze.rows === 1 ? '' : 's'} and ${freeze.cols} column${freeze.cols === 1 ? '' : 's'}` : 'Panes unfrozen');
      return true;
    },
    /** The sheet's zoom in percent (10 to 400, as Excel allows); a view change, not undone. */
    zoom(percent) {
      const z = Math.round(Math.min(400, Math.max(10, percent)));
      ctx.workbook().setView(ctx.sheetIndex(), { zoom: z });
      ctx.grid.setView({ zoom: z });
      ctx.grid.scrollIntoView(state.focus.r, state.focus.c);
      ctx.changed?.();
      ctx.announce(`Zoom ${z}%`);
      return true;
    },
    /** Zoom to Selection: the zoom at which the selection fills the window (cut to the used range when it runs past
     *  it), as near as whole percent allows. */
    zoomToSelection() {
      const s = state.sel;
      const sh = sheet();
      const r1 = s.r1 - s.r0 > 10000 ? Math.max(s.r0, Math.min(s.r1, sh.maxRow)) : s.r1;
      const c1 = s.c1 - s.c0 > 1000 ? Math.max(s.c0, Math.min(s.c1, sh.maxCol)) : s.c1;
      const g = ctx.grid.geometry();
      const w = g.cols.start(c1 + 1) - g.cols.start(s.c0);
      const h = g.rows.start(r1 + 1) - g.rows.start(s.r0);
      const view = ctx.grid.el;
      // (The headings scale with the sheet, so they count with it.)
      const z = Math.floor(Math.min(view.clientWidth / (Math.max(1, w) + g.rowHeadW), view.clientHeight / (Math.max(1, h) + g.colHeadH)) * 100);
      commands.zoom(z);
      ctx.grid.scrollToCell(s.r0, s.c0);
      return true;
    },
    /** Go to a found cell ({ si, r, c }): its sheet shown, the cell selected. */
    goTo(hit) {
      if (hit.si !== ctx.sheetIndex() && !commands.showSheet(hit.si)) return false;
      commands.select(hit.r, hit.c);
      return true;
    },
    /** Changes of several sheets ([{ si, contents, styles, merges }], findops.js replaceChanges) as one step. */
    applyChanges(steps) { return runSteps(steps); },
    /** Refresh (Alt+F5: the PivotTable at the active cell) or Refresh All (Ctrl+Alt+F5: every one), from their
     *  sources as Excel's Refresh does (pivot.js), as one step of the history; a pivot that can't be refreshed here
     *  leaves everything as it was, and says so. */
    refreshPivots(all = false) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const here = pivotAt(wb, ctx.sheetIndex(), state.sel.ar, state.sel.ac);
      const list = all ? readPivots(wb).filter((p, k, a) => a.findIndex((q) => q.cachePath === p.cachePath) === k) : here ? [here] : [];
      if (!list.length) { ctx.announce(all ? 'There are no PivotTables to refresh' : 'The active cell is not in a PivotTable'); return false; }
      const before = wb.snapshot();
      for (const p of list) {
        const out = refreshPivot(wb, p);
        if (out.error) {
          wb.restore(before);
          ctx.reloaded?.();
          ctx.alert(out.why ? `${out.error}: ${out.why}.` : out.error);
          return false;
        }
        for (const ch of out.changes) if (ch.contents.length) wb.apply(ch.si, ch.contents);
      }
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
      state.redo = [];
      ctx.edited?.({ book: true });
      ctx.reloaded?.();
      ctx.changed?.();
      ctx.announce(list.length === 1 ? 'PivotTable refreshed' : `${list.length} PivotTables refreshed`);
      return true;
    },
    /** A change to the sheet's notes or rules, `run(wb, si)` (noteedit.js setNote, deleteNote; Workbook
     *  setValidation), as one step of the history. */
    sheetChange(run) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const before = wb.snapshot();
      if (run(wb, ctx.sheetIndex()) === false) return false;
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
      state.redo = [];
      ctx.edited?.({ book: true });
      ctx.reloaded?.();
      ctx.changed?.();
      return true;
    },
    /** Cell contents [{ r, c, value } or { r, c, formula }] of sheet si that another editor of a shared workbook
     *  changed (collab.js): made here outside the history, and drawn. */
    remoteContents(si, contents) {
      ctx.workbook().apply(si, contents);
      if (si === ctx.sheetIndex()) ctx.grid.refresh();
      ctx.changed?.();
      return true;
    },
    /** A change by the agent (agent.js), `run(wb)`: of rows or columns ('lines': one step of the history) or of the
     *  sheets ('sheets', as the tab menu's). A refusal throws, in Excel's words, for the agent to read. */
    agentChange(run, kind) {
      if (ctx.readOnly()) throw new SheetError('The workbook is read-only');
      if (kind === 'sheets') return sheetOp(run, null, true, true);
      const wb = ctx.workbook();
      const before = wb.snapshot();
      run(wb);
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
    ctx.edited?.({ book: true });
      state.redo = [];
      commands.dropClip();
      ctx.reloaded?.();
      ctx.changed?.();
      return true;
    },
    /** A heading edge dragged to `px` pixels: line `index`, or every selected whole line when it is one of them (as
     *  in Excel). 0 hides. */
    resizeLines(axis, index, px) {
      const [a, b] = linesOf(axis, index);
      const size = Math.max(0, Math.round(px));
      const change = axis === 'col' ? { width: size ? pxWidth(size) : 0 } : { height: size ? size * 0.75 : 0 };
      return sizeStep((wb, si) => wb.setLines(si, axis, a, b, change), `${axis === 'col' ? 'Column width' : 'Row height'} changed`);
    },
    /** Column Width (characters) and Row Height (points) of the selection's columns or rows; 0 hides. */
    columnWidth(chars) {
      const [a, b] = linesOf('col');
      return sizeStep((wb, si) => wb.setLines(si, 'col', a, b, { width: chars > 0 ? typedWidth(chars) : 0 }), 'Column width changed');
    },
    rowHeight(points) {
      const [a, b] = linesOf('row');
      return sizeStep((wb, si) => wb.setLines(si, 'row', a, b, { height: points }), 'Row height changed');
    },
    /** AutoFit: columns as wide as their widest text (an empty one left as it is); rows back to the height their text
     *  needs. A heading edge's double click fits line `index` (or the selected lines it is one of). */
    fitLines(axis, index = null) {
      const [a, b] = linesOf(axis, index);
      if (axis === 'row') return sizeStep((wb, si) => wb.setLines(si, 'row', a, b, { height: null }), 'Row height fitted');
      const widths = [];
      for (let c = a; c <= b; c++) { const px = ctx.fitWidth?.(c); if (px) widths.push([c, px]); }
      if (!widths.length) return false;
      return sizeStep((wb, si) => { for (const [c, px] of widths) wb.setLines(si, 'col', c, c, { width: pxWidth(px), bestFit: true }); }, 'Column width fitted');
    },
    /** Hide (Ctrl+9 rows, Ctrl+0 columns) or show again (Ctrl+Shift+9, Ctrl+Shift+0) the selection's rows or columns. */
    hideLines(axis, hidden = true) {
      const [a, b] = linesOf(axis);
      return sizeStep((wb, si) => wb.setLines(si, axis, a, b, { hidden }), `${axis === 'col' ? 'Columns' : 'Rows'} ${hidden ? 'hidden' : 'shown'}`);
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
    /** Paste Special (Ctrl+Alt+V) of the cells copied here at the selection (clipops.js pasteSpecial: `paste`, `op`,
     *  `skip`, `transpose`; Column widths, `paste` 'widths'), or Paste Link (`link`), as one step of the history. The
     *  copy stays to be pasted again, as in Excel. Not for cells cut, nor for text from another program. */
    pasteSpecial({ paste = 'all', op = null, skip = false, transpose = false, link = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const copy = state.clip;
      if (!copy) { ctx.announce('Copy cells first: Paste Special pastes cells copied in this workbook or another one open here'); return false; }
      if (copy.cut) { ctx.announce("Paste Special isn't available for cells that were cut: paste them, or copy them instead"); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      if (paste === 'widths' && !link) {
        const widths = clip.pasteWidths(state.sel, copy);
        return sizeStep((book, k) => { for (const { c, width } of widths) book.setLines(k, 'col', c, c, { width }); }, 'Pasted column widths');
      }
      if (link && copy.wb !== wb) { ctx.announce("Paste Link to another workbook isn't available yet"); return false; }
      const steps = link ? clip.pasteLink(wb, si, state.sel, copy) : clip.pasteSpecial(wb, si, state.sel, copy, { paste, op, skip, transpose });
      if (!runSteps(steps)) return false;
      const area = steps[0].area;
      show({ ar: area.r0, ac: area.c0, ...area }, null, link ? 'Pasted links, ' : 'Pasted, ');
      return true;
    },
    /** Fill Down ('down', Ctrl+D), Right ('right', Ctrl+R), Up or Left over the selection (fillops.js fillCopy), as one
     *  step of the history. */
    fill(dir) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const steps = fillCopy(ctx.workbook(), ctx.sheetIndex(), state.sel, dir);
      if (!runSteps(steps)) return false;
      const a = steps[0].area;
      show({ ...state.sel, r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }, state.focus, `Filled ${dir}, `);
      return true;
    },
    /** The fill handle dragged from the selection out to `to` (fillops.js autoFill): Excel's guess, or with Ctrl held
     *  (`ctrl`) the other way, as Excel's Ctrl-drag: a single cell counts on, more cells are copied. The filled range
     *  is selected. */
    autoFill(to, { ctrl = false, type = null } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const s = state.sel;
      const one = s.r0 === s.r1 && s.c0 === s.c1;
      const how = type || (ctrl ? (one ? 'series' : 'copy') : 'default');
      const steps = autoFill(ctx.workbook(), ctx.sheetIndex(), s, to, how);
      if (!steps || !runSteps(steps)) return false;
      show({ ...s, r0: to.r0, c0: to.c0, r1: to.r1, c1: to.c1 }, null, `Filled ${refName(to.r0, to.c0)} to ${refName(to.r1, to.c1)}, `);
      return true;
    },
    /** Sort (sortops.js), as one step of the history; the selection stays. With `how` ({ area, keys, header,
     *  matchCase, leftToRight }, as the Sort box gives it) that sort; else A to Z ('asc') or Z to A ('desc') as
     *  Excel's buttons sort, the data beside a selection taken in too when `expand`. Excel's refusals in its words. */
    sort(how = 'asc', { expand = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const plan = typeof how === 'string' ? quickSort(wb, si, state.sel, how === 'desc', { expand }) : how;
      const step = plan.error ? plan : sortChanges(wb, si, plan.area, plan);
      if (step.error) { ctx.alert(step.error); return false; }
      if (step.contents.length && !runSteps([step])) return false;
      const order = typeof how === 'string' ? (how === 'desc' ? ', Z to A' : ', A to Z') : '';
      ctx.announce(`Sorted ${refName(plan.area.r0, plan.area.c0)} to ${refName(plan.area.r1, plan.area.c1)}${order}`);
      return true;
    },
    /** AutoSum (Alt+=) and its Average, Count Numbers, Max and Min (`fn`), as one step of the history, the totals
     *  selected with what they sum (autosumops.js). Where Excel leaves a formula being edited (one cell, or nothing
     *  found to sum) that cell is selected and { edit: text } returned, for the cell editor to open with it. */
    autoSum(fn = 'sum') {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const step = autoSumOf(ctx.workbook(), ctx.sheetIndex(), state.sel, fn);
      if (step.edit) { commands.select(step.edit.r, step.edit.c); return { edit: step.edit.text }; }
      if (!runSteps([step])) return false;
      show({ ...state.sel, ...step.area }, state.focus, 'AutoSum, ');
      return true;
    },
    /** The Name Box: go to the reference or name typed (goto.js), on its sheet; Excel's words when it can't. */
    nameBox(text) {
      const to = goToTarget(ctx.workbook(), ctx.sheetIndex(), text);
      if (to.error) { if (to.name) ctx.announce(to.error); else ctx.alert(to.error); return false; }
      if (to.si !== ctx.sheetIndex() && !commands.showSheet(to.si)) return false;
      show(to.sel, null, 'Went to ');
      return true;
    },
    /** Filter (Ctrl+Shift+L): the sheet's AutoFilter turned on over what Sort would take, or off (filterops.js). */
    filterToggle() {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const sh = sheet();
      if (sh.autoFilter) return moved((wb, si) => wb.setAutoFilter(si, null)) === true && (ctx.announce('Filter off'), true);
      const { area, error } = filterArea(sh, state.sel);
      if (error) { ctx.alert(error); return false; }
      return moved((wb, si) => wb.setAutoFilter(si, { ...area, columns: new Map() })) === true && (ctx.announce(`Filter on ${refName(area.r0, area.c0)} to ${refName(area.r1, area.c1)}`), true);
    },
    /** Column c of the filter shows rows with `show` ({ values: shown texts, blank }), or all of them (null). */
    filterBy(c, show) {
      const f = sheet().autoFilter;
      if (!f || ctx.readOnly()) return false;
      const columns = new Map(f.columns);
      if (show) columns.set(c, { values: show.values, blank: !!show.blank, other: false });
      else columns.delete(c);
      return moved((wb, si) => wb.setAutoFilter(si, { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1, columns })) === true;
    },
    /** Home > Sort & Filter > Clear (every column's filter taken off, the buttons kept), and Reapply (the rows
     *  filtered again, after edits). */
    filterClear() { const f = sheet().autoFilter; return !!f && f.columns.size > 0 && commands.filterAll(new Map()); },
    filterReapply() { const f = sheet().autoFilter; return !!f && commands.filterAll(f.columns); },
    filterAll(columns) {
      const f = sheet().autoFilter;
      if (!f || ctx.readOnly()) return false;
      return moved((wb, si) => wb.setAutoFilter(si, { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1, columns: new Map(columns) })) === true;
    },
    /** A filter button's Sort A to Z (or Z to A, `desc`): the filter's range by column c, its heading row kept. Not
     *  yet while rows are filtered out (Excel then sorts only the rows shown). */
    filterSort(c, desc) {
      const sh = sheet();
      const f = sh.autoFilter;
      if (!f) return false;
      for (let r = f.r0 + 1; r <= f.r1; r++) if (sh.rows.get(r)?.hidden) { ctx.announce("Sorting while rows are filtered out isn't available yet"); return false; }
      return commands.sort({ area: { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1 }, keys: [{ at: c, desc }], header: true });
    },
    /** Clear Formats: the selected cells back to the Normal style's format. */
    clearFormats() { return commands.clearAs('formats'); },
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
  function sheetOp(run, show = null, clearsHistory = false, quiet = false) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    const wb = ctx.workbook();
    const was = ctx.sheetIndex();
    const shownId = idOf(was);
    try {
      run(wb);
    } catch (e) {
      if (!(e instanceof SheetError) || quiet) throw e;
      ctx.alert(e.message);
      return false;
    }
    if (clearsHistory) { state.undo = []; state.redo = []; }
    ctx.edited?.({ book: true });
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
    if (entry.snap) {
      ctx.workbook().restore(entry[side]);
      ctx.edited?.({ book: true });
      ctx.reloaded?.();
      ctx.changed?.();
      to.push(entry);
      show(state.sel, state.focus, `${word}, `);
      return true;
    }
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
