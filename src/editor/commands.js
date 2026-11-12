// Every action is a named command, shared by the keyboard, the mouse, the formula bar and the API: moving and
// extending the selection, entering and clearing values and formulas, undo and redo, changing sheet, and adding,
// renaming, moving, hiding, colouring and deleting sheets. Cell changes go through the workbook (only the sheet's XML
// is touched, as everywhere; the formulas that read a changed cell are worked out again) and into an undo history of
// cell contents, one for the workbook, as in Excel; the selection, kept for each sheet, is view state and never reaches
// the history or the file. Sheets are known by their sheetId in both, as indexes change when sheets move.
import { cellSelection, extendSelection, addArea, areasOf } from './selection.js';
import { sheetCommands } from './commands-sheets.js';
import { dataCommands } from './commands-data.js';
import { partCommands } from './commands-parts.js';
import { pivotCommands } from './commands-pivot.js';
import { viewCommands } from './commands-view.js';
import { step, jump, lastCell } from './navigate.js';
import { entryChanges } from '../ooxml/entry.js';
import { refName, colName, MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { SheetError } from '../ooxml/sheetops.js';
import * as fmt from '../ooxml/formatops.js';
import { mergeChange } from '../ooxml/mergeops.js';
import { typedWidth, pxWidth } from '../ooxml/sizes.js';
import { shiftCells } from '../ooxml/rowcol.js';
import { clearChanges } from '../ooxml/clearops.js';
import { partsClear, PARTS_ONLY } from '../ooxml/clearparts.js';
import { traceLevel } from './tracing.js';
import { ruleAt, valid, VALIDATION_MESSAGE } from '../ooxml/validation.js';
import { protectionRefusal } from '../ooxml/protection.js';

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
  const contents = (r, c) => contentsOn(sheet(), r, c, ctx.sheetIndex());
  const inArray = (r, c) => sheet().arrays.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1);
  const idOf = (i) => ctx.workbook().sheets[i]?.sheetId;
  /** Sheet i's tracer arrows and how far each trace has gone ({ arrows, precedents, dependents }), by its sheetId. */
  state.traces = new Map();
  const traceOf = (i) => { const id = idOf(i); if (!state.traces.has(id)) state.traces.set(id, { arrows: [], precedents: null, dependents: null }); return state.traces.get(id); };
  const indexOf = (id) => ctx.workbook().sheets.findIndex((s) => s.sheetId === id);

  /** Whether a change of `kind` (protection.js: 'contents' with its `cells`, 'formatCells', 'insertRows'...; 'always'
   *  for one a protected sheet never allows) may be made to sheet si (the one shown); if not, Excel's refusal said. */
  function allowed(kind, cells = [], si = ctx.sheetIndex()) {
    const why = protectionRefusal(ctx.workbook(), si, kind === 'always' ? 'never' : kind, cells);
    if (!why) return true;
    (ctx.alert || ctx.announce)(why);
    return false;
  }

  /** Apply changes [{ r, c, value } or { r, c, formula }] as one step of the history; refused, whole, if a cell is part
   *  of an array formula or locked on a protected sheet. */
  function apply(changes, record = true) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    if (!allowed('contents', changes)) return false;
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
    if (!allowed('formatCells')) return false;
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
  /** What a cell holds, as a change that would put it back: a formula (of the old kind, `legacy`, when it doesn't
   *  spill), or a value; a cell a formula spilled into holds nothing of its own. */
  function contentsOn(sh, r, c, si = ctx.workbook().sheets.findIndex((s, i) => s.kind === 'worksheet' && s.path && ctx.workbook().sheet(i) === sh)) {
    const cell = sh.cells.get(r)?.get(c);
    if (cell && cell.f != null && cell.f !== '') return { r, c, formula: cell.f, ...(cell.dynamic || cell.array ? {} : { legacy: true }) };
    if (ctx.workbook().isSpilled(si, r, c)) return { r, c, value: null };
    return { r, c, value: cell && cell.f == null ? cell.v ?? null : null };
  }
  /** Values, formats and merged ranges changed together, sheet by sheet ([{ sheet (its sheetId), contents, styles,
   *  merges }]: a merge, a paste, or their undo). Refused, whole, if a cell is part of an array formula. */
  function applySteps(list) {
    const wb = ctx.workbook();
    for (const st of list) {
      const sh = wb.sheet(indexOf(st.sheet));
      if (st.contents.some(({ r, c }) => sh.arrays.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1))) { ctx.announce("You can't change part of an array"); return false; }
      if (!allowed('contents', st.contents, indexOf(st.sheet))) return false;
      if ((st.styles.length || st.merges) && !allowed('formatCells', [], indexOf(st.sheet))) return false;
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
  /** Sizes changed by `apply(wb, si)`, as one step of the history (a snapshot of the workbook before and after);
   *  `kind`, what a protected sheet must allow ('formatColumns' or 'formatRows'). */
  function sizeStep(apply, note, kind = 'formatColumns') {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    if (!allowed(kind)) return false;
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
  function moved(change, kind = 'always') {
    if (!allowed(kind)) return false;
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
  /** The selection's ranges, each as a selection with the active cell (formatting and Clear go over each, as Excel's
   *  do: borders outline each range, a font toggled goes by the active cell). */
  const eachArea = () => areasOf(state.sel).map((a) => ({ ...a, ar: state.sel.ar, ac: state.sel.ac }));
  /** A format command: formatops.js's function over each range of the selection on the sheet shown. */
  const formatting = (fn, ...args) => (ctx.readOnly() ? (ctx.announce('The workbook is read-only'), false) : restyle(eachArea().flatMap((a) => fn(ctx.workbook(), ctx.sheetIndex(), a, ...args))));

  const commands = {
    /** Select one cell. */
    select(r, c) { show(cellSelection(r, c)); },
    /** Ctrl+click: cell (r, c) a new range, the ranges selected before kept. */
    addArea(r, c) { show(addArea(state.sel, r, c), { r, c }, 'Added '); },
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
    /** Home > Clear: Clear All ('all'), Clear Contents ('contents'), Clear Formats ('formats'), Clear Comments and
     *  Notes ('comments'), Clear Hyperlinks ('hyperlinks') or Remove Hyperlinks ('removeHyperlinks') of the selection
     *  (clearops.js, clearparts.js), as one step of the history. */
    clearAs(what) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      // (Notes, comments and hyperlinks are parts of the sheet, not cells: a change of them goes back as a whole.)
      if (PARTS_ONLY.includes(what)) {
        const run = partsClear(wb, si, eachArea(), what);
        if (!run) { ctx.announce(what === 'comments' ? 'No notes or comments to clear' : 'No hyperlinks to clear'); return false; }
        return commands.sheetChange(run);
      }
      const steps = eachArea().map((a) => clearChanges(wb, si, a, what));
      const refused = steps.find((x) => x.error);
      if (refused) { ctx.alert(refused.error); return false; }
      const step = steps.length === 1 ? steps[0] : { ...steps[0], contents: steps.flatMap((x) => x.contents), styles: steps.flatMap((x) => x.styles), merges: steps[0].merges.filter((m) => steps.every((x) => x.merges.includes(m))) };
      const run = partsClear(wb, si, eachArea(), what, step);
      if (run) return commands.sheetChange(run);
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
      const done = moved((wb, si) => wb.changeLines(si, axis, rows ? s.r0 : s.c0, dir * count), `${dir > 0 ? 'insert' : 'delete'}${rows ? 'Rows' : 'Columns'}`);
      if (done === true) show(state.sel, state.focus, `${dir > 0 ? 'Inserted' : 'Deleted'} ${count} ${rows ? 'row' : 'column'}${count > 1 ? 's' : ''}, `);
      return done;
    },
    /** Insert (dir 1) or delete (-1) the selected cells, the rest of their columns shifting down or up (axis 'row') or
     *  of their rows right or left ('col'), as one step of the history (rowcol.js shiftCells). Returns 'unmerge' when
     *  merged cells would be split: asked first, as Excel asks, and done again with `unmerge`. */
    cells(dir, axis, { unmerge = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const s = state.sel;
      const done = moved((wb, si) => shiftCells(wb, si, s, dir, axis, { unmerge }), `${dir > 0 ? 'insert' : 'delete'}${axis === 'row' ? 'Rows' : 'Columns'}`);
      if (done !== true) return done;
      // (Excel keeps the inserted cells selected; after a delete, the cells that took their place.)
      const way = dir > 0 ? (axis === 'row' ? 'shifted down' : 'shifted right') : (axis === 'row' ? 'shifted up' : 'shifted left');
      show(state.sel, state.focus, `${dir > 0 ? 'Inserted' : 'Deleted'} cells, the rest ${way}, `);
      return true;
    },
    /** Every selected cell with the same format change (stylesedit.js: font, fill, border, align, numFmt, style). */
    format(change) { return formatting(fmt.formatEach, change); },
    /** The Format Cells box's answer (formatops.js formatCells: borders by outline and inside lines). */
    formatCells(change) { return formatting(fmt.formatCells, change); },
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
    /** Formulas > Trace Precedents ('precedents') or Trace Dependents ('dependents') from the active cell: one level
     *  more each time, as Excel's arrows (ooxml/trace.js, arrows.js). Not a step of the history; not saved. */
    trace(kind) {
      const added = traceLevel(ctx.workbook(), ctx.sheetIndex(), traceOf(ctx.sheetIndex()), kind, { r: state.sel.ar, c: state.sel.ac });
      ctx.grid.setArrows(traceOf(ctx.sheetIndex()).arrows);
      if (!added) ctx.announce(kind === 'precedents' ? 'No more precedents to trace' : 'No more dependents to trace');
      return added > 0;
    },
    /** Remove Arrows (kind null), Remove Precedent Arrows or Remove Dependent Arrows, on the sheet shown. */
    removeArrows(kind = null) {
      const t = traceOf(ctx.sheetIndex());
      if (kind) { t.arrows = t.arrows.filter((a) => a.kind !== kind); t[kind] = null; } else Object.assign(t, { arrows: [], precedents: null, dependents: null });
      ctx.grid.setArrows(t.arrows);
      return true;
    },
    /** A heading edge dragged to `px` pixels: line `index`, or every selected whole line when it is one of them (as
     *  in Excel). 0 hides. */
    resizeLines(axis, index, px) {
      const [a, b] = linesOf(axis, index);
      const size = Math.max(0, Math.round(px));
      const change = axis === 'col' ? { width: size ? pxWidth(size) : 0 } : { height: size ? size * 0.75 : 0 };
      return sizeStep((wb, si) => wb.setLines(si, axis, a, b, change), `${axis === 'col' ? 'Column width' : 'Row height'} changed`, axis === 'col' ? 'formatColumns' : 'formatRows');
    },
    /** Column Width (characters) and Row Height (points) of the selection's columns or rows; 0 hides. */
    columnWidth(chars) {
      const [a, b] = linesOf('col');
      return sizeStep((wb, si) => wb.setLines(si, 'col', a, b, { width: chars > 0 ? typedWidth(chars) : 0 }), 'Column width changed');
    },
    rowHeight(points) {
      const [a, b] = linesOf('row');
      return sizeStep((wb, si) => wb.setLines(si, 'row', a, b, { height: points }), 'Row height changed', 'formatRows');
    },
    /** AutoFit: columns as wide as their widest text (an empty one left as it is); rows back to the height their text
     *  needs. A heading edge's double click fits line `index` (or the selected lines it is one of). */
    fitLines(axis, index = null) {
      const [a, b] = linesOf(axis, index);
      if (axis === 'row') return sizeStep((wb, si) => wb.setLines(si, 'row', a, b, { height: null }), 'Row height fitted', 'formatRows');
      const widths = [];
      for (let c = a; c <= b; c++) { const px = ctx.fitWidth?.(c); if (px) widths.push([c, px]); }
      if (!widths.length) return false;
      return sizeStep((wb, si) => { for (const [c, px] of widths) wb.setLines(si, 'col', c, c, { width: pxWidth(px), bestFit: true }); }, 'Column width fitted');
    },
    /** Hide (Ctrl+9 rows, Ctrl+0 columns) or show again (Ctrl+Shift+9, Ctrl+Shift+0) the selection's rows or columns. */
    hideLines(axis, hidden = true) {
      const [a, b] = linesOf(axis);
      return sizeStep((wb, si) => wb.setLines(si, axis, a, b, { hidden }), `${axis === 'col' ? 'Columns' : 'Rows'} ${hidden ? 'hidden' : 'shown'}`, axis === 'col' ? 'formatColumns' : 'formatRows');
    },
    /** Clear Formats: the selected cells back to the Normal style's format. */
    clearFormats() { return commands.clearAs('formats'); },
    undo() { return replay(state.undo, state.redo, 'before', 'Undone'); },
    redo() { return replay(state.redo, state.undo, 'after', 'Redone'); },
  };


  function replay(from, to, side, word) {
    const entry = from.pop();
    if (!entry) return false;
    // (The history is the workbook's: undoing a change on another sheet shows that sheet first, as Excel does.)
    const at = indexOf(entry.sheet);
    if (entry.sheets) {
      // (Sheets added, renamed, moved, hidden or deleted: the workbook back with its sheets, the tabs again, and the
      // sheet that was shown then.)
      const shownId = idOf(ctx.sheetIndex());
      const was = ctx.sheetIndex();
      ctx.workbook().restore(entry[side]);
      ctx.edited?.({ book: true });
      env.sheetsMoved(ctx.workbook(), shownId, was, () => indexOf(entry.shown[side]));
      ctx.reloaded?.();
      ctx.changed?.();
      to.push(entry);
      ctx.announce(`${word}: the sheets as they were`);
      return true;
    }
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

  // (The other commands, from modules sharing this state and these helpers: sheets first, whose sheetOp the parts
  // use.)
  const env = { allowed, commands, ctx, state, sheet, limits, show, contents, inArray, idOf, indexOf, apply, restyle, contentsOn, applySteps, runSteps, linesOf, sizeStep, moved, formatting };
  Object.assign(commands, sheetCommands(env));
  Object.assign(commands, dataCommands(env), partCommands(env), pivotCommands(env), viewCommands(env));

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
