// Outlines as commands (commands.js shares its state and helpers as `env`): Group and Ungroup, the outline bar's level
// and +/- buttons, Show Detail and Hide Detail, Clear Outline, Auto Outline, the Settings, Subtotal and its Remove All;
// each one step of the history (a snapshot of the workbook, as Excel's Undo takes each back at once).
import { MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { groupLines, ungroupLines, showLevels, hideDetail, showDetail, clearOutline, autoOutline, setSummary, setOutlineSymbols, outlineOf } from '../ooxml/outline.js';
import { subtotal, removeSubtotals } from '../ooxml/subtotal.js';

export function outlineCommands(env) {
  const { ctx, state, commands, linesOf, moved } = env;
  const kind = (axis) => (axis === 'col' ? 'formatColumns' : 'formatRows');
  /** `run(wb, si)` as one step, `note` said when it changed something, `refused` when it didn't. */
  const step = (run, note, axis = 'row', refused = null) => {
    const done = commands.sheetChange(run, kind(axis));
    if (done) ctx.announce(note); else if (refused && !ctx.readOnly()) ctx.announce(refused);
    return done;
  };
  /** The axis the selection's Show Detail, Hide Detail and Clear Outline go by: whole columns selected, columns; whole
   *  rows, rows; else the rows when the active cell's row is in a row group (or sums one up), the columns otherwise. */
  const axisOf = () => {
    const s = state.sel;
    if (s.r0 === 0 && s.r1 >= MAX_ROWS - 1 && !(s.c0 === 0 && s.c1 >= MAX_COLS - 1)) return 'col';
    if (s.c0 === 0 && s.c1 >= MAX_COLS - 1) return 'row';
    const o = outlineOf(ctx.workbook().sheet(ctx.sheetIndex()));
    const near = (groups, i) => groups.some((g) => (i >= g.start && i <= g.end) || g.summary === i);
    return near(o.rows.groups, s.ar) || !near(o.cols.groups, s.ac) ? 'row' : 'col';
  };
  const lineOf = (axis) => (axis === 'row' ? state.sel.ar : state.sel.ac);
  return {
    /** Data > Group (Shift+Alt+Right): the selection's rows or columns one level deeper. */
    group(axis) {
      const [a, b] = linesOf(axis);
      return step((wb, si) => groupLines(wb, si, axis, a, b), `${axis === 'col' ? 'Columns' : 'Rows'} grouped`, axis, 'No more than 7 levels can be grouped');
    },
    /** Data > Ungroup (Shift+Alt+Left): the selection's rows or columns one level out. */
    ungroup(axis) {
      const [a, b] = linesOf(axis);
      return step((wb, si) => ungroupLines(wb, si, axis, a, b), `${axis === 'col' ? 'Columns' : 'Rows'} ungrouped`, axis, 'The selection isn\'t grouped');
    },
    /** The outline bar's level button `n` for rows or columns. */
    outlineLevel(axis, n) { return step((wb, si) => showLevels(wb, si, axis, n), `Level ${n} shown`, axis); },
    /** The outline bar's + or − at summary line `line`: its group shown or hidden. */
    outlineToggle(axis, line) {
      const sh = ctx.workbook().sheet(ctx.sheetIndex());
      const g = outlineOf(sh)[axis === 'row' ? 'rows' : 'cols'].groups.find((x) => x.summary === line);
      const size = axis === 'row' ? (i) => !sh.rows.get(i)?.hidden : (i) => !sh.cols.find((x) => i >= x.min && i <= x.max)?.hidden;
      const hidden = g && size(g.start) === false;
      return hidden ? step((wb, si) => showDetail(wb, si, axis, line), 'Detail shown', axis) : step((wb, si) => hideDetail(wb, si, axis, line), 'Detail hidden', axis);
    },
    /** Data > Show Detail and Hide Detail at the active cell. */
    showDetail() { const axis = axisOf(); return step((wb, si) => showDetail(wb, si, axis, lineOf(axis)), 'Detail shown', axis, 'There is no hidden detail here'); },
    hideDetail() { const axis = axisOf(); return step((wb, si) => hideDetail(wb, si, axis, lineOf(axis)), 'Detail hidden', axis, 'There is no outline here'); },
    /** Ungroup > Clear Outline: one cell selected, the sheet's whole outline; whole rows or columns, theirs; else the
     *  outline of the selection's rows and columns. */
    clearOutline() {
      const s = state.sel;
      const one = s.r0 === s.r1 && s.c0 === s.c1;
      const rows = s.c0 === 0 && s.c1 >= MAX_COLS - 1;
      const cols = s.r0 === 0 && s.r1 >= MAX_ROWS - 1;
      return step((wb, si) => {
        const r = !cols || rows ? clearOutline(wb, si, 'row', one ? 0 : s.r0, one ? Infinity : s.r1) : false;
        const c = !rows || cols ? clearOutline(wb, si, 'col', one ? 0 : s.c0, one ? Infinity : s.c1) : false;
        return r || c;
      }, 'Outline cleared', 'row', 'There is no outline to clear');
    },
    /** Group > Auto Outline: the outline made from the sheet's formulas. */
    autoOutline() { return step((wb, si) => autoOutline(wb, si), 'Outline made', 'row', 'Cannot create an outline: no formulas sum up rows or columns'); },
    /** The Outline Settings: summary rows below their detail, summary columns right of theirs. */
    outlineSettings({ below, right }) { return step((wb, si) => setSummary(wb, si, { below, right }), 'Outline settings changed'); },
    /** Ctrl+8: the outline symbols shown or hidden (no argument: turned over), on the sheet shown; a view setting saved
     *  with the file, as Gridlines is, outside the history. */
    outlineSymbols(on) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const sh = ctx.workbook().sheet(ctx.sheetIndex());
      const now = on ?? sh.outline?.symbols === false;
      setOutlineSymbols(ctx.workbook(), ctx.sheetIndex(), now);
      ctx.reloaded?.();
      ctx.changed?.();
      ctx.announce(now ? 'Outline symbols shown' : 'Outline symbols hidden');
      return true;
    },
    /** Data > Subtotal over the list `area` (subtotal.js's options); refused when it has no rows under its headings.
     *  (Rows are inserted: `moved`, which says a refusal in Excel's words.) */
    subtotal(area, opts) {
      if (area.r1 <= area.r0) { ctx.announce('The list has no rows under its headings'); return false; }
      const done = moved((wb, si) => subtotal(wb, si, area, opts), 'formatRows');
      if (done) ctx.announce('Subtotals added');
      return done;
    },
    /** The Subtotal box's Remove All. */
    removeSubtotals(area) {
      const done = moved((wb, si) => removeSubtotals(wb, si, area), 'formatRows');
      if (done) ctx.announce('Subtotals removed');
      return done;
    },
  };
}
