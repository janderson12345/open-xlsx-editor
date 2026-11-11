// Commands on the workbook's sheets (commands.js shares its state with them): show, next, insert, rename, move,
// hide, colour and delete, through sheetOp.
import { cellSelection } from './selection.js';
import { SheetError, nearest } from '../ooxml/sheetops.js';
import { copySheet } from '../ooxml/sheetcopy.js';

/** The sheet commands, and sheetOp (for the others): `env` is createCommands's shared state and helpers. */
export function sheetCommands(env) {
  const { commands, ctx, state, show, idOf, indexOf } = env;
  /** After the sheets changed: the tabs again, and the sheet shown `show()` if given, else the sheet `shownId` wherever
   *  it went, else (gone or hidden) the nearest to where it was (`was`). */
  function sheetsMoved(wb, shownId, was, show = null) {
    for (const id of state.views.keys()) if (indexOf(id) < 0) state.views.delete(id);
    const now = indexOf(shownId);
    ctx.sheetsChanged(now); // (the tabs again, and the sheet shown at its new index, or -1 if it went)
    let target = show ? show() : now;
    if (target < 0 || !wb.canShow(target)) target = nearest(wb, target < 0 ? Math.min(was, wb.sheets.length - 1) : target);
    if (target !== now || now < 0) commands.showSheet(target);
  }
  /** Run a change to the workbook's sheets, as one step of the history (the workbook before and after, its sheets
   *  with it, and the sheet shown each side). Afterwards the sheet shown is `show()` if given, else the same sheet
   *  wherever it went, else (gone or hidden) its nearest neighbour. */
  function sheetOp(run, show = null, quiet = false) {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    const wb = ctx.workbook();
    const was = ctx.sheetIndex();
    const shownId = idOf(was);
    const before = wb.snapshot();
    try {
      run(wb);
    } catch (e) {
      wb.restore(before);
      ctx.reloaded?.();
      if (!(e instanceof SheetError) || quiet) throw e;
      ctx.alert(e.message);
      return false;
    }
    ctx.edited?.({ book: true });
    sheetsMoved(wb, shownId, was, show);
    state.undo.push({ sheet: shownId, sheets: true, snap: true, before, after: wb.snapshot(), shown: { before: shownId, after: idOf(ctx.sheetIndex()) } });
    state.redo = [];
    ctx.changed?.();
    return true;
  }
  env.sheetOp = sheetOp;
  env.sheetsMoved = sheetsMoved;
  return {
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
    /** Move or Copy > Create a copy: sheet i copied before index `before` (sheetcopy.js), the copy shown. */
    copySheet(i, before) {
      let at = -1;
      return sheetOp((wb) => { at = copySheet(wb, i, before); }, () => at);
    },
    /** Rename sheet i (Excel's refusals said in its words). */
    renameSheet(i, name) { return sheetOp((wb) => wb.renameSheet(i, name)); },
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
    deleteSheet(i) { return sheetOp((wb) => wb.deleteSheet(i)); },
  };
}
