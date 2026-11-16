// Home > Cell Styles' New Cell Style, Modify, Duplicate, Delete and Merge Styles as commands (commands.js shares its state and helpers as `env`): each
// one step of the history, a snapshot of the workbook with its styles part before and after, since a style changed
// in place changes the cell formats that use it (ooxml/cellstyledefine.js).
import { newCellStyle, modifyCellStyle } from '../ooxml/cellstyledefine.js';
import { deleteCellStyle, duplicateCellStyle, mergeCellStyles } from '../ooxml/cellstylemanage.js';
import { Workbook } from '../ooxml/workbook.js';

export function styleCommands(env) {
  const { ctx, state, allowed, idOf } = env;
  /** `run(wb)` ({ error } to refuse) as one step of the history. */
  const styleStep = (run, said) => {
    if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
    if (!allowed('formatCells')) return false;
    const wb = ctx.workbook();
    const before = wb.snapshot({ styles: true });
    const out = run(wb);
    if (out?.error) { wb.restore(before); ctx.announce(out.error); return out; }
    state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot({ styles: true }) });
    state.redo = [];
    ctx.edited?.({ book: true });
    ctx.restyled?.();
    ctx.reloaded?.();
    ctx.changed?.();
    ctx.announce(said);
    return out || true;
  };
  return {
    /** New Cell Style: style `name` made from the active cell's format by example, with `includes` ({ number,
     *  alignment, font, border, fill, protection }: false for a part left out) and `change` (Format..., the Format
     *  Cells box's change). The cell keeps its own style, as in Excel. */
    newCellStyle(name, includes = {}, change = null) {
      const s = ctx.workbook().sheet(ctx.sheetIndex()).cells.get(state.sel.ar)?.get(state.sel.ac)?.s || 0;
      return styleStep((wb) => newCellStyle(wb, s, name, includes, change), `Style ${name} made`);
    },
    /** Modify: style `name` given `change` and `includes` (null: as it was), renamed to `rename`; the cells using it
     *  change with it. */
    modifyCellStyle(name, change = null, includes = null, rename = null) {
      return styleStep((wb) => modifyCellStyle(wb, name, change, includes, rename), `Style ${rename || name} changed`);
    },
    /** Duplicate: style `newName` made a copy of style `name`, with `includes` and `change` as the Style box gives them
     *  (null: as the style's). */
    duplicateCellStyle(name, newName, includes = null, change = null) {
      return styleStep((wb) => duplicateCellStyle(wb, name, newName, includes, change), `Style ${newName} made`);
    },
    /** Delete: style `name` taken out (a built-in one kept hidden, as Excel keeps it); the cells using it back to
     *  Normal but for the parts they hold as their own. Normal is refused. */
    deleteCellStyle(name) {
      return styleStep((wb) => deleteCellStyle(wb, name), `Style ${name} deleted`);
    },
    /** Merge Styles: the styles of `source` (another workbook: a Workbook, or an .xlsx file's bytes) taken into this
     *  one; `replace`: Yes to "Merge styles that have the same names?" (false: this workbook's kept). */
    mergeCellStyles(source, replace = true) {
      let from = source;
      if (!(from instanceof Workbook)) {
        try { from = Workbook.fromBytes(source instanceof ArrayBuffer ? new Uint8Array(source) : source); } catch { ctx.announce('That workbook could not be read'); return false; }
      }
      return styleStep((wb) => {
        const out = mergeCellStyles(wb, from, replace);
        return out.added.length || out.replaced.length ? out : { error: 'No styles to merge' };
      }, 'Styles merged');
    },
  };
}
