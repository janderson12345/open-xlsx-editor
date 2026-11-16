// Home > Cell Styles' New Cell Style and Modify as commands (commands.js shares its state and helpers as `env`): each
// one step of the history, a snapshot of the workbook with its styles part before and after, since a style changed
// in place changes the cell formats that use it (ooxml/cellstyledefine.js).
import { newCellStyle, modifyCellStyle } from '../ooxml/cellstyledefine.js';

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
  };
}
