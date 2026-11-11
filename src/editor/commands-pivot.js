// Insert > PivotTable (commands.js shares its state with it): a new PivotTable made (pivotnew.js) on a new sheet or at a
// cell of one there is, its row labels aligned left and its columns fitted, as Excel leaves a new one.
import { createPivot, pivotSourceFields, layoutRefusal, applyPivot } from '../ooxml/pivotnew.js';
import { readPivots } from '../ooxml/pivot.js';
import { SheetError } from '../ooxml/sheetops.js';
import { pxWidth } from '../ooxml/sizes.js';
import { parseRange } from '../ooxml/refs.js';
import { areasSelection } from './selection.js';

/** Excel's refusal of a PivotTable over another. */
export const PIVOT_OVERLAP = 'A PivotTable report cannot overlap another PivotTable report.';
/** Excel's question before a new PivotTable covers cells with something in them (Excel 365's wording, from its
 *  dialog, which COM doesn't reach: not measured). */
export const PIVOT_REPLACE = 'There\'s already data in the destination. Do you want to replace it?';

/** These commands: `env` is createCommands's shared state and helpers (commands.js). */
export function pivotCommands(env) {
  const { commands, ctx, state, idOf, sheetOp } = env;

  /** The cells and look of a new pivot: its cells made, its row labels aligned left, its columns fitted. */
  function finish(wb, out) {
    applyPivot(wb, out);
    ctx.restyled?.(); // (the grid learns the new cell format)
  }
  /** The pivot's columns fitted (on the sheet shown, where the grid measures text), as Excel fits a new one. */
  function fit(wb, si, area) {
    ctx.reloaded?.();
    for (let c = area.c0; c <= area.c1; c++) {
      const px = ctx.fitWidth?.(c);
      if (px) wb.setLines(si, 'col', c, c, { width: pxWidth(px), bestFit: true });
    }
    ctx.reloaded?.();
  }
  /** Whether the new pivot's cells (`area`; its part `own`) would cover another pivot, or cells with something in
   *  them. */
  function covers(wb, si, area, own) {
    for (const p of readPivots(wb)) {
      if (p.si !== si || p.path === own) continue;
      const loc = parseRange(/<location ref="([^"]+)"/.exec(wb.pkg.text(p.path))?.[1] || '');
      if (loc && loc.r0 <= area.r1 && area.r0 <= loc.r1 && loc.c0 <= area.c1 && area.c0 <= loc.c1) return 'pivot';
    }
    const sheet = wb.sheet(si);
    for (let r = area.r0; r <= area.r1; r++) for (let c = area.c0; c <= area.c1; c++) { const v = sheet.cells.get(r)?.get(c); if (v && (v.v != null || v.f != null)) return 'data'; }
    return null;
  }

  return {
    /** Insert > PivotTable: a PivotTable of `src` ({ si, r0, c0, r1, c1 }) with `layout` (pivotnew.js), at `at` ({ si,
     *  r, c }) or, with none, on a new sheet before the one shown at A3 (as Excel's New Worksheet puts it). On a sheet
     *  there is, one step of the history; Excel's question first ({ ask }) when it would cover cells with something in
     *  them, unless `replace`. A new sheet clears the history, as inserting and deleting sheets here do. */
    newPivot(src, layout, at = null, { replace = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const fields = pivotSourceFields(wb, src);
      const why = fields.error || layoutRefusal(layout);
      if (why) { ctx.alert(why); return false; }
      let out = null;
      if (at) {
        const before = wb.snapshot();
        out = createPivot(wb, src, at, layout);
        const blocked = !out.error && covers(wb, at.si, out.changes.find((ch) => ch.si === at.si).area, out.path);
        if (out.error || blocked === 'pivot' || (blocked === 'data' && !replace)) {
          wb.restore(before);
          if (out.error) { ctx.alert(out.error); return false; }
          if (blocked === 'pivot') { ctx.alert(PIVOT_OVERLAP); return false; }
          return { ask: PIVOT_REPLACE };
        }
        finish(wb, out);
        if (at.si !== ctx.sheetIndex()) commands.showSheet(at.si);
        fit(wb, at.si, out.changes[0].area);
        state.undo.push({ sheet: idOf(at.si), snap: true, before, after: wb.snapshot() });
        state.redo = [];
        ctx.edited?.({ book: true });
        ctx.changed?.();
      } else {
        let si = -1;
        const done = sheetOp((book) => {
          si = book.addSheet(ctx.sheetIndex());
          // (The sheets from the new one on moved up one: the source's with them.)
          out = createPivot(book, { ...src, si: src.si >= si ? src.si + 1 : src.si }, { si, r: 2, c: 0 }, layout);
          if (out.error) throw new SheetError(out.error);
          finish(book, out);
        }, () => si);
        if (!done) return false;
        fit(wb, si, out.changes[0].area);
        // (The widths fitted on the new sheet, once shown, belong to the same step.)
        state.undo[state.undo.length - 1].after = wb.snapshot();
      }
      const area = out.changes[0].area;
      env.show(areasSelection([{ r0: area.r0, c0: area.c0, r1: area.r0, c1: area.c0 }]), null, `${out.name} made, `);
      return true;
    },
  };
}
