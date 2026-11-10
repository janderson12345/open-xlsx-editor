// Home > Clear as Excel does it (measured over COM, docs/FORMATS.md): Clear All empties the cells and takes their
// formats away, Clear Contents empties them, Clear Formats takes the formats away; All and Formats also unmerge the
// merged ranges inside. A selection that covers part of a merged range is refused, whichever it is.
import { cellsOf } from './formatops.js';

export const CLEAR_REFUSAL = "We can't do that to a merged cell.";

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const inside = (a, b) => a.r0 >= b.r0 && a.r1 <= b.r1 && a.c0 >= b.c0 && a.c1 <= b.c1;

/** The cells a selection stands for: a merged range's whole, when the selection is one of its cells (as Excel selects
 *  a merged cell). */
export function clearArea(sheet, sel) {
  const one = sel.r0 === sel.r1 && sel.c0 === sel.c1;
  const m = one && sheet.merges.find((x) => sel.r0 >= x.r0 && sel.r0 <= x.r1 && sel.c0 >= x.c0 && sel.c0 <= x.c1);
  return m ? { r0: m.r0, c0: m.c0, r1: m.r1, c1: m.c1 } : { r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 };
}

/**
 * Clear `what` ('all', 'contents' or 'formats') of selection `sel` on sheet si: a step as clipops.js's ({ si, contents,
 * styles, merges }), or { error } with Excel's refusal. Only the cells there are are emptied, so a whole column is
 * quick.
 */
export function clearChanges(wb, si, sel, what) {
  const sheet = wb.sheet(si);
  const area = clearArea(sheet, sel);
  if (sheet.merges.some((m) => overlaps(m, area) && !inside(m, area))) return { error: CLEAR_REFUSAL };
  const out = { si, contents: [], styles: [], merges: sheet.merges, area };
  if (what !== 'formats') {
    for (const [r, line] of sheet.cells) {
      if (r < area.r0 || r > area.r1) continue;
      for (const [c, cell] of line) if (c >= area.c0 && c <= area.c1 && (cell.v != null || cell.f != null)) out.contents.push({ r, c, value: null });
    }
  }
  if (what !== 'contents') {
    out.styles = cellsOf(sheet, area).filter(({ r, c }) => sheet.cells.get(r)?.get(c)?.s).map(({ r, c }) => ({ r, c, s: 0 }));
    out.merges = sheet.merges.filter((m) => !inside(m, area));
  }
  return out;
}
