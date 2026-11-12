// Trace Precedents and Trace Dependents a level at a time, as Excel's do: the first click draws the active cell's
// arrows, each click after draws the next level's from the formulas just reached, until nothing more is found. A
// click on another cell starts its own trace, the arrows drawn so far staying. ooxml/trace.js finds what is read and
// what reads; arrows.js draws them.
import { precedents, dependents, formulaCells } from '../ooxml/trace.js';

const one = (r, c) => ({ r0: r, c0: c, r1: r, c1: c });

/** Whether a range of sheet si holds an error value (its arrow is red). */
function hasError(wb, si, rect) {
  const cells = wb.sheet(si).cells;
  if ((rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1) > 4096) return false;
  for (let r = rect.r0; r <= rect.r1; r++) for (let c = rect.c0; c <= rect.c1; c++) if (cells.get(r)?.get(c)?.t === 'e') return true;
  return false;
}

/** One more level of `kind` ('precedents' or 'dependents') on sheet si from `at` ({ r, c }, the active cell), its
 *  arrows added to `trace` ({ arrows, precedents, dependents }). Returns how many arrows were added. */
export function traceLevel(wb, si, trace, kind, at) {
  let side = trace[kind];
  if (!side || side.from.r !== at.r || side.from.c !== at.c) side = trace[kind] = { from: at, frontier: [at], done: new Set() };
  const next = [];
  let added = 0;
  const add = (arrow) => { trace.arrows.push({ kind, ...arrow }); added++; };
  for (const cell of side.frontier) {
    const k = `${cell.r}:${cell.c}`;
    if (side.done.has(k)) continue;
    side.done.add(k);
    if (kind === 'precedents') {
      const p = precedents(wb, si, cell.r, cell.c);
      if (!p) continue;
      for (const rect of p.here) {
        add({ from: rect, to: one(cell.r, cell.c), color: hasError(wb, si, rect) ? 'red' : 'blue' });
        next.push(...formulaCells(wb, si, rect));
      }
      if (p.elsewhere.length) add({ from: null, to: one(cell.r, cell.c), color: 'black' });
    } else {
      const d = dependents(wb, si, cell.r, cell.c);
      const red = hasError(wb, si, one(cell.r, cell.c));
      for (const x of d.here) { add({ from: one(cell.r, cell.c), to: one(x.r, x.c), color: red ? 'red' : 'blue' }); next.push(x); }
      if (d.elsewhere.length) add({ from: one(cell.r, cell.c), to: null, color: 'black' });
    }
  }
  side.frontier = next.filter((x) => !side.done.has(`${x.r}:${x.c}`));
  return added;
}
