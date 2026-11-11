// Spilling, as Excel 365 does it (measured: tools/excel-compare/excel-dynamic.ps1, test/dynamic-excel.json). A dynamic
// array formula whose result is an array fills the cells below and to the right of it with the rest: its spill range.
// A cell there that holds something of its own (a value, a formula, another formula's spill), a merged range or the
// sheet's edge stops it: the formula shows #SPILL! and fills nothing. When the result shrinks, grows or goes, the cells
// it no longer fills are emptied; formulas reading them, or reading the spill as D1#, are worked out again (engine.js).
import { E } from './values.js';
import { MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';

const key = (si, r, c) => `${si}:${r}:${c}`;
const inside = (a, r, c) => !!a && r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1;
const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** Whether something other than formula `e`'s own spill is in the way of range `want`. */
function blocked(engine, e, want) {
  if (want.r1 >= MAX_ROWS || want.c1 >= MAX_COLS) return true;
  const sheet = engine.book.sheet(e.si);
  if (sheet.merges.some((m) => overlaps(m, want)) || sheet.arrays.some((a) => overlaps(a, want))) return true;
  for (let r = want.r0; r <= want.r1; r++) {
    const line = sheet.cells.get(r);
    for (let c = want.c0; c <= want.c1; c++) {
      if (r === e.r && c === e.c) continue;
      const owner = engine.spilledBy.get(key(e.si, r, c));
      if (owner) { if (owner !== e) return true; continue; }
      const cell = line?.get(c);
      if (cell && (cell.v != null || cell.f != null)) return true;
    }
  }
  return false;
}

/** The cells of `e`'s spill, but its own, no longer its (not in `keep`): emptied and forgotten. */
function clear(engine, e, keep = null) {
  const s = e.spill;
  if (!s) return;
  for (let r = s.r0; r <= s.r1; r++) {
    for (let c = s.c0; c <= s.c1; c++) {
      if ((r === e.r && c === e.c) || inside(keep, r, c)) continue;
      const k = key(e.si, r, c);
      if (engine.spilledBy.get(k) !== e) continue;
      engine.spilledBy.delete(k);
      if (engine.at.get(k) === e) engine.at.delete(k);
      engine.book.clearSpilled(e.si, r, c);
    }
  }
}

/** Formula `e`'s spill range now `next` (or none): the cells taken, the anchor's file form, the moved area noted. */
function settle(engine, e, next, wanted = null) {
  const was = e.spill;
  if (next) {
    for (let r = next.r0; r <= next.r1; r++) {
      for (let c = next.c0; c <= next.c1; c++) {
        if (r === e.r && c === e.c) continue;
        engine.spilledBy.set(key(e.si, r, c), e);
        engine.at.set(key(e.si, r, c), e);
      }
    }
  }
  e.spill = next;
  e.wanted = wanted;
  if (wanted) engine.blocked.add(e); else engine.blocked.delete(e);
  engine.book.setDynamic(e.si, e.r, e.c, next || (wanted ? { r0: e.r, c0: e.c, r1: e.r, c1: e.c } : null));
  const same = (a, b) => (!a && !b) || (a && b && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1);
  if (!same(was, next)) for (const a of [was, next]) if (a) engine.moved.push({ si: e.si, ...a });
}

/** Put a dynamic formula's result (evaluate.js dynamicResult) in place: one value in its cell, or an array spilled
 *  from it (#SPILL! when something is in the way). */
export function place(engine, e, res) {
  if (!res.rows) {
    clear(engine, e);
    engine.book.writeResult(e.si, e.r, e.c, res.value);
    settle(engine, e, null);
    return;
  }
  const rows = res.rows;
  const want = { r0: e.r, c0: e.c, r1: e.r + rows.length - 1, c1: e.c + rows[0].length - 1 };
  if (blocked(engine, e, want)) {
    clear(engine, e);
    engine.book.writeResult(e.si, e.r, e.c, E.spill);
    settle(engine, e, null, want);
    return;
  }
  clear(engine, e, want);
  for (let r = want.r0; r <= want.r1; r++) for (let c = want.c0; c <= want.c1; c++) engine.book.writeResult(e.si, r, c, rows[r - want.r0][c - want.c0]);
  settle(engine, e, want);
}

/** A dynamic formula taken away (cleared, or given something else): the cells it spilled into emptied. */
export function release(engine, e) {
  clear(engine, e);
  if (e.spill) engine.moved.push({ si: e.si, ...e.spill });
  engine.blocked.delete(e);
  e.spill = null;
}

/** The spill range of the formula in cell (r, c) of sheet si (D1#), or null when it spills nothing there. */
export function spillOf(engine, si, r, c) {
  const e = engine.entries?.get(key(si, r, c));
  return e && e.dynamic && e.spill ? e.spill : null;
}

/** Cells just changed by hand (`rects`): any that a spill filled is the person's now, and the formula whose spill it
 *  was, or whose spill it stood in the way of, is worked out again (it may now be #SPILL!, or spill again). */
export function touched(engine, rects) {
  const again = new Set();
  for (const a of rects) {
    if ((a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) <= 4096) {
      for (let r = a.r0; r <= a.r1; r++) {
        for (let c = a.c0; c <= a.c1; c++) {
          const owner = engine.spilledBy.get(key(a.si, r, c));
          if (owner && !(owner.r === r && owner.c === c)) { engine.spilledBy.delete(key(a.si, r, c)); again.add(owner); }
        }
      }
    }
    for (const e of engine.blocked) if (e.si === a.si && overlaps(e.wanted, a)) again.add(e);
  }
  return again;
}
