// Formulas > Trace Precedents and Trace Dependents: what a formula reads and what reads a cell, as Excel's arrows
// show them (measured with a hidden Excel's ShowPrecedents, ShowDependents and NavigateArrow): one arrow for each
// different cell or range read (A1 read twice is one), a defined name traced to what it names, and everything on other
// sheets as one arrow of its own. The arrows themselves are the editor's (editor/arrows.js); none is saved.

/** Sheet si's cell (r, c): the precedents of its formula, { here: [rects], elsewhere: [{ si, ...rect }] }, null when
 *  it holds no formula. */
export function precedents(wb, si, r, c) {
  const calc = wb.calc;
  calc.build();
  const e = calc.entries.get(`${si}:${r}:${c}`);
  if (!e) return null;
  const seen = new Set();
  const here = [];
  const elsewhere = [];
  for (const d of e.deps) {
    const k = `${d.si}:${d.r0}:${d.c0}:${d.r1}:${d.c1}`;
    if (seen.has(k)) continue;
    seen.add(k);
    (d.si === si ? here : elsewhere).push({ ...d });
  }
  return { here, elsewhere };
}

/** The formula cells that read sheet si's cell (r, c): { here: [{ r, c }], elsewhere: [{ si, r, c }] }. */
export function dependents(wb, si, r, c) {
  const calc = wb.calc;
  calc.build();
  const here = [];
  const elsewhere = [];
  for (const e of calc.readers({ si, r0: r, c0: c, r1: r, c1: c })) {
    if (e.si === si && e.r === r && e.c === c) continue;
    (e.si === si ? here : elsewhere).push({ si: e.si, r: e.r, c: e.c });
  }
  return { here, elsewhere };
}

/** The formula cells of a range of sheet si (at most `limit` cells looked at), to trace on from. */
export function formulaCells(wb, si, rect, limit = 4096) {
  const out = [];
  const cells = wb.sheet(si).cells;
  if ((rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1) > limit) return out;
  for (let r = rect.r0; r <= rect.r1; r++) for (let c = rect.c0; c <= rect.c1; c++) if (cells.get(r)?.get(c)?.f != null) out.push({ r, c });
  return out;
}
