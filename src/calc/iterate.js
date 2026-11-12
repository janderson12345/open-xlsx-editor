// Circular references, as Excel works them out (measured with a hidden Excel, probe iter.ps1 and iter2.ps1). With
// iterative calculation off (calcPr iterate), a formula in a circle isn't worked out: it keeps its value, and one just
// entered shows 0. With it on, every recalculation goes round: the formulas in circles, and those reading them, worked
// out in turn, row by row of each sheet in order (B1 before A2), each pass reading the last; a formula with no value
// yet reads as 0 (=IF(J1="",1,J1&"x") makes "0x" first). The passes stop when no number changed by Maximum Change
// or more (the whole workbook's: one formula still changing keeps them all going), or after Maximum Iterations.
// Excel works them out again at every recalculation, whatever changed (=A1+1 is 100, then 200 after any edit).
import { cellValue } from './values.js';
import { iteration } from '../ooxml/calcmode.js';

const valueAt = (calc, e) => cellValue(calc.book.sheet(e.si).cells.get(e.r)?.get(e.c));

/** The formulas of `order` (each after those it reads) that are in `circular` or read one, and the others. */
function split(calc, order, circular, dirty) {
  const round = new Set(circular);
  const rest = [];
  for (const e of order) {
    if (round.has(e)) continue;
    let reads = false;
    for (const p of calc._precedents(e, dirty)) if (round.has(p)) { reads = true; break; }
    if (reads) round.add(e); else rest.push(e);
  }
  return { round: [...round].sort((a, b) => a.si - b.si || a.r - b.r || a.c - b.c), rest };
}

/** Work out `order` (from Calc._run), its circles as the workbook's settings have them. Returns the circular
 *  formulas to announce: none when they were iterated. */
export function runWithCircles(calc, order, circular, dirty) {
  const { on, count, delta } = iteration(calc.book);
  calc.cyclic = on ? new Set(circular) : new Set();
  if (!circular.size) { for (const e of order) calc._evaluate(e); return []; }
  if (!on) {
    for (const e of order) if (!circular.has(e)) calc._evaluate(e);
    for (const e of circular) if (valueAt(calc, e) === null) calc.book.writeResult(e.si, e.r, e.c, 0);
    return [...circular];
  }
  const { round, rest } = split(calc, order, circular, dirty);
  for (const e of rest) calc._evaluate(e);
  for (const e of round) if (valueAt(calc, e) === null) calc.book.writeResult(e.si, e.r, e.c, 0);
  for (let n = 0; n < count; n++) {
    let most = 0;
    for (const e of round) {
      const was = valueAt(calc, e);
      calc._evaluate(e);
      const now = valueAt(calc, e);
      if (typeof was === 'number' && typeof now === 'number') most = Math.max(most, Math.abs(now - was));
    }
    if (most < delta) break;
  }
  return [];
}
