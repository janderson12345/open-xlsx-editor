// A cell whose formula has HYPERLINK in it, as Excel treats one (measured over COM, docs/LINKS.md): typed, the cell
// gets the Hyperlink style when HYPERLINK is at the formula's top level (=HYPERLINK(...), =1+HYPERLINK(...); not inside
// another function, =IF(TRUE,HYPERLINK(...))), and no <hyperlink> is written: a click follows HYPERLINK's
// link_location, worked out where the cell is ("#Sheet2!B5" a place in the workbook).
import { parseFormula } from './parse.js';
import { result } from './evaluate.js';
import { toText, isError } from './values.js';
import { hyperlinkAt } from '../ooxml/links.js';

const isHyperlink = (n) => n && n.k === 'fn' && n.name.toUpperCase() === 'HYPERLINK';
const children = (n) => Object.values(n).flatMap((v) => (Array.isArray(v) ? v : [v])).filter((v) => v && typeof v === 'object' && typeof v.k === 'string');

/** Whether a formula (as the file has it, without "=") has HYPERLINK at its top level: through operators and
 *  brackets, not inside another function. */
export function topLevelHyperlink(formula) {
  if (!formula || !/HYPERLINK\s*\(/i.test(formula)) return false;
  let tree;
  try { tree = parseFormula(formula); } catch { return false; }
  const walk = (n) => isHyperlink(n) || (n.k !== 'fn' && children(n).some(walk));
  return walk(tree);
}

/** The first HYPERLINK call in a tree, depth first, or null. */
function firstHyperlink(n) {
  if (isHyperlink(n)) return n;
  for (const k of children(n)) { const f = firstHyperlink(k); if (f) return f; }
  return null;
}

/** Cell (r, c) of sheet si's HYPERLINK, as links.js describes a hyperlink ({ address, location, display, tooltip }),
 *  or null: its link_location worked out in the cell's place. */
export function formulaLink(wb, si, r, c) {
  const f = wb.sheet(si).cells.get(r)?.get(c)?.f;
  if (!f || !/HYPERLINK\s*\(/i.test(f)) return null;
  let call;
  try { call = firstHyperlink(parseFormula(f)); } catch { return null; }
  if (!call?.args?.[0]) return null;
  let to;
  try { to = result(call.args[0], wb.calc.context({ si, r, c })); } catch { return null; }
  if (to == null || isError(to)) return null;
  const text = toText(to);
  if (!text) return null;
  return text.startsWith('#') ? { address: null, location: text.slice(1), display: null, tooltip: null } : { address: text, location: null, display: null, tooltip: null };
}

/** The hyperlink a click on cell (r, c) follows: its own (links.js), else its formula's HYPERLINK. */
export const cellLink = (wb, si, r, c) => hyperlinkAt(wb, si, r, c) || formulaLink(wb, si, r, c);
