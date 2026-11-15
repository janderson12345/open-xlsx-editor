// The Excel 2010 half of conditional formats a paste or a cut brings (clipcf.js): a data bar's settings, and icon sets
// Excel 2010 added, kept in the worksheet's extLst as x14 rules, each with an id (a data bar's base rule names it) and
// its own xm:sqref. Measured over COM: a data bar copied from A1:A3 to C1 came as a new base rule and a new x14 rule
// over C1:C3, a new id in both; three stars (x14 only) as a new x14 rule first in priority; extended along a line and
// moved by a cut, both halves' cells change together.
import { kids, kid } from './xml.js';
import { rangeName } from './refs.js';
import { X14, XM } from './cfext.js';
import { x14Rules, addX14Rule } from './cfbarsedit.js';

/** A new rule id, as Excel writes them: {8-4-4-4-12} in capitals. */
export const newId = () => `{${globalThis.crypto.randomUUID().toUpperCase()}}`;

/** The x14 rule of id `id` in the sheet's part, or null. */
export const x14Rule = (doc, id) => x14Rules(doc).find((r) => r.getAttribute('id') === id) || null;

/** A base rule's x14 id (a data bar's), or null. */
export const baseX14Id = (rule) => [...rule.getElementsByTagNameNS(X14, 'id')].map((e) => e.textContent)[0] || null;

/** The xm:sqref of the x14 conditionalFormatting holding rule `r`. */
const sqrefOf = (r) => kid(r.parentNode, 'sqref', XM);

/** The x14 rule with id `id` given the cells `ranges`. */
export function setX14Ranges(doc, id, ranges) {
  const r = x14Rule(doc, id);
  const sq = r && sqrefOf(r);
  if (sq) sq.textContent = ranges.map(rangeName).join(' ');
}

/** The x14 rules that are rules of their own (Excel 2010's icon sets: a priority, no base rule), by their
 *  conditionalFormatting: [{ cf, rules, ranges }]. */
export function x14Blocks(doc, parse) {
  const seen = new Map();
  for (const r of x14Rules(doc)) {
    if (!r.hasAttribute('priority')) continue;
    const cf = r.parentNode;
    if (!seen.has(cf)) seen.set(cf, { cf, rules: kids(cf, 'cfRule', X14), ranges: (sqrefOf(r)?.textContent || '').split(/\s+/).map(parse).filter(Boolean) });
  }
  return [...seen.values()];
}

/** Rule `el` (an x14 cfRule, from any document) added to the sheet's extLst over `ranges` with id `id` (and
 *  `priority`, when it has one). */
export function addX14Clone(doc, el, ranges, id, priority = null) {
  addX14Rule(doc, '', ranges.map(rangeName).join(' '));
  const cfs = [...doc.getElementsByTagNameNS(X14, 'conditionalFormatting')];
  const cf = cfs[cfs.length - 1];
  const rule = doc.importNode(el, true);
  rule.setAttribute('id', id);
  if (priority != null) rule.setAttribute('priority', String(priority));
  cf.insertBefore(rule, kid(cf, 'sqref', XM));
}
