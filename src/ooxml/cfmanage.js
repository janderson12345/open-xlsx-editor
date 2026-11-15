// The Conditional Formatting Rules Manager's changes, as Excel makes them (measured over COM:
// tools/excel-compare/excel-cf-manage.ps1, test/cfmanage-excel.json). Rules are named by their priority, which is
// unique on a sheet. A rule deleted goes with its extLst part, the others keeping their priorities (Excel doesn't
// number them again); a rule given priority p comes before the rule that had it, every rule from p on one later; a rule
// sent last takes one more than the last; Stop If True is the rule's attribute; a rule given other ranges moves to a
// conditionalFormatting of its own just after the one it was in. Edit Rule: see editRule.
import { kids, kid } from './xml.js';
import { rangeName, parseRange } from './refs.js';
import { addConditional, addPresetDxf, cfRuleXml, ruleElement } from './cfedit.js';
import { X14, XM, EXT_CF } from './cfext.js';
import { x14Rules, dropX14Rules } from './cfbarsedit.js';

/** Every rule of the sheet's part, base and extLst-only: [{ el, priority, cf (its conditionalFormatting), x14 }]. */
function rulesOf(doc) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const base = kids(root, 'conditionalFormatting', ns).flatMap((cf) => kids(cf, 'cfRule', ns).map((el) => ({ el, cf, x14: false })));
  const extra = x14Rules(doc).filter((el) => el.hasAttribute('priority')).map((el) => ({ el, cf: el.parentNode, x14: true }));
  return [...base, ...extra].map((x) => ({ ...x, priority: Number(x.el.getAttribute('priority') || 0) }));
}
const find = (doc, priority) => {
  const r = rulesOf(doc).find((x) => x.priority === priority);
  if (!r) throw new Error('No such rule.');
  return r;
};
/** A base rule's x14 id (a data bar's settings in the extLst). */
const x14IdOf = (el) => [...el.getElementsByTagNameNS(X14, 'id')].map((e) => e.textContent)[0] || null;

/** The extLst's conditional formats ext, and the extLst, gone when left empty. */
function tidyExt(doc) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const extLst = kid(root, 'extLst', ns);
  const ext = kids(extLst, 'ext', ns).find((e) => e.getAttribute('uri') === EXT_CF);
  if (ext && !kids(kid(ext, 'conditionalFormattings', X14), 'conditionalFormatting', X14).length) extLst.removeChild(ext);
  if (extLst && !kids(extLst, 'ext', ns).length) root.removeChild(extLst);
}

/** Delete the rule of priority `p` on sheet si. */
export function deleteRule(wb, si, p) {
  const doc = wb._part(si).doc;
  const r = find(doc, p);
  const id = r.x14 ? r.el.getAttribute('id') : x14IdOf(r.el);
  if (r.x14) dropX14Rules(doc, new Set([id]));
  else {
    r.cf.removeChild(r.el);
    if (!kids(r.cf, 'cfRule', r.cf.namespaceURI).length) r.cf.parentNode.removeChild(r.cf);
    if (id) dropX14Rules(doc, new Set([id]));
  }
  tidyExt(doc);
}

/** Give the rule of priority `p` priority `to` (Excel's FormatCondition.Priority): every other rule from `to` on one
 *  later. `to` 'last': one more than the last. */
export function moveRule(wb, si, p, to) {
  const doc = wb._part(si).doc;
  const all = rulesOf(doc);
  const r = all.find((x) => x.priority === p);
  if (!r) throw new Error('No such rule.');
  if (to === 'last') { r.el.setAttribute('priority', String(Math.max(...all.map((x) => x.priority)) + 1)); return; }
  for (const x of all) if (x !== r && x.priority >= to) x.el.setAttribute('priority', String(x.priority + 1));
  r.el.setAttribute('priority', String(to));
}

/** The rules of sheet si in priority order: their priorities (to name them by). */
export const rulePriorities = (wb, si) => rulesOf(wb._part(si).doc).map((x) => x.priority).sort((a, b) => a - b);

/** Stop If True on or off for the rule of priority `p` (an attribute after the priority, as Excel orders them). */
export function setStopIfTrue(wb, si, p, on) {
  const r = find(wb._part(si).doc, p);
  if (r.x14) return;
  const el = r.el;
  const all = [...el.attributes].map((a) => [a.name, a.value]);
  for (const [n] of all) el.removeAttribute(n);
  const attrs = all.filter(([n]) => n !== 'stopIfTrue');
  const at = attrs.findIndex(([n]) => n === 'priority') + 1;
  if (on) attrs.splice(at, 0, ['stopIfTrue', '1']);
  for (const [n, v] of attrs) el.setAttribute(n, v);
}

/** The rule of priority `p` given the cells `ranges` (Applies to): moved to a conditionalFormatting of its own, just
 *  after the one it was in (that one gone when left empty); an extLst rule's own sqref changed. */
export function setRuleRanges(wb, si, p, ranges) {
  const doc = wb._part(si).doc;
  const r = find(doc, p);
  const sqref = ranges.map(rangeName).join(' ');
  if (r.x14) { const sq = kid(r.cf, 'sqref', XM); if (sq) sq.textContent = sqref; return; }
  const ns = r.cf.namespaceURI;
  const id = x14IdOf(r.el);
  const cf = doc.createElementNS(ns, r.cf.tagName);
  cf.setAttribute('sqref', sqref);
  r.cf.parentNode.insertBefore(cf, r.cf.nextSibling);
  r.cf.removeChild(r.el);
  cf.appendChild(r.el);
  if (!kids(r.cf, 'cfRule', ns).length) r.cf.parentNode.removeChild(r.cf);
  // (A data bar's settings in the extLst go to the same cells.)
  if (id) for (const x of x14Rules(doc)) if (x.getAttribute('id') === id) { const sq = kid(x.parentNode, 'sqref', XM); if (sq) sq.textContent = sqref; }
}

/**
 * Edit Rule: the rule of priority `p` given the settings of `rule` (cfedit.js's, `preset` null to keep the rule's own
 * format). As Excel edits one (tools/excel-compare/excel-cf-edit.ps1, FormatCondition.Modify): in its place among the
 * rules of its conditionalFormatting, its ranges and its format (dxfId) kept, its type and criteria new. COM also sends
 * an edited rule last and sets Stop If True, as its Add does; Excel's Rules Manager keeps its priority and its own
 * Stop If True (measured in its window, test/cfedit-ui-excel.json), and so does this. A data bar or icon set (with its part in the extLst) is replaced instead: the new
 * rule added on the same ranges and given the old one's place.
 */
export function editRule(wb, si, p, rule) {
  const doc = wb._part(si).doc;
  const r = find(doc, p);
  const sqref = r.x14 ? [...r.cf.getElementsByTagNameNS(XM, 'sqref')].map((e) => e.textContent)[0] || '' : r.cf.getAttribute('sqref');
  const ranges = sqref.split(/\s+/).map(parseRange).filter(Boolean);
  const bars = (k) => k === 'dataBar' || k === 'iconSet';
  if (r.x14 || x14IdOf(r.el) || bars(r.el.getAttribute('type')) || bars(rule.kind)) {
    const stop = !r.x14 && r.el.getAttribute('stopIfTrue') === '1';
    deleteRule(wb, si, p);
    addConditional(wb, si, ranges, { ...rule, preset: rule.preset || 'lightRed' });
    // (Added first, every other rule one later: put back after the rules that came before it.)
    moveRule(wb, si, 1, p + 1);
    if (stop && !bars(rule.kind)) setStopIfTrue(wb, si, p + 1, true);
    return;
  }
  const own = r.el.getAttribute('dxfId');
  const dxfId = rule.kind === 'colorScale' ? null : rule.preset == null && own != null ? Number(own) : addPresetDxf(wb, rule.preset || 'lightRed');
  const first = ranges[0];
  let xml = cfRuleXml(rule, dxfId, p, rangeName({ r0: first.r0, c0: first.c0, r1: first.r0, c1: first.c0 }));
  if (r.el.getAttribute('stopIfTrue') === '1' && rule.kind !== 'colorScale') xml = xml.replace(/ priority="(\d+)"/, ' priority="$1" stopIfTrue="1"');
  r.cf.replaceChild(ruleElement(doc, xml), r.el);
}
