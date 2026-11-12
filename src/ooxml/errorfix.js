// What Excel's error button offers for a cell with an error indicator (errorcheck.js): Ignore Error, written as Excel
// writes it (measured: one ignoredError for each rule, its cells joined into ranges, sqref="A1:A2 A4"
// numberStoredAsText="1", in the order the rules were first ignored), and the fixes as changes for the editor to
// make: Convert to Number, Copy Formula from Above or Left, Lock Cell.
import { kid, kids, attr } from './xml.js';
import { parseRange, rangeName } from './refs.js';
import { textToNumber } from '../calc/values.js';
import { shiftFormula } from '../calc/tokens.js';

// (ignoredErrors comes after these in a worksheet; everything else in the list below goes after it.)
const AFTER = ['smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** Cells [{ r, c }] as Excel joins them for a sqref: runs down each column, then runs of the same rows side by side. */
export function sqrefOf(cells) {
  const byCol = new Map();
  for (const { r, c } of cells) { if (!byCol.has(c)) byCol.set(c, new Set()); byCol.get(c).add(r); }
  const runs = [];
  for (const c of [...byCol.keys()].sort((a, b) => a - b)) {
    const rows = [...byCol.get(c)].sort((a, b) => a - b);
    for (let k = 0; k < rows.length;) {
      let e = k;
      while (e + 1 < rows.length && rows[e + 1] === rows[e] + 1) e++;
      runs.push({ r0: rows[k], r1: rows[e], c0: c, c1: c });
      k = e + 1;
    }
  }
  // (Side by side: a run beside one of the same rows widens it.)
  const out = [];
  for (const g of runs) {
    const prev = out.find((o) => o.r0 === g.r0 && o.r1 === g.r1 && o.c1 === g.c0 - 1);
    if (prev) prev.c1 = g.c1; else out.push({ ...g });
  }
  return out.map(rangeName).join(' ');
}

/** Ignore Error: rule `rule` (errorcheck.js's names) ignored for cells [{ r, c }] of sheet si. */
export function ignoreErrors(wb, si, cells, rule) {
  const doc = wb._part(si).doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const make = (name) => doc.createElementNS(ns, root.prefix ? `${root.prefix}:${name}` : name);
  let list = kid(root, 'ignoredErrors', ns);
  if (!list) {
    list = make('ignoredErrors');
    const next = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName));
    root.insertBefore(list, next || null);
  }
  // (The rule's own element, its cells joined with the new ones; only it of the rules, as Excel keeps one per rule.)
  let el = kids(list, 'ignoredError', ns).find((x) => attr(x, rule) === '1' && [...x.attributes].filter((a) => a.name !== 'sqref').length === 1);
  const all = [...cells];
  if (el) {
    for (const part of (attr(el, 'sqref') || '').split(/\s+/).filter(Boolean)) {
      const g = parseRange(part);
      if (g) for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) all.push({ r, c });
    }
  } else {
    el = make('ignoredError');
    list.appendChild(el);
  }
  const seen = new Set();
  el.setAttribute('sqref', sqrefOf(all.filter(({ r, c }) => !seen.has(`${r},${c}`) && seen.add(`${r},${c}`))));
  el.setAttribute(rule, '1');
  wb.dirty.add(wb.sheets[si].path);
  wb.edited = true;
}

/** Convert to Number: the changes [{ r, c, value }] that make text cells that read as numbers numbers. */
export function toNumberChanges(wb, si, cells) {
  const out = [];
  for (const { r, c } of cells) {
    const cell = wb.sheet(si).cells.get(r)?.get(c);
    const n = cell && cell.f == null && typeof cell.v === 'string' ? textToNumber(cell.v) : null;
    if (n != null) out.push({ r, c, value: n });
  }
  return out;
}

/** Copy Formula from Above (dr 1) or Left (dc 1): the change that gives cell (r, c) its neighbour's formula, moved. */
export function copyFormulaChange(wb, si, r, c, dr, dc) {
  const from = wb.sheet(si).cells.get(r - dr)?.get(c - dc);
  if (!from?.f) return null;
  return { r, c, formula: shiftFormula(from.f, dr, dc) };
}

/** Which way an inconsistent formula's neighbours run: { dr: 1, dc: 0 } above and below, { dr: 0, dc: 1 } left and
 *  right, as errorcheck.js found them. */
export function inconsistentFrom(wb, si, r, c) {
  const cells = wb.sheet(si).cells;
  const at = (rr, cc) => (rr >= 0 && cc >= 0 ? cells.get(rr)?.get(cc) : null);
  for (const [dr, dc] of [[1, 0], [0, 1]]) {
    const a = at(r - dr, c - dc);
    const b = at(r + dr, c + dc);
    if (a?.f && b?.f) { try { if (shiftFormula(a.f, dr * 2, dc * 2) === b.f) return { dr, dc }; } catch { /* (not this way) */ } }
  }
  return null;
}

