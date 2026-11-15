// Conditional formats coming with a paste, as Excel's do. Measured over COM in a hidden Excel (copy, PasteSpecial or
// Cut, then the conditionalFormatting Excel saved; docs/CLIPBOARD.md):
// - Paste, All, Formats, All except borders, All using Source theme and All merging conditional formats bring them,
//   with Skip blanks or an operation too; Values, Formulas, the number-format pairs and Validation don't.
// - The pasted cells first leave every rule over them (Excel's All took E1:E3 out of E1:E4's rule, E4 left); All
//   merging conditional formats leaves those rules as they are.
// - Each conditionalFormatting the copy meets comes as a new one over the pasted cells, all its rules, their formats
//   copied, their formulas moved as a copied formula moves (=A1>B1 over A1:A3, pasted at C1: =C1>D1; with Transpose,
//   C1:E1 and the same =C1>D1), first in priority (the sheet's others one later each, in their order).
// - On the sheet copied from, a rule over a single row or column, pasted along that line against or over it, takes the
//   cells in instead (A1:A3 pasted at A4: A1:A6; at A2: A1:A4; A3:A4 at A1 of A3:A5: A1:A5; A1:C1 at D1: A1:F1), its
//   formulas written for its new first cell. Anywhere else a copy of its own (A1:A3 at B1, at A5; A1:B2 at A3 or C1).
// - A cut moves the rules over the cells cut, their formulas as a cut formula's (=A1>B1 cut from A1:A3 to C1:
//   =C1>B1). To another sheet, the cut cells leave their rules and the copy's come there as a paste's.
// Rules with Excel 2010's settings in the extLst (data bars, its icon sets) aren't brought yet.
import { kids, kid, parseXml, serializeXml } from './xml.js';
import { parseRange, rangeName } from './refs.js';
import { shiftFormula, moveRefs } from '../calc/tokens.js';
import { clearConditional, AFTER_CF } from './cfedit.js';
import { readDxf } from './styles.js';
import { coalesce } from './dvedit.js';

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const clipTo = (a, b) => (overlaps(a, b) ? { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) } : null);
const inside = (a, b) => a.r0 >= b.r0 && a.r1 <= b.r1 && a.c0 >= b.c0 && a.c1 <= b.c1;
const rangesOf = (cf) => (cf.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
const blocksOf = (doc) => kids(doc.documentElement, 'conditionalFormatting', doc.documentElement.namespaceURI);
const hasExt = (rule) => !!kid(rule, 'extLst', rule.namespaceURI);
const dxfEls = (wb) => kids(kid(wb.stylesDoc.documentElement, 'dxfs', wb.stylesDoc.documentElement.namespaceURI), 'dxf', wb.stylesDoc.documentElement.namespaceURI);

/** The conditional formats over `rect` of sheet si, for a copy: [{ ranges, parts, priorities, rules: [{ xml, dxf }] }],
 *  each block the copy meets (`parts`: its ranges within the copy), its rules as XML and their formats' XML. */
export function copyConditionals(wb, si, rect) {
  const doc = wb._part(si).doc;
  const dxfs = dxfEls(wb);
  const out = [];
  for (const cf of blocksOf(doc)) {
    const ranges = rangesOf(cf);
    const parts = ranges.map((a) => clipTo(a, rect)).filter(Boolean);
    const rules = kids(cf, 'cfRule', cf.namespaceURI);
    if (!parts.length || !rules.length || rules.some(hasExt)) continue;
    out.push({
      ranges, parts,
      priorities: rules.map((r) => Number(r.getAttribute('priority') || 0)),
      rules: rules.map((r) => ({ xml: serializeXml(r), dxf: r.hasAttribute('dxfId') ? (dxfs[Number(r.getAttribute('dxfId'))] ? serializeXml(dxfs[Number(r.getAttribute('dxfId'))]) : null) : null })),
    });
  }
  return out;
}

/** A rule's formulas rewritten by `fn`. */
function mapFormulas(rule, fn) {
  for (const f of kids(rule, 'formula', rule.namespaceURI)) f.textContent = fn(f.textContent);
}

/** A copied format added to the workbook's dxfs: its new id. */
function addDxf(wb, xml) {
  const sroot = wb.stylesDoc.documentElement;
  const sns = sroot.namespaceURI;
  let dxfs = kid(sroot, 'dxfs', sns);
  if (!dxfs) {
    dxfs = wb.stylesDoc.createElementNS(sns, sroot.prefix ? `${sroot.prefix}:dxfs` : 'dxfs');
    const after = [...sroot.childNodes].find((n) => n.nodeType === 1 && ['tableStyles', 'colors', 'extLst'].includes(n.localName));
    sroot.insertBefore(dxfs, after || null);
  }
  const el = wb.stylesDoc.importNode(parseXml(xml).documentElement, true);
  dxfs.appendChild(el);
  const id = kids(dxfs, 'dxf', sns).length - 1;
  dxfs.setAttribute('count', String(id + 1));
  wb.styles.dxfs.push(readDxf(el, sns));
  wb.dirty.add(wb.stylesPath);
  return id;
}

/** Where each part of the copy lands in `area` (repeated over it as the cells are; with Transpose turned, once). */
function landing(parts, rect, area, transpose) {
  if (transpose) {
    return parts.map((p) => ({ r0: area.r0 + (p.c0 - rect.c0), c0: area.c0 + (p.r0 - rect.r0), r1: area.r0 + (p.c1 - rect.c0), c1: area.c0 + (p.r1 - rect.r0) }));
  }
  const h = rect.r1 - rect.r0 + 1;
  const w = rect.c1 - rect.c0 + 1;
  const out = [];
  for (let tr = area.r0; tr <= area.r1; tr += h) {
    for (let tc = area.c0; tc <= area.c1; tc += w) {
      for (const p of parts) {
        const moved = clipTo({ r0: p.r0 + tr - rect.r0, c0: p.c0 + tc - rect.c0, r1: p.r1 + tr - rect.r0, c1: p.c1 + tc - rect.c0 }, area);
        if (moved) out.push(moved);
      }
    }
  }
  return out;
}

/** The block on the sheet copied from that has the rules of priorities `ps`, now. */
const blockWith = (doc, ps) => blocksOf(doc).find((cf) => kids(cf, 'cfRule', cf.namespaceURI).map((r) => Number(r.getAttribute('priority') || 0)).join() === ps.join());

/** A single-row or single-column range `a` and the ranges `to` making one line with it, against or over it: their
 *  union, else null. */
function alongLine(a, to) {
  const all = [a, ...to];
  const r0 = Math.min(...all.map((x) => x.r0));
  const r1 = Math.max(...all.map((x) => x.r1));
  const c0 = Math.min(...all.map((x) => x.c0));
  const c1 = Math.max(...all.map((x) => x.c1));
  const row = a.r0 === a.r1 && all.every((x) => x.r0 === a.r0 && x.r1 === a.r0);
  const col = a.c0 === a.c1 && all.every((x) => x.c0 === a.c0 && x.c1 === a.c0);
  if (!row && !col) return null;
  // (No gap: the cells covered, end to end.)
  const covered = row ? all.map((x) => [x.c0, x.c1]) : all.map((x) => [x.r0, x.r1]);
  covered.sort((x, y) => x[0] - y[0]);
  let reach = covered[0][1];
  for (const [s, e] of covered.slice(1)) { if (s > reach + 1) return null; reach = Math.max(reach, e); }
  return { r0, c0, r1, c1 };
}

/** The copy's conditional formats pasted at `area` of sheet si (`merge`: All merging conditional formats, the rules
 *  there kept; `transpose`). Changes the sheet's XML and the styles; returns whether anything changed. */
export function pasteConditionals(wb, si, area, clip, { merge = false, transpose = false } = {}) {
  const blocks = clip.conditionals || [];
  const doc = wb._part(si).doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const had = blocksOf(doc).some((cf) => rangesOf(cf).some((a) => overlaps(a, area)));
  if (!blocks.length && (merge || !had)) return false;
  wb.changeConditional(si, () => {
    if (!merge && had) clearConditional(wb, si, [area]);
    const sameSheet = clip.wb === wb && clip.si === si;
    const fresh = [];
    for (const b of blocks) {
      const to = landing(b.parts, clip.rect, area, transpose);
      if (!to.length) continue;
      // (On the sheet copied from, a line of cells takes the pasted ones in.)
      const own = sameSheet && !transpose ? blockWith(doc, b.priorities) : null;
      const now = own ? rangesOf(own) : [];
      const line = now.length === 1 ? alongLine(now[0], to) : null;
      if (line) {
        own.setAttribute('sqref', rangeName(line));
        for (const r of kids(own, 'cfRule', ns)) mapFormulas(r, (f) => shiftFormula(f, line.r0 - now[0].r0, line.c0 - now[0].c0));
        continue;
      }
      fresh.push({ b, to: coalesce(to) });
    }
    if (!fresh.length) return;
    // (The pasted rules first, in their order; every other rule of the sheet that many later.)
    const n = fresh.reduce((k, x) => k + x.b.rules.length, 0);
    for (const r of doc.getElementsByTagNameNS(ns, 'cfRule')) r.setAttribute('priority', String(Number(r.getAttribute('priority') || 0) + n));
    const order = fresh.flatMap(({ b }, bi) => b.priorities.map((p, ri) => ({ p, bi, ri }))).sort((x, y) => x.p - y.p);
    const prio = new Map(order.map((x, k) => [`${x.bi},${x.ri}`, k + 1]));
    const at = [...root.childNodes].find((x) => x.nodeType === 1 && AFTER_CF.includes(x.localName)) || null;
    fresh.forEach(({ b, to }, bi) => {
      const cf = doc.createElementNS(ns, root.prefix ? `${root.prefix}:conditionalFormatting` : 'conditionalFormatting');
      cf.setAttribute('sqref', to.map(rangeName).join(' '));
      // (Formulas written for the rule's first range's top-left cell: moved from the old one to the new one.)
      const base = b.ranges[0];
      const dr = to[0].r0 - base.r0;
      const dc = to[0].c0 - base.c0;
      b.rules.forEach((x, ri) => {
        const rule = doc.importNode(parseXml(x.xml).documentElement, true);
        rule.setAttribute('priority', String(prio.get(`${bi},${ri}`)));
        if (x.dxf) rule.setAttribute('dxfId', String(addDxf(wb, x.dxf)));
        mapFormulas(rule, (f) => shiftFormula(f, dr, dc));
        cf.appendChild(rule);
      });
      root.insertBefore(cf, at);
    });
  });
  return true;
}

/** A cut's conditional formats, after its cells were pasted at `area` of sheet si: on the same sheet, the rules wholly
 *  over the cells cut move with them, their formulas as a cut formula's, the pasted cells leaving the others; to
 *  another sheet, the cut cells leave their rules and the copy's come as a paste's. Returns whether anything changed. */
export function moveCutConditionals(wb, si, area, clip) {
  if (clip.si !== si || clip.wb !== wb) {
    const left = (clip.conditionals || []).length > 0;
    if (left) wb.changeConditional(clip.si, () => clearConditional(wb, clip.si, [clip.rect]));
    return pasteConditionals(wb, si, area, clip) || left;
  }
  const doc = wb._part(si).doc;
  const blocks = blocksOf(doc);
  const moving = blocks.filter((cf) => { const rs = rangesOf(cf); return rs.length && rs.every((a) => inside(a, clip.rect)); });
  const others = blocks.filter((cf) => !moving.includes(cf) && rangesOf(cf).some((a) => overlaps(a, area)));
  if (!moving.length && !others.length) return false;
  const dr = area.r0 - clip.rect.r0;
  const dc = area.c0 - clip.rect.c0;
  const name = wb.sheets[si].name;
  wb.changeConditional(si, () => {
    // (The pasted cells leave the rules that stay: their blocks set aside while the cleared ones are trimmed.)
    for (const cf of moving) cf.parentNode.removeChild(cf);
    if (others.length) clearConditional(wb, si, [area]);
    const at = [...doc.documentElement.childNodes].find((x) => x.nodeType === 1 && AFTER_CF.includes(x.localName)) || null;
    for (const cf of moving) {
      cf.setAttribute('sqref', rangesOf(cf).map((a) => rangeName({ r0: a.r0 + dr, c0: a.c0 + dc, r1: a.r1 + dr, c1: a.c1 + dc })).join(' '));
      for (const r of kids(cf, 'cfRule', cf.namespaceURI)) mapFormulas(r, (f) => moveRefs(f, { here: name, from: { sheet: name, ...clip.rect }, to: { sheet: name, dr, dc } }));
      doc.documentElement.insertBefore(cf, at);
    }
  });
  return true;
}
