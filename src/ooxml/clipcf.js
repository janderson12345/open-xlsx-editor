// Conditional formats coming with a paste, as Excel's do. Measured over COM in a hidden Excel (copy, PasteSpecial or
// Cut, then the conditionalFormatting Excel saved; docs/CLIPBOARD.md):
// - Paste, All, Formats, All except borders, All using Source theme and All merging conditional formats bring them,
//   with Skip blanks or an operation too; Values, Formulas, the number-format pairs and Validation don't.
// - The pasted cells first leave every rule over them (Excel's All took E1:E3 out of E1:E4's rule, E4 left); All
//   merging conditional formats leaves those rules as they are.
// - Each conditionalFormatting the copy meets comes as a new one over the pasted cells, all its rules, their formats
//   copied, their formulas moved as a copied formula moves (=A1>B1 over A1:A3, pasted at C1: =C1>D1; with Transpose,
//   C1:E1 and the same =C1>D1), first in priority (the sheet's others one later each, in their order). Data bars and
//   Excel 2010's icon sets with their extLst half too, new ids (clipcfx.js).
// - On the sheet copied from, a rule over a single row or column, pasted along that line against or over it, takes the
//   cells in instead (A1:A3 pasted at A4: A1:A6; at A2: A1:A4; A3:A4 at A1 of A3:A5: A1:A5; A1:C1 at D1: A1:F1), its
//   formulas written for its new first cell. Anywhere else a copy of its own (A1:A3 at B1, at A5; A1:B2 at A3 or C1).
// - A cut moves the rules over the cells cut, their formulas as a cut formula's (=A1>B1 cut from A1:A3 to C1:
//   =C1>B1). A rule only partly cut keeps its cells left (=A1>B1 over A1:A3, A1:A2 cut to C1: =A3>B3 over A3) and the
//   moved ones get a rule of their own after it (=C1>B1 over C1:C2); with none left (A1:A2 cut to A2, over A3) the
//   rule moves (=A2>B1 over A2:A3). To another sheet, the cut cells leave their rules and the copy's come there as a
//   paste's.
import { kids, parseXml, serializeXml } from './xml.js';
import { parseRange, rangeName } from './refs.js';
import { shiftFormula, moveRefs } from '../calc/tokens.js';
import { clearConditional, AFTER_CF } from './cfedit.js';
import { readDxf } from './styles.js';
import { coalesce, subtract } from './dvedit.js';
import { x14Rules } from './cfbarsedit.js';
import { newId, x14Rule, baseX14Id, setX14Ranges, x14Blocks, addX14Clone } from './clipcfx.js';
import { maxOf, minOf } from './listmath.js';

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const clipTo = (a, b) => (overlaps(a, b) ? { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) } : null);
const shifted = (a, dr, dc) => ({ r0: a.r0 + dr, c0: a.c0 + dc, r1: a.r1 + dr, c1: a.c1 + dc });
const rangesOf = (cf) => (cf.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
const blocksOf = (doc) => kids(doc.documentElement, 'conditionalFormatting', doc.documentElement.namespaceURI);
const dxfEls = (wb) => kids(kids(wb.stylesDoc.documentElement, 'dxfs', wb.stylesDoc.documentElement.namespaceURI)[0], 'dxf', wb.stylesDoc.documentElement.namespaceURI);
const prioOf = (r) => Number(r.getAttribute('priority') || 0);

/** The conditional formats over `rect` of sheet si, for a copy: [{ ranges, parts, priorities, rules, x14only }], each
 *  block the copy meets (`parts`: its ranges within the copy); a base block's rules as { xml, dxf, x14 } (its format's
 *  XML, its extLst half as an element), an Excel 2010 icon set's as { x14 }. */
export function copyConditionals(wb, si, rect) {
  const doc = wb._part(si).doc;
  const dxfs = dxfEls(wb);
  const out = [];
  for (const cf of blocksOf(doc)) {
    const ranges = rangesOf(cf);
    const parts = ranges.map((a) => clipTo(a, rect)).filter(Boolean);
    const rules = kids(cf, 'cfRule', cf.namespaceURI);
    if (!parts.length || !rules.length) continue;
    out.push({
      ranges, parts, priorities: rules.map(prioOf), x14only: false,
      rules: rules.map((r) => {
        const dxf = r.hasAttribute('dxfId') ? dxfs[Number(r.getAttribute('dxfId'))] : null;
        const id = baseX14Id(r);
        return { xml: serializeXml(r), dxf: dxf ? serializeXml(dxf) : null, x14: id ? x14Rule(doc, id)?.cloneNode(true) ?? null : null };
      }),
    });
  }
  for (const b of x14Blocks(doc, parseRange)) {
    const parts = b.ranges.map((a) => clipTo(a, rect)).filter(Boolean);
    if (parts.length) out.push({ ranges: b.ranges, parts, priorities: b.rules.map(prioOf), x14only: true, rules: b.rules.map((r) => ({ x14: r.cloneNode(true) })) });
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
  let dxfs = kids(sroot, 'dxfs', sns)[0];
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
  if (transpose) return parts.map((p) => ({ r0: area.r0 + (p.c0 - rect.c0), c0: area.c0 + (p.r0 - rect.r0), r1: area.r0 + (p.c1 - rect.c0), c1: area.c0 + (p.r1 - rect.r0) }));
  const h = rect.r1 - rect.r0 + 1;
  const w = rect.c1 - rect.c0 + 1;
  const out = [];
  for (let tr = area.r0; tr <= area.r1; tr += h) {
    for (let tc = area.c0; tc <= area.c1; tc += w) {
      for (const p of parts) {
        const moved = clipTo(shifted(p, tr - rect.r0, tc - rect.c0), area);
        if (moved) out.push(moved);
      }
    }
  }
  return out;
}

/** A single-row or single-column range `a` and the ranges `to` making one line with it, against or over it: their
 *  union, else null. */
function alongLine(a, to) {
  const all = [a, ...to];
  const row = a.r0 === a.r1 && all.every((x) => x.r0 === a.r0 && x.r1 === a.r0);
  const col = a.c0 === a.c1 && all.every((x) => x.c0 === a.c0 && x.c1 === a.c0);
  if (!row && !col) return null;
  const covered = (row ? all.map((x) => [x.c0, x.c1]) : all.map((x) => [x.r0, x.r1])).sort((x, y) => x[0] - y[0]);
  let reach = covered[0][1];
  for (const [s, e] of covered.slice(1)) { if (s > reach + 1) return null; reach = Math.max(reach, e); }
  return { r0: minOf(all.map((x) => x.r0)), c0: minOf(all.map((x) => x.c0)), r1: maxOf(all.map((x) => x.r1)), c1: maxOf(all.map((x) => x.c1)) };
}

/** Every rule of the sheet `n` later in priority, base and extLst. */
function makeRoom(doc, n) {
  const ns = doc.documentElement.namespaceURI;
  for (const r of [...doc.getElementsByTagNameNS(ns, 'cfRule'), ...x14Rules(doc).filter((x) => x.hasAttribute('priority'))]) r.setAttribute('priority', String(prioOf(r) + n));
}

/** A new conditionalFormatting over `ranges` with copies of `rules` (copyConditionals'), their priorities from
 *  `prio(ri)`, their formulas moved by `fix`, before the elements that come after them. */
function addBlock(wb, doc, ranges, rules, prio, fix) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const cf = doc.createElementNS(ns, root.prefix ? `${root.prefix}:conditionalFormatting` : 'conditionalFormatting');
  cf.setAttribute('sqref', ranges.map(rangeName).join(' '));
  rules.forEach((x, ri) => {
    const rule = doc.importNode(parseXml(x.xml).documentElement, true);
    rule.setAttribute('priority', String(prio(ri)));
    if (x.dxf) rule.setAttribute('dxfId', String(addDxf(wb, x.dxf)));
    mapFormulas(rule, fix);
    // (A data bar's extLst half, under a new id named in both.)
    if (x.x14) {
      const id = newId();
      for (const e of rule.getElementsByTagNameNS('*', 'id')) e.textContent = id;
      addX14Clone(doc, x.x14, ranges, id);
    }
    cf.appendChild(rule);
  });
  root.insertBefore(cf, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_CF.includes(n.localName)) || null);
}

/** A block's cells changed to `ranges`, its extLst halves too. */
function setBlockRanges(doc, cf, ranges) {
  cf.setAttribute('sqref', ranges.map(rangeName).join(' '));
  for (const r of kids(cf, 'cfRule', cf.namespaceURI)) { const id = baseX14Id(r); if (id) setX14Ranges(doc, id, ranges); }
}

/** The copy's conditional formats pasted at `area` of sheet si (`merge`: All merging conditional formats, the rules
 *  there kept; `transpose`). Changes the sheet's XML and the styles; returns whether anything changed. */
export function pasteConditionals(wb, si, area, clip, { merge = false, transpose = false } = {}) {
  const blocks = clip.conditionals || [];
  const doc = wb._part(si).doc;
  const ns = doc.documentElement.namespaceURI;
  const had = blocksOf(doc).some((cf) => rangesOf(cf).some((a) => overlaps(a, area))) || x14Blocks(doc, parseRange).some((b) => b.ranges.some((a) => overlaps(a, area)));
  if (!blocks.length && (merge || !had)) return false;
  wb.changeConditional(si, () => {
    if (!merge && had) clearConditional(wb, si, [area]);
    const sameSheet = clip.wb === wb && clip.si === si && !transpose;
    const fresh = [];
    for (const b of blocks) {
      const to = landing(b.parts, clip.rect, area, transpose);
      if (!to.length) continue;
      // (On the sheet copied from, a line of cells takes the pasted ones in.)
      const key = b.priorities.join();
      const own = !sameSheet ? null : b.x14only
        ? x14Blocks(doc, parseRange).find((x) => x.rules.map(prioOf).join() === key)
        : blocksOf(doc).map((cf) => ({ cf, ranges: rangesOf(cf) })).find((x) => kids(x.cf, 'cfRule', ns).map(prioOf).join() === key);
      const line = own && own.ranges.length === 1 ? alongLine(own.ranges[0], to) : null;
      if (line && b.x14only) { for (const r of own.rules) setX14Ranges(doc, r.getAttribute('id'), [line]); continue; }
      if (line) {
        setBlockRanges(doc, own.cf, [line]);
        for (const r of kids(own.cf, 'cfRule', ns)) mapFormulas(r, (f) => shiftFormula(f, line.r0 - own.ranges[0].r0, line.c0 - own.ranges[0].c0));
        continue;
      }
      fresh.push({ b, to: coalesce(to) });
    }
    if (!fresh.length) return;
    // (The pasted rules first, in their order; every other rule of the sheet that many later.)
    makeRoom(doc, fresh.reduce((k, x) => k + x.b.rules.length, 0));
    const order = fresh.flatMap(({ b }, bi) => b.priorities.map((p, ri) => ({ p, bi, ri }))).sort((x, y) => x.p - y.p);
    const prio = new Map(order.map((x, k) => [`${x.bi},${x.ri}`, k + 1]));
    fresh.forEach(({ b, to }, bi) => {
      if (b.x14only) { b.rules.forEach((x, ri) => addX14Clone(doc, x.x14, to, newId(), prio.get(`${bi},${ri}`))); return; }
      // (Formulas written for the rule's first range's top-left cell: moved from the old one to the new one.)
      const base = b.ranges[0];
      addBlock(wb, doc, to, b.rules, (ri) => prio.get(`${bi},${ri}`), (f) => shiftFormula(f, to[0].r0 - base.r0, to[0].c0 - base.c0));
    });
  });
  return true;
}

/** A cut's conditional formats, after its cells were pasted at `area` of sheet si: on the same sheet, each rule's cut
 *  cells move, its formulas as a cut formula's, and the pasted cells leave the others (the rules above); to another
 *  sheet, the cut cells leave their rules and the copy's come as a paste's. Returns whether anything changed. */
export function moveCutConditionals(wb, si, area, clip) {
  if (clip.si !== si || clip.wb !== wb) {
    const left = (clip.conditionals || []).length > 0;
    if (left) wb.changeConditional(clip.si, () => clearConditional(wb, clip.si, [clip.rect]));
    return pasteConditionals(wb, si, area, clip) || left;
  }
  const doc = wb._part(si).doc;
  const ns = doc.documentElement.namespaceURI;
  const dr = area.r0 - clip.rect.r0;
  const dc = area.c0 - clip.rect.c0;
  const name = wb.sheets[si].name;
  const move = (f) => moveRefs(f, { here: name, from: { sheet: name, ...clip.rect }, to: { sheet: name, dr, dc } });
  const cut = blocksOf(doc).map((cf) => ({ cf, ranges: rangesOf(cf) })).filter((x) => x.ranges.some((a) => overlaps(a, clip.rect)));
  const cutX = x14Blocks(doc, parseRange).filter((x) => x.ranges.some((a) => overlaps(a, clip.rect)));
  const others = blocksOf(doc).some((cf) => !cut.some((x) => x.cf === cf) && rangesOf(cf).some((a) => overlaps(a, area)));
  if (!cut.length && !cutX.length && !others) return false;
  wb.changeConditional(si, () => {
    // (The cut rules set aside while the pasted cells leave the rest.)
    for (const x of cut) x.cf.parentNode.removeChild(x.cf);
    const xs = cutX.map((x) => ({ ...x, parent: x.cf.parentNode, ids: x.rules.map((r) => r.getAttribute('id')) }));
    for (const x of xs) x.parent.removeChild(x.cf);
    if (others) clearConditional(wb, si, [area]);
    const at = [...doc.documentElement.childNodes].find((n) => n.nodeType === 1 && AFTER_CF.includes(n.localName)) || null;
    const last = Math.max(0, ...[...doc.getElementsByTagNameNS(ns, 'cfRule'), ...x14Rules(doc)].map(prioOf), ...cut.flatMap((x) => kids(x.cf, 'cfRule', ns).map(prioOf)));
    let next = last;
    for (const { cf, ranges } of cut) {
      const movedFrom = ranges.map((a) => clipTo(a, clip.rect)).filter(Boolean);
      const moved = movedFrom.map((a) => shifted(a, dr, dc));
      let left = ranges.flatMap((a) => subtract(a, clip.rect)).flatMap((a) => subtract(a, area));
      const anchor = ranges[0];
      const ruleEls = kids(cf, 'cfRule', ns);
      // (Part of a data bar: its rule stays as it was, not measured.)
      if (left.length && ruleEls.some(baseX14Id)) { doc.documentElement.insertBefore(cf, at); continue; }
      if (!left.length) {
        // (None of its cells left: it moves, its formulas as a cut formula's, from its first cut cell.)
        for (const r of ruleEls) mapFormulas(r, (f) => move(shiftFormula(f, movedFrom[0].r0 - anchor.r0, movedFrom[0].c0 - anchor.c0)));
        setBlockRanges(doc, cf, coalesce(moved));
        doc.documentElement.insertBefore(cf, at);
        continue;
      }
      // (Cells left: the rule keeps them, written for its new first cell; the moved ones a rule of their own after.)
      left = coalesce(left);
      const copy = cf.cloneNode(true);
      cf.setAttribute('sqref', left.map(rangeName).join(' '));
      for (const r of ruleEls) mapFormulas(r, (f) => shiftFormula(f, left[0].r0 - anchor.r0, left[0].c0 - anchor.c0));
      doc.documentElement.insertBefore(cf, at);
      copy.setAttribute('sqref', coalesce(moved).map(rangeName).join(' '));
      for (const r of kids(copy, 'cfRule', ns)) { r.setAttribute('priority', String(++next)); mapFormulas(r, (f) => move(shiftFormula(f, movedFrom[0].r0 - anchor.r0, movedFrom[0].c0 - anchor.c0))); }
      doc.documentElement.insertBefore(copy, at);
    }
    // (Excel 2010's own icon sets: moved when wholly cut, else left as they were.)
    for (const x of xs) {
      x.parent.appendChild(x.cf);
      if (x.ranges.every((a) => a.r0 >= clip.rect.r0 && a.r1 <= clip.rect.r1 && a.c0 >= clip.rect.c0 && a.c1 <= clip.rect.c1)) for (const id of x.ids) setX14Ranges(doc, id, x.ranges.map((a) => shifted(a, dr, dc)));
    }
  });
  return true;
}
