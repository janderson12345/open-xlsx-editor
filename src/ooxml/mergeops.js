// Merge & Center, Merge Across, Merge Cells and Unmerge Cells, as Excel does them (measured over COM, docs/FORMATS.md).
// Each returns what changes, for the editor to apply and undo as one step: { contents: [{ r, c, value } or
// { r, c, formula }], styles: [{ r, c, s }], merges: the sheet's merged ranges afterwards }.
import { cellsOf } from './formatops.js';

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const hasContent = (cell) => !!cell && (cell.v != null || (cell.f != null && cell.f !== ''));

/** `area` grown until it takes in every merged range it touches (Excel merges them into one). */
function grow(area, merges) {
  let a = { ...area };
  for (let again = true; again;) {
    again = false;
    for (const m of merges) {
      if (!overlaps(a, m)) continue;
      const b = { r0: Math.min(a.r0, m.r0), c0: Math.min(a.c0, m.c0), r1: Math.max(a.r1, m.r1), c1: Math.max(a.c1, m.c1) };
      if (b.r0 !== a.r0 || b.c0 !== a.c0 || b.r1 !== a.r1 || b.c1 !== a.c1) { a = b; again = true; }
    }
  }
  return a;
}

/** The areas a merge makes (each grown to take in the merged ranges it touches), or null when the command unmerges. */
export function mergeAreas(sheet, sel, kind) {
  if (kind === 'unmerge' || (kind === 'center' && sheet.merges.some((m) => overlaps(m, sel)))) return null;
  const rows = kind === 'across' ? Array.from({ length: sel.r1 - sel.r0 + 1 }, (_, k) => ({ r0: sel.r0 + k, c0: sel.c0, r1: sel.r0 + k, c1: sel.c1 })) : [sel];
  const areas = [];
  let merges = sheet.merges;
  for (const row of rows) {
    const area = grow(row, merges);
    if (area.r0 === area.r1 && area.c0 === area.c1) continue;
    merges = [...merges.filter((m) => !overlaps(m, area)), area];
    areas.push(area);
  }
  return areas;
}

/** How many values a merge would discard (Excel asks first when there are any). */
export function mergeLoss(sheet, sel, kind) {
  const areas = mergeAreas(sheet, sel, kind) || [];
  return areas.reduce((n, a) => n + Math.max(0, cellsOf(sheet, a).filter(({ r, c }) => hasContent(sheet.cells.get(r)?.get(c))).length - 1), 0);
}

/**
 * `kind`: 'center' (Merge & Center), 'across' (Merge Across: each row on its own), 'cells' (Merge Cells) or 'unmerge'.
 * Merging keeps one value, the first in reading order (moved into the top-left cell), and gives every cell of the
 * area that cell's format; Merge & Center centres it too. Merge & Center on a selection that has a merged range in it
 * unmerges instead, and sets the alignment back to General.
 */
export function mergeChange(wb, si, sel, kind) {
  const sheet = wb.sheet(si);
  const touched = sheet.merges.filter((m) => overlaps(m, sel));
  const out = { contents: [], styles: [], merges: sheet.merges.filter((m) => !overlaps(m, sel)) };
  const areas = mergeAreas(sheet, sel, kind);
  if (!areas) {
    if (!touched.length) return null;
    if (kind === 'center') {
      for (const m of touched) out.styles.push(...cellsOf(sheet, m).map(({ r, c }) => ({ r, c, s: wb.styleWith(sheet.cells.get(r)?.get(c)?.s || 0, { align: { horizontal: 'general' } }) })));
    }
    return out;
  }
  if (!areas.length) return null;
  let merges = sheet.merges;
  for (const area of areas) {
    merges = merges.filter((m) => !overlaps(m, area));
    const cells = cellsOf(sheet, area);
    const at = (r, c) => sheet.cells.get(r)?.get(c);
    const kept = cells.find(({ r, c }) => hasContent(at(r, c))) || cells[0];
    const keptCell = at(kept.r, kept.c);
    const s = keptCell?.s || 0;
    const look = kind === 'center' ? wb.styleWith(s, { align: { horizontal: 'center' } }) : s;
    for (const { r, c } of cells) {
      const cell = at(r, c);
      const top = r === area.r0 && c === area.c0;
      if (top && (kept.r !== r || kept.c !== c)) out.contents.push(keptCell.f ? { r, c, formula: keptCell.f } : { r, c, value: keptCell.v });
      else if (!top && hasContent(cell)) out.contents.push({ r, c, value: null });
      out.styles.push({ r, c, s: look });
    }
    merges = [...merges, area];
  }
  out.merges = merges;
  return out;
}

/**
 * The merge over a selection of several areas: each area merged (or unmerged) as alone, all in one change, as Excel
 * did (Merge & Center over A2:B3 and D6:E7 merged both; Merge Across each one's rows). Areas that overlap: null, as
 * Excel left them (nothing merged, nothing said).
 */
export function mergeAreasChange(wb, si, areas, kind) {
  if (areas.length < 2) return mergeChange(wb, si, areas[0], kind);
  for (let i = 0; i < areas.length; i++) for (let j = i + 1; j < areas.length; j++) if (overlaps(areas[i], areas[j])) return null;
  const sheet = wb.sheet(si);
  const out = { contents: [], styles: [], merges: sheet.merges.filter((m) => !areas.some((a) => overlaps(m, a))) };
  let any = false;
  for (const a of areas) {
    const one = mergeChange(wb, si, a, kind);
    if (!one) continue;
    any = true;
    out.contents.push(...one.contents);
    out.styles.push(...one.styles);
    // (Its own new merged ranges; the ones it touched are gone from out.merges already.)
    out.merges.push(...one.merges.filter((m) => overlaps(m, a)));
  }
  return any ? out : null;
}
