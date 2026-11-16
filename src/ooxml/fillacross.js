// Home > Fill > Across Worksheets, as Excel's Sheets(...).FillAcrossSheets does it (measured over COM:
// tools/excel-compare/excel-fillacross.ps1, test/fillacross-excel.json, docs/FILL.md). The selection on the sheet
// shown is copied to the same cells of each other grouped sheet: All (contents and formats), Contents or Formats.
// Formulas are copied as they are written (the cells are the same), a blank cell clears the one it lands on; All and
// Formats bring the cell formats and the merged ranges, and the conditional formats; All brings hyperlinks and data
// validation too. Notes, column widths and row heights stay as they were on each sheet.
import { snapshot, pasteValidation, pasteConditionals } from './clipops.js';
import { copyHyperlinks } from './links.js';

/** Excel's three choices, by the names its box gives them. */
export const FILL_ACROSS = ['all', 'contents', 'formats'];

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const inside = (a, b) => a.r0 >= b.r0 && a.r1 <= b.r1 && a.c0 >= b.c0 && a.c1 <= b.c1;

/** `sel` cut to where any of the sheets has something (a whole column selected isn't a million cells). */
function areaOf(wb, sheets, sel) {
  const maxRow = Math.max(0, ...sheets.map((k) => wb.sheet(k).maxRow));
  const maxCol = Math.max(0, ...sheets.map((k) => wb.sheet(k).maxCol));
  return { r0: sel.r0, c0: sel.c0, r1: Math.max(sel.r0, Math.min(sel.r1, maxRow)), c1: Math.max(sel.c0, Math.min(sel.c1, maxCol)) };
}

/**
 * The cells' changes of a fill across worksheets: `sel` of sheet si to each sheet in `targets` (indices; si itself is
 * passed over), `type` 'all', 'contents' or 'formats'. Steps as clipops.js gives them: [{ si, contents, styles,
 * merges, area }], one for each other sheet.
 */
export function fillAcrossSteps(wb, si, sel, targets, type = 'all') {
  if (!FILL_ACROSS.includes(type)) throw new Error(`not a way to fill across worksheets: ${type}`);
  const others = [...new Set(targets)].filter((k) => k !== si && wb.sheets[k]?.kind === 'worksheet');
  const area = areaOf(wb, [si, ...others], sel);
  const src = wb.sheet(si);
  const contents = type !== 'formats';
  const formats = type !== 'contents';
  return others.map((k) => {
    const dst = wb.sheet(k);
    const out = { si: k, contents: [], styles: [], merges: dst.merges, area };
    for (let r = area.r0; r <= area.r1; r++) {
      for (let c = area.c0; c <= area.c1; c++) {
        const from = src.cells.get(r)?.get(c);
        const was = dst.cells.get(r)?.get(c);
        if (!from && !was) continue;
        if (contents) {
          if (from?.f != null && from.f !== '') out.contents.push({ r, c, formula: from.f });
          else if (from?.v != null || was?.v != null || was?.f != null) out.contents.push({ r, c, value: from?.t === 'e' ? { code: String(from.v) } : from?.v ?? null });
        }
        if (formats && (from?.s || 0) !== (was?.s || 0)) out.styles.push({ r, c, s: from?.s || 0 });
      }
    }
    // (The merged ranges in the area as the shown sheet has them, measured; those that met it there go.)
    if (formats) out.merges = [...dst.merges.filter((m) => !overlaps(m, area)), ...src.merges.filter((m) => inside(m, area)).map((m) => ({ ...m }))];
    return out;
  });
}

/**
 * The rest of a fill across worksheets, after its steps: All's hyperlinks and data validation, All's and Formats'
 * conditional formats, put over the same cells of each other sheet (as a paste puts them: clipops.js). Changes the
 * workbook; returns whether anything changed.
 */
export function fillAcrossParts(wb, si, sel, targets, type = 'all') {
  if (type === 'contents') return false;
  const others = [...new Set(targets)].filter((k) => k !== si && wb.sheets[k]?.kind === 'worksheet');
  const area = areaOf(wb, [si, ...others], sel);
  const clip = snapshot(wb, si, area);
  // (The snapshot cut to the shown sheet's cells: the other sheets' area is the whole one.)
  clip.rect = area;
  let changed = false;
  for (const k of others) {
    if (type === 'all') {
      changed = copyHyperlinks(wb, si, k, area) || changed;
      changed = pasteValidation(wb, k, area, clip) || changed;
    }
    changed = pasteConditionals(wb, k, area, clip, { merge: false }) || changed;
  }
  return changed;
}

/** Both, on a workbook (the API's, and the tests'): the steps made and the parts put over. */
export function fillAcrossSheets(wb, si, sel, targets, type = 'all') {
  const steps = fillAcrossSteps(wb, si, sel, targets, type);
  for (const st of steps) {
    if (st.contents.length) wb.apply(st.si, st.contents);
    if (st.merges !== wb.sheet(st.si).merges) wb.setMerges(st.si, st.merges);
    if (st.styles.length) wb.setStyles(st.si, st.styles);
  }
  fillAcrossParts(wb, si, sel, targets, type);
  return steps;
}
