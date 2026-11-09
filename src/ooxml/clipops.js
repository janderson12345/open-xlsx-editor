// Copy, cut and paste, as Excel does them (measured over COM, docs/CLIPBOARD.md). A copy is a snapshot of cells:
// their values, formulas, formats and merged ranges. Pasting returns what changes, sheet by sheet, for the editor to
// apply and undo as one step: [{ si, contents: [{ r, c, value } or { r, c, formula }], styles: [{ r, c, s }], merges }].
import { shiftFormula, moveRefs } from '../calc/tokens.js';
import { parseInput } from './input.js';
import { cellsOf } from './formatops.js';

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const inside = (a, b) => a.r0 >= b.r0 && a.r1 <= b.r1 && a.c0 >= b.c0 && a.c1 <= b.c1;

/** A snapshot of the selected cells of sheet si (the selection as it is, cut to the used range when it runs past it). */
export function snapshot(wb, si, sel, cut = false) {
  const sheet = wb.sheet(si);
  const cells = cellsOf(sheet, sel);
  const rect = { r0: sel.r0, c0: sel.c0, r1: Math.max(...cells.map((x) => x.r)), c1: Math.max(...cells.map((x) => x.c)) };
  return {
    wb, si, rect, cut,
    sheetName: wb.sheets[si].name,
    cells: cells.map(({ r, c }) => {
      const cell = sheet.cells.get(r)?.get(c);
      return { dr: r - rect.r0, dc: c - rect.c0, f: cell?.f || null, v: cell?.f ? null : cell?.v ?? null, t: cell?.t, s: cell?.s || 0 };
    }),
    merges: sheet.merges.filter((m) => inside(m, rect)).map((m) => ({ r0: m.r0 - rect.r0, c0: m.c0 - rect.c0, r1: m.r1 - rect.r0, c1: m.c1 - rect.c0 })),
  };
}

/** Where a paste goes: the copy repeated over the selection when the selection is a whole number of copies in both
 *  directions, else one copy from the selection's top-left cell (Excel's rule). */
export function pasteArea(sel, rows, cols) {
  const h = sel.r1 - sel.r0 + 1;
  const w = sel.c1 - sel.c0 + 1;
  const tile = h % rows === 0 && w % cols === 0;
  return { r0: sel.r0, c0: sel.c0, r1: sel.r0 + (tile ? h : rows) - 1, c1: sel.c0 + (tile ? w : cols) - 1 };
}

const valueOf = (cell) => (cell.t === 'e' ? String(cell.v) : cell.v);

/** Paste a copy (not a cut) into sheet si at `sel`: values, formulas moved as their cells are (relative references
 *  follow, $ ones stay, one off the sheet is #REF!), formats and merged ranges. A copy from another workbook brings
 *  its values and formulas but not its formats. */
export function pasteCopy(wb, si, sel, clip) {
  const rows = clip.rect.r1 - clip.rect.r0 + 1;
  const cols = clip.rect.c1 - clip.rect.c0 + 1;
  const area = pasteArea(sel, rows, cols);
  const sameBook = clip.wb === wb;
  const out = { si, contents: [], styles: [], merges: [], area };
  for (let tr = area.r0; tr <= area.r1; tr += rows) {
    for (let tc = area.c0; tc <= area.c1; tc += cols) {
      for (const cell of clip.cells) {
        const r = tr + cell.dr;
        const c = tc + cell.dc;
        const dr = r - (clip.rect.r0 + cell.dr);
        const dc = c - (clip.rect.c0 + cell.dc);
        out.contents.push(cell.f ? { r, c, formula: shiftFormula(cell.f, dr, dc) } : { r, c, value: valueOf(cell) });
        if (sameBook) out.styles.push({ r, c, s: cell.s });
      }
      for (const m of clip.merges) out.merges.push({ r0: tr + m.r0, c0: tc + m.c0, r1: tr + m.r1, c1: tc + m.c1 });
    }
  }
  out.merges = [...wb.sheet(si).merges.filter((m) => !overlaps(m, area)), ...out.merges];
  return [out];
}

/**
 * Paste a cut: the cells move to `sel`'s top-left cell (the copy's size, whatever is selected), leaving their places
 * empty and unformatted, with their merged ranges. Every formula in the workbook that reads them follows them (a
 * reference wholly inside the cut cells); a reference to a single cell pasted over becomes #REF!. Only within the
 * workbook the cut was made in.
 */
export function pasteCut(wb, si, sel, clip) {
  const rows = clip.rect.r1 - clip.rect.r0 + 1;
  const cols = clip.rect.c1 - clip.rect.c0 + 1;
  const area = { r0: sel.r0, c0: sel.c0, r1: sel.r0 + rows - 1, c1: sel.c0 + cols - 1 };
  const dr = area.r0 - clip.rect.r0;
  const dc = area.c0 - clip.rect.c0;
  const from = { sheet: clip.sheetName, ...clip.rect };
  const to = { sheet: wb.sheets[si].name, dr, dc };
  const bySheet = new Map();
  const step = (k) => {
    if (!bySheet.has(k)) bySheet.set(k, { si: k, contents: new Map(), styles: new Map(), merges: wb.sheet(k).merges });
    return bySheet.get(k);
  };
  const key = (r, c) => `${r}:${c}`;
  // (Every other formula of the workbook that reads the cut cells, or a cell pasted over.)
  wb.sheets.forEach((info, k) => {
    if (info.kind !== 'worksheet' || !info.path || !wb.pkg.has(info.path)) return;
    const sheet = wb.sheet(k);
    for (const [r, line] of sheet.cells) {
      for (const [c, cell] of line) {
        if (!cell.f) continue;
        if (k === clip.si && r >= clip.rect.r0 && r <= clip.rect.r1 && c >= clip.rect.c0 && c <= clip.rect.c1) continue;
        const f = moveRefs(cell.f, { here: info.name, from, to });
        if (f !== cell.f) step(k).contents.set(key(r, c), { r, c, formula: f });
      }
    }
  });
  // The cut cells' places emptied and unformatted, then the cells written where they go.
  const src = step(clip.si);
  for (const cell of clip.cells) {
    const r = clip.rect.r0 + cell.dr;
    const c = clip.rect.c0 + cell.dc;
    src.contents.set(key(r, c), { r, c, value: null });
    src.styles.set(key(r, c), { r, c, s: 0 });
  }
  src.merges = src.merges.filter((m) => !inside(m, clip.rect));
  const dst = step(si);
  for (const cell of clip.cells) {
    const r = area.r0 + cell.dr;
    const c = area.c0 + cell.dc;
    const f = cell.f && moveRefs(cell.f, { here: clip.sheetName, home: to.sheet, from, to });
    dst.contents.set(key(r, c), f ? { r, c, formula: f } : { r, c, value: valueOf(cell) });
    dst.styles.set(key(r, c), { r, c, s: cell.s });
  }
  dst.merges = [...dst.merges.filter((m) => !overlaps(m, area)), ...clip.merges.map((m) => ({ r0: area.r0 + m.r0, c0: area.c0 + m.c0, r1: area.r0 + m.r1, c1: area.c0 + m.c1 }))];
  return [...bySheet.values()].map((x) => ({ si: x.si, contents: [...x.contents.values()], styles: [...x.styles.values()], merges: x.merges, area: x.si === si ? area : null }));
}

/** Text from outside (tabs between cells, a line for each row) pasted at `sel`'s top-left cell, each cell read as if it
 *  were typed there (a number, a date, =a formula...). Formats stay as they are. */
export function pasteText(wb, si, sel, text) {
  const lines = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
  const contents = [];
  lines.forEach((line, dr) => line.forEach((t, dc) => {
    const r = sel.r0 + dr;
    const c = sel.c0 + dc;
    const p = parseInput(t);
    contents.push(p.error || p.formula == null ? { r, c, value: p.error ? t : p.value } : { r, c, formula: p.formula });
  }));
  const area = { r0: sel.r0, c0: sel.c0, r1: sel.r0 + lines.length - 1, c1: sel.c0 + Math.max(...lines.map((l) => l.length)) - 1 };
  return [{ si, contents, styles: [], merges: wb.sheet(si).merges, area }];
}
