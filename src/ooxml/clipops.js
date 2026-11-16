// Copy, cut and paste, as Excel does them (measured over COM, docs/CLIPBOARD.md). A copy is a snapshot of cells:
// their values, formulas, formats and merged ranges. Pasting returns what changes, sheet by sheet, for the editor to
// apply and undo as one step: [{ si, contents: [{ r, c, value } or { r, c, formula }], styles: [{ r, c, s }], merges }].
import { shiftFormula, moveRefs, transposeFormula, sheetPrefix } from '../calc/tokens.js';
import { entryChanges } from './entry.js';
import { cellsOf } from './formatops.js';
import { formatCodeOf } from './styles.js';
import { refName, colName } from './refs.js';
import { readNotes } from './notes.js';
import { setNote, deleteNote } from './noteedit.js';
import { moveNamesForCut } from './names.js';
import { addThread, replyThread, resolveThread } from './threadedit.js';
import { boxRule, coalesce } from './dvedit.js';
import { copyConditionals, moveCutConditionals } from './clipcf.js';
import { maxOf } from './listmath.js';

export { pasteConditionals } from './clipcf.js';

const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
const inside = (a, b) => a.r0 >= b.r0 && a.r1 <= b.r1 && a.c0 >= b.c0 && a.c1 <= b.c1;

/** A snapshot of the selected cells of sheet si (the selection as it is, cut to the used range when it runs past it). */
export function snapshot(wb, si, sel, cut = false) {
  const sheet = wb.sheet(si);
  const cells = cellsOf(sheet, sel);
  const rect = { r0: sel.r0, c0: sel.c0, r1: maxOf(cells.map((x) => x.r)), c1: maxOf(cells.map((x) => x.c)) };
  return {
    wb, si, rect, cut,
    sheetName: wb.sheets[si].name,
    cells: cells.map(({ r, c }) => {
      const cell = sheet.cells.get(r)?.get(c);
      // (`result`: what a formula worked out to, for Paste Special's Values.)
      return { dr: r - rect.r0, dc: c - rect.c0, f: cell?.f || null, v: cell?.f ? null : cell?.v ?? null, t: cell?.t, s: cell?.s || 0, result: cell?.v ?? null };
    }),
    // (Each column's width when it isn't the standard one, for Paste Special's Column widths.)
    widths: Array.from({ length: rect.c1 - rect.c0 + 1 }, (_, dc) => sheet.cols.find((x) => x.min <= rect.c0 + dc && x.max >= rect.c0 + dc && x.width != null)?.width ?? null),
    merges: sheet.merges.filter((m) => inside(m, rect)).map((m) => ({ r0: m.r0 - rect.r0, c0: m.c0 - rect.c0, r1: m.r1 - rect.r0, c1: m.c1 - rect.c0 })),
    // (The notes of the copied cells, with their authors: a paste brings them, as Excel's does.)
    // (Comment threads too: their comments, authors and times, and whether resolved, as Excel copies them.)
    notes: [...readNotes(wb, si)].map(([key, n]) => { const [r, c] = key.split(',').map(Number); return { r, c, n }; })
      .filter(({ r, c }) => r >= rect.r0 && r <= rect.r1 && c >= rect.c0 && c <= rect.c1)
      .map(({ r, c, n }) => ({ dr: r - rect.r0, dc: c - rect.c0, ...(n.kind === 'thread' ? { thread: { items: n.items, resolved: !!n.resolved } } : { author: n.author, text: noteBody(n) }) })),
    // (The validation rules over the copied cells, each with the parts of it in the copy: a paste brings them.)
    validations: (sheet.validations || []).filter((v) => v.type !== 'none').map((v) => ({ v, parts: v.ranges.map((a) => clipTo(a, rect)).filter(Boolean) })).filter((x) => x.parts.length),
    // (And the conditional formats over them, rules and formats: clipcf.js.)
    conditionals: copyConditionals(wb, si, rect),
  };
}

/** Range `a` cut to `b`, or null when they don't meet. */
const clipTo = (a, b) => (overlaps(a, b) ? { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) } : null);

/** What a paste does to validation, as Excel's Paste, Paste Special's All and Validation do (measured over COM): the
 *  pasted cells end with the copy's rules, each over the cells it covered, moved with them; its formulas move as a
 *  copied formula does (relative references follow); cells the copy had no rule for, none. `area` and the copy as
 *  pasteNotes's. With Transpose (measured too) the same over the turned cells, a rule of its own moved by the
 *  corner only (=A1>B1 over A1:A3, turned onto another sheet's C1: =C1>D1 over C1:E1). Changes the workbook; returns
 *  whether anything changed. */
export function pasteValidation(wb, si, area, clip, { transpose = false } = {}) {
  const h = clip.rect.r1 - clip.rect.r0 + 1;
  const w = clip.rect.c1 - clip.rect.c0 + 1;
  const had = (wb.sheet(si).validations || []).some((v) => v.type !== 'none' && v.ranges.some((a) => overlaps(a, area)));
  if (!had && !clip.validations?.length) return false;
  if (had) wb.setValidation(si, [area], null);
  const sameSheet = clip.wb === wb && clip.si === si;
  for (const { v, parts } of clip.validations || []) {
    const ranges = [];
    // (Turned: each part's rows become columns from the pasted cells' corner, once.)
    const turn = (a) => ({ r0: area.r0 + (a.c0 - clip.rect.c0), c0: area.c0 + (a.r0 - clip.rect.r0), r1: area.r0 + (a.c1 - clip.rect.c0), c1: area.c0 + (a.r1 - clip.rect.r0) });
    if (transpose) for (const a of parts) ranges.push(turn(a));
    for (let tr = area.r0; !transpose && tr <= area.r1; tr += h) {
      for (let tc = area.c0; tc <= area.c1; tc += w) {
        for (const a of parts) {
          const moved = clipTo({ r0: a.r0 + tr - clip.rect.r0, c0: a.c0 + tc - clip.rect.c0, r1: a.r1 + tr - clip.rect.r0, c1: a.c1 + tc - clip.rect.c0 }, area);
          if (moved) ranges.push(moved);
        }
      }
    }
    if (!ranges.length) continue;
    const base = v.ranges[0];
    // (On the sheet copied from, the cells join the rule itself, as Excel's do; elsewhere a rule of their own, its
    // formulas written for its first cell as a formula copied there is: shifted from the rule's first cell to the
    // copied cell, and on as far as it was pasted.)
    if (sameSheet && wb.addToValidation(si, { first: { r: base.r0, c: base.c0 }, type: v.type, formula1: v.formulas[0] }, coalesce(ranges))) continue;
    const first = transpose ? turn(parts[0]) : { r0: parts[0].r0 + area.r0 - clip.rect.r0, c0: parts[0].c0 + area.c0 - clip.rect.c0 };
    const formulas = v.formulas.map((f) => (f == null ? f : shiftFormula(f, first.r0 - base.r0, first.c0 - base.c0)));
    wb.setValidation(si, coalesce(ranges), boxRule({ ...v, formulas }, wb.date1904));
  }
  return true;
}

/** A cut pasted at `area` of sheet si: its notes and validation move with its cells, as Excel's do (measured over
 *  COM): the notes taken from the cells cut and put on the pasted ones (a note pasted over goes), the rules as
 *  dvedit.js moveValidation moves them (to another sheet: taken off the cut cells and pasted there as a copy's are).
 *  Changes the workbook; returns whether anything changed. */
export function moveCutParts(wb, si, area, clip) {
  let changed = false;
  for (const x of clip.notes || []) { deleteNote(wb, clip.si, clip.rect.r0 + x.dr, clip.rect.c0 + x.dc); changed = true; }
  changed = pasteNotes(wb, si, area, clip) || changed;
  // (Defined names follow the cut cells as formulas do: names.js moveNamesForCut.)
  if (clip.wb === wb && moveNamesForCut(wb, { sheet: clip.sheetName, ...clip.rect }, { sheet: wb.sheets[si].name, dr: area.r0 - clip.rect.r0, dc: area.c0 - clip.rect.c0 })) changed = true;
  const had = (sheetOf) => (wb.sheet(sheetOf).validations || []).some((v) => v.type !== 'none');
  if (clip.si === si) {
    if (had(si)) { wb.moveValidation(si, clip.rect, area.r0 - clip.rect.r0, area.c0 - clip.rect.c0); changed = true; }
  } else {
    if (clip.validations?.length) { wb.setValidation(clip.si, [clip.rect], null); changed = true; }
    changed = pasteValidation(wb, si, area, clip) || changed;
  }
  changed = moveCutConditionals(wb, si, area, clip) || changed;
  return changed;
}

/** A note's own text, without the bold author line Excel puts before it. */
const noteBody = (n) => n.runs.filter((x, k) => !(k === 0 && x.bold)).map((x) => x.text).join('').replace(/^\n/, '');

/** What a paste does to notes and comment threads, as Excel's Paste, Paste Special's All and Comments and Notes do
 *  (measured over COM): each pasted cell ends with the copy's note or thread there (a thread with its replies, their
 *  authors and times, and resolved if it was; new ids), or none (one there before is taken away). `area` the pasted cells, the copy repeated over it every `rows` by `cols`; `transpose` as
 *  pasteSpecial's. Changes the workbook; returns whether anything changed. */
export function pasteNotes(wb, si, area, clip, { transpose = false } = {}) {
  const h = clip.rect.r1 - clip.rect.r0 + 1;
  const w = clip.rect.c1 - clip.rect.c0 + 1;
  const [rows, cols] = transpose ? [w, h] : [h, w];
  const want = new Map();
  for (let tr = area.r0; tr <= area.r1; tr += rows) {
    for (let tc = area.c0; tc <= area.c1; tc += cols) {
      for (const x of clip.notes || []) {
        const [dr, dc] = transpose ? [x.dc, x.dr] : [x.dr, x.dc];
        want.set(`${tr + dr},${tc + dc}`, x);
      }
    }
  }
  let changed = false;
  for (const [key, n] of readNotes(wb, si)) {
    const [r, c] = key.split(',').map(Number);
    if (r < area.r0 || r > area.r1 || c < area.c0 || c > area.c1) continue;
    const next = want.get(key);
    const same = next && (next.thread ? n.kind === 'thread' && !!n.resolved === next.thread.resolved && JSON.stringify(n.items) === JSON.stringify(next.thread.items)
      : n.kind === 'note' && n.author === next.author && noteBody(n) === next.text);
    if (same) { want.delete(key); continue; }
    deleteNote(wb, si, r, c);
    changed = true;
  }
  for (const [key, x] of want) {
    const [r, c] = key.split(',').map(Number);
    if (x.thread) {
      // (A thread: its comment, its replies, each with its author and time, then resolved if it was.)
      const [top, ...replies] = x.thread.items;
      const at = (d) => ({ when: d ? new Date(d) : new Date() });
      addThread(wb, si, r, c, top.text, top.author || 'Author', at(top.date));
      for (const y of replies) replyThread(wb, si, r, c, y.text, y.author || 'Author', at(y.date));
      if (x.thread.resolved) resolveThread(wb, si, r, c, true);
    } else setNote(wb, si, r, c, x.text, x.author || 'Author');
    changed = true;
  }
  return changed;
}

/** Where a paste goes: the copy repeated over the selection when the selection is a whole number of copies in both
 *  directions, else one copy from the selection's top-left cell (Excel's rule). */
export function pasteArea(sel, rows, cols) {
  const h = sel.r1 - sel.r0 + 1;
  const w = sel.c1 - sel.c0 + 1;
  const tile = h % rows === 0 && w % cols === 0;
  return { r0: sel.r0, c0: sel.c0, r1: sel.r0 + (tile ? h : rows) - 1, c1: sel.c0 + (tile ? w : cols) - 1 };
}

const valueOf = (cell) => (cell.t === 'e' ? { code: String(cell.v) } : cell.v);

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

/** Text from the clipboard as rows of cells: tabs between cells, a line for each row, and a cell starting with a
 *  quote running to its closing quote ("" a quote inside), line breaks and tabs and all, as Excel reads it (measured:
 *  "Plums, long text" pasted as Plums, long text; Excel's own copy quotes a cell with a line break so). A quote not
 *  closed, or with more after it, leaves the cell as written. */
export function tsvRows(text) {
  const t = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  const rows = [[]];
  let i = 0;
  while (i <= t.length) {
    let cell = null;
    if (t[i] === '"') {
      let j = i + 1;
      let inner = '';
      for (;;) {
        const q = t.indexOf('"', j);
        if (q < 0) { j = -1; break; }
        if (t[q + 1] === '"') { inner += t.slice(j, q + 1); j = q + 2; continue; }
        inner += t.slice(j, q);
        j = q + 1;
        break;
      }
      if (j >= 0 && (j >= t.length || t[j] === '\t' || t[j] === '\n')) { cell = inner; i = j; }
    }
    if (cell == null) {
      let j = i;
      while (j < t.length && t[j] !== '\t' && t[j] !== '\n') j++;
      cell = t.slice(i, j);
      i = j;
    }
    rows[rows.length - 1].push(cell);
    if (t[i] === '\n') rows.push([]);
    i++;
  }
  return rows;
}

/** Text from outside (tabs between cells, a line for each row; tsvRows) pasted at `sel`'s top-left cell, each cell
 *  read as if it were typed there (a number, a date, =a formula...), with the format typing would give it. */
export function pasteText(wb, si, sel, text) {
  const lines = tsvRows(text);
  const contents = [];
  const styles = [];
  lines.forEach((line, dr) => line.forEach((t, dc) => {
    const r = sel.r0 + dr;
    const c = sel.c0 + dc;
    const ch = entryChanges(wb, si, r, c, t);
    // (A formula Excel wouldn't take comes in as its text.)
    contents.push(ch.error ? { r, c, value: t } : ch.content);
    if (ch.style) styles.push(ch.style);
  }));
  const area = { r0: sel.r0, c0: sel.c0, r1: sel.r0 + lines.length - 1, c1: sel.c0 + maxOf(lines.map((l) => l.length)) - 1 };
  return [{ si, contents, styles, merges: wb.sheet(si).merges, area }];
}

/** What Paste Special pastes (its Paste choices): contents (formulas, or `values` their results), the whole format,
 *  or only the number format. */
const PASTES = {
  all: { contents: 'formulas', format: 'all' },
  formulas: { contents: 'formulas' },
  values: { contents: 'values' },
  formats: { format: 'all' },
  comments: {},
  validation: {},
  noBorders: { contents: 'formulas', format: 'noBorders' },
  formulasFormats: { contents: 'formulas', format: 'number' },
  valuesFormats: { contents: 'values', format: 'number' },
};
/** A number as a formula writes it (15 digits, as Excel keeps them). */
const numText = (n) => String(Number(n.toPrecision(15)));
const OPS = { add: '+', subtract: '-', multiply: '*', divide: '/' };
const operate = (op, a, b) => (op === 'add' ? a + b : op === 'subtract' ? a - b : op === 'multiply' ? a * b : b === 0 ? null : a / b);

/** A cell's content after Paste Special's operation: `dest` the cell there (or undefined), `src` { formula } or
 *  { value } from the copy. Text, on either side, is left as it was (null: no change); a blank is 0; a formula is
 *  combined as a formula, =(destination)+(copy), else the numbers worked out (dividing by 0, #DIV/0!). */
function operated(op, dest, src, r, c) {
  const destFormula = dest?.f != null && dest.f !== '' ? `(${dest.f})` : null;
  const destText = !destFormula && dest && dest.v != null && typeof dest.v !== 'number';
  const srcText = src.value != null && typeof src.value !== 'number';
  if (destText || srcText) return null;
  const dn = destFormula ? null : dest?.v ?? 0;
  if (src.formula != null) return { r, c, formula: `${destFormula ?? numText(dn)}${OPS[op]}(${src.formula})` };
  const sn = src.value ?? 0;
  if (destFormula) return { r, c, formula: `${destFormula}${OPS[op]}${numText(sn)}` };
  const out = operate(op, dn, sn);
  return { r, c, value: out == null ? { code: '#DIV/0!' } : out };
}

/**
 * Paste Special of a copy (not a cut) at `sel` on sheet si, as Excel's: `paste` one of all, formulas, values, formats,
 * noBorders (all except borders), formulasFormats and valuesFormats (with the number format only); `op` add,
 * subtract, multiply or divide (Excel's Operation; with formats alone, none); `skip` (Skip blanks: an empty cell of the
 * copy changes nothing); `transpose` (rows become columns, formulas as transposeFormula rewrites them). Steps as
 * pasteCopy's; merged ranges come with the whole format, and aren't transposed. Formats only within one workbook.
 */
export function pasteSpecial(wb, si, sel, clip, { paste = 'all', op = null, skip = false, transpose = false } = {}) {
  const what = PASTES[paste];
  const h = clip.rect.r1 - clip.rect.r0 + 1;
  const w = clip.rect.c1 - clip.rect.c0 + 1;
  const [rows, cols] = transpose ? [w, h] : [h, w];
  const area = pasteArea(sel, rows, cols);
  const sameBook = clip.wb === wb;
  const sheet = wb.sheet(si);
  const out = { si, contents: [], styles: [], merges: sheet.merges, area };
  const operation = op && what.contents ? op : null;
  for (let tr = area.r0; tr <= area.r1; tr += rows) {
    for (let tc = area.c0; tc <= area.c1; tc += cols) {
      for (const cell of clip.cells) {
        const [dr, dc] = transpose ? [cell.dc, cell.dr] : [cell.dr, cell.dc];
        const r = tr + dr;
        const c = tc + dc;
        if (skip && cell.f == null && cell.v == null) continue;
        const dest = sheet.cells.get(r)?.get(c);
        if (what.contents) {
          const from = { r: clip.rect.r0 + cell.dr, c: clip.rect.c0 + cell.dc };
          let src;
          if (!cell.f) src = { value: valueOf(cell) };
          else if (what.contents === 'values') src = { value: cell.t === 'e' ? { code: String(cell.result) } : cell.result };
          else if (transpose) src = { formula: transposeFormula(cell.f, { here: clip.sheetName, from: { sheet: clip.sheetName, ...clip.rect }, src: from, dst: { r, c }, to: { r: tr, c: tc } }) };
          else src = { formula: shiftFormula(cell.f, r - from.r, c - from.c) };
          const next = operation ? operated(operation, dest, src, r, c) : { r, c, ...src };
          if (next) out.contents.push(next);
        }
        if (!sameBook || !what.format) continue;
        const was = dest?.s || 0;
        if (what.format === 'all') out.styles.push({ r, c, s: cell.s });
        else if (what.format === 'number') out.styles.push({ r, c, s: wb.styleWith(was, { numFmt: formatCodeOf(wb.styles, cell.s) }) });
        else {
          // (All except borders: the copy's format with the destination's borders.)
          const sides = wb.styles.borders[wb.styles.cellXfs[was]?.borderId ?? 0];
          const border = Object.fromEntries(['left', 'right', 'top', 'bottom'].map((k) => [k, sides[k]?.style ? sides[k] : null]));
          out.styles.push({ r, c, s: wb.styleWith(cell.s, { border }) });
        }
      }
      if (what.format && what.format !== 'number' && sameBook && !transpose) {
        const here = { r0: tr, c0: tc, r1: tr + rows - 1, c1: tc + cols - 1 };
        out.merges = [...out.merges.filter((m) => !overlaps(m, here)), ...clip.merges.map((m) => ({ r0: tr + m.r0, c0: tc + m.c0, r1: tr + m.r1, c1: tc + m.c1 }))];
      }
    }
  }
  return [out];
}

/** Paste Special's Column widths: each column from `sel`'s first takes the width of the copy's column above it ([{ c,
 *  width }] for wb.setLines; a column of the copy with the standard width leaves its column as it is). */
export function pasteWidths(sel, clip) {
  return clip.widths.map((width, dc) => ({ c: sel.c0 + dc, width })).filter((x) => x.width != null);
}

/**
 * Paste Link at `sel` on sheet si, as Excel's: each cell of the paste area (the copy's size, or the selection when it is
 * a whole number of copies) a formula reading the cell of the copy as far from its top-left, relative (=A1, =B1...),
 * or for a single cell one absolute reference (=$A$1); with its sheet's name when the copy is on another. Cells in the
 * General format take the number format of the copy's first cell, as Excel gives them.
 */
export function pasteLink(wb, si, sel, clip) {
  const h = clip.rect.r1 - clip.rect.r0 + 1;
  const w = clip.rect.c1 - clip.rect.c0 + 1;
  const area = pasteArea(sel, h, w);
  const sheet = wb.sheet(si);
  const prefix = clip.si === si ? '' : sheetPrefix(clip.sheetName);
  const one = h === 1 && w === 1;
  const first = clip.cells.find((x) => x.dr === 0 && x.dc === 0);
  const code = formatCodeOf(wb.styles, first?.s || 0);
  const out = { si, contents: [], styles: [], merges: sheet.merges, area };
  for (let r = area.r0; r <= area.r1; r++) {
    for (let c = area.c0; c <= area.c1; c++) {
      const ref = one ? `$${colName(clip.rect.c0)}$${clip.rect.r0 + 1}` : refName(clip.rect.r0 + r - area.r0, clip.rect.c0 + c - area.c0);
      out.contents.push({ r, c, formula: prefix + ref });
      const s = sheet.cells.get(r)?.get(c)?.s || 0;
      if (code !== 'General' && formatCodeOf(wb.styles, s) === 'General') out.styles.push({ r, c, s: wb.styleWith(s, { numFmt: code }) });
    }
  }
  return [out];
}
