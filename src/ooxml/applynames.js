// Formulas > Define Name > Apply Names...: references in formulas replaced by the names that refer to them, as
// Excel's Range.ApplyNames does. Measured over COM in a hidden Excel, each formula on its own (docs/NAMES.md):
// - Only references without a sheet name: Sheet1!B2 and Data!$A$1 stay, even with names for them. Names that refer
//   to one area of the sheet, all of it $ (a relative name, a constant, a formula: never applied).
// - Ignore Relative/Absolute (on in Excel's box): B2, $B$2, B$2 and $C2 all become Qty (=$B$2); off, only $B$2.
// - A range that is a name's becomes it (B2:B5, $B$2:$B$5: Amounts). In a range that isn't, an end whose column is
//   relative and a cell's name becomes the name (B2:B4: Qty:B4; A1:B2: A1:Qty; B2:C2: Qty:Price) and a $ column's
//   doesn't ($B$2:$B$4, $B2:$B4 stay) - Ignore Relative/Absolute on only. A one-cell range (C2:C2), which Excel made
//   @Price, is left here.
// - Use row and column names (on in Excel's box): a cell in a column's name (Jan: C2:C4) and a row's (South: C3:D3)
//   becomes the two, row first ("South Jan"; Column row order: "Jan South"). From the cell's own row the column's name
//   alone (Omit row name if same row: =C2 in H2, Jan), from its own column the row's alone (Omit column name if same
//   column: =C3 in C6, South). This comes before a name of the cell itself (=B2 in H2 with Qty =$B$2 and Amounts
//   =$B$2:$B$5: Amounts); without a row and a column, the cell's own name.
// - Text in quotes, names already there and everything else stay as typed.
import { tokenize, areaText } from '../calc/tokens.js';
import { namesOnSheet } from './names.js';
import { colName } from './refs.js';
import { maxOf, minOf } from './listmath.js';

/** The names Apply Names can use on sheet si: those the sheet sees that refer to one area of it, all $:
 *  [{ name, area }]. */
export function applicableNames(wb, si) {
  const out = [];
  const here = wb.sheets[si].name.toLowerCase();
  for (const n of namesOnSheet(wb, si)) {
    let toks;
    try { toks = tokenize(n.text.replace(/^=/, '')); } catch { continue; }
    const t = toks[0];
    if (toks.length !== 1 || t.type !== 'ref' || !t.area || t.book != null || String(t.sheet ?? '').toLowerCase() !== here) continue;
    const a = t.area;
    if (a.kind !== 'cell' && a.kind !== 'area') continue;
    if (!(a.ar0 && a.ac0 && (a.kind === 'cell' || (a.ar1 && a.ac1)))) continue;
    out.push({ name: n.name, area: { r0: a.r0, c0: a.c0, r1: a.kind === 'cell' ? a.r0 : a.r1, c1: a.kind === 'cell' ? a.c0 : a.c1 } });
  }
  return out;
}

const cellText = (c, r, ac, ar) => `${ac ? '$' : ''}${colName(c)}${ar ? '$' : ''}${r + 1}`;
const sameArea = (a, b) => a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;
const inside = (a, r, c) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1;

/** Formula `text` (without "=") in cell `at` ({ r, c }) with `names` (applicableNames', or those chosen) applied, as the
 *  rules above; null when nothing changes. Options as Excel's box: `ignore` (Ignore Relative/Absolute), `rowCol` (Use
 *  row and column names), `omitCol`, `omitRow`, `colFirst` (Column row order). */
export function applyNamesTo(text, at, names, { ignore = true, rowCol = true, omitCol = true, omitRow = true, colFirst = false } = {}) {
  let toks;
  try { toks = tokenize(text); } catch { return null; }
  const cellName = (r, c) => names.find((n) => n.area.r0 === r && n.area.c0 === c && n.area.r1 === r && n.area.c1 === c)?.name;
  const allAbs = (a) => a.ar0 && a.ac0 && (a.kind === 'cell' || (a.ar1 && a.ac1));
  let out = '';
  let last = 0;
  let changed = false;
  for (const t of toks) {
    if (t.type !== 'ref' || t.prefix || t.book != null || !t.area || t.spill) continue;
    const a = t.area;
    let next = null;
    if (a.kind === 'cell') {
      const { r0: r, c0: c } = a;
      if (rowCol && (ignore || allAbs(a))) {
        const col = names.find((n) => n.area.c0 === c && n.area.c1 === c && n.area.r1 > n.area.r0 && inside(n.area, r, c))?.name;
        const row = names.find((n) => n.area.r0 === r && n.area.r1 === r && n.area.c1 > n.area.c0 && inside(n.area, r, c))?.name;
        if (col && omitRow && at.r === r) next = col;
        else if (row && omitCol && at.c === c) next = row;
        else if (row && col) next = colFirst ? `${col} ${row}` : `${row} ${col}`;
      }
      if (!next && (ignore || allAbs(a))) next = cellName(r, c) ?? null;
    } else if (a.kind === 'area') {
      const whole = names.find((n) => sameArea(n.area, a) && !(a.r0 === a.r1 && a.c0 === a.c1));
      if (whole && (ignore || allAbs(a))) next = whole.name;
      else if (ignore && !(a.r0 === a.r1 && a.c0 === a.c1)) {
        // (Each end by itself, when its column is relative.)
        const first = !a.ac0 ? cellName(a.r0, a.c0) : null;
        const second = !a.ac1 ? cellName(a.r1, a.c1) : null;
        if (first || second) next = `${first ?? cellText(a.c0, a.r0, a.ac0, a.ar0)}:${second ?? cellText(a.c1, a.r1, a.ac1, a.ar1)}`;
      }
    }
    if (next == null || next === areaText(a)) continue;
    out += text.slice(last, t.start) + next;
    last = t.end;
    changed = true;
  }
  return changed ? out + text.slice(last) : null;
}

/** Apply Names over `sel` of sheet si (one cell: the whole sheet, as Excel's): { steps, count } for runSteps, or
 *  { none: true } when no reference takes a name. `names`: those chosen ([{ name, area }]); `opts` as applyNamesTo. */
export function applyNamesSteps(wb, si, sel, names, opts = {}) {
  const sheet = wb.sheet(si);
  const whole = sel.r0 === sel.r1 && sel.c0 === sel.c1;
  const inArray = (r, c) => sheet.arrays?.some((a) => inside(a, r, c));
  const contents = [];
  for (const [r, line] of sheet.cells) {
    if (!whole && (r < sel.r0 || r > sel.r1)) continue;
    for (const [c, cell] of line) {
      if (!whole && (c < sel.c0 || c > sel.c1)) continue;
      if (cell.f == null || cell.f === '' || inArray(r, c)) continue;
      const next = applyNamesTo(cell.f, { r, c }, names, opts);
      if (next != null) contents.push({ r, c, formula: next });
    }
  }
  if (!contents.length) return { none: true };
  const rs = contents.map((x) => x.r);
  const cs = contents.map((x) => x.c);
  return { count: contents.length, steps: [{ si, contents, styles: [], area: { r0: minOf(rs), c0: minOf(cs), r1: maxOf(rs), c1: maxOf(cs) } }] };
}
