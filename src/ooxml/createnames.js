// Formulas > Defined Names > Create from Selection (Ctrl+Shift+F3): names made from the labels in a selection's top
// row, left column, bottom row or right column, each for the cells beside it, as Excel's Range.CreateNames makes them.
// Measured over COM in a hidden Excel, 86 labels in three probes (docs/NAMES.md):
// - Only text labels name anything: a number, TRUE or FALSE, an error or an empty cell names nothing.
// - Characters a name can't have are dropped from the label's ends (spaces too: "  padded  ", "Total!", "Price ($)",
//   "(x)"), and each one between becomes _ ("Unit Price" Unit_Price, "a!-!b" a___b, "$A$1" A_1, a line break _).
// - A digit, . or ? first gets _ before it (1st _1st, .dot _.dot, the text '2020 _2020); so does R or C and a digit,
//   or RC and a digit (R1C1 _R1C1, R2D2 _R2D2, C3 _C3, rc2x _rc2x).
// - R, C or RC alone, a cell's address, and the text TRUE or FALSE get _ after (R_, rc_, A1_, XFD1_, TRUE_).
// - At most 255 characters; a label still no name (\x, _ alone) names nothing; Print_Area and Print_Titles, which
//   Excel makes its own built-in names of, are left out here.
// - A name there is already (workbook-wide, capitals aside) is replaced, as Excel replaced it without its question.
// - Each name is workbook-wide and refers to the cells absolutely, with their sheet: Top row, the cells under each
//   label; Left column, those right of it; Bottom row, above; Right column, left; where two of them meet, the corner's
//   label names all the cells inside them.
import { validName, findName, defineName, editName } from './names.js';
import { colIndex, colName, MAX_ROWS, MAX_COLS } from './refs.js';
import { sheetPrefix } from '../calc/tokens.js';

/** What a name can have after its first character (names.js validName). */
const OK = /[A-Za-z0-9_.?\\À-￿]/;

/** The name Excel makes of the text label `text`, or null when it makes none (the rules above). */
export function labelName(text) {
  const chars = [...String(text ?? '')];
  let a = 0;
  let b = chars.length;
  while (a < b && !OK.test(chars[a])) a++;
  while (b > a && !OK.test(chars[b - 1])) b--;
  let n = chars.slice(a, b).map((ch) => (OK.test(ch) ? ch : '_')).join('');
  if (!n || n === '_') return null;
  if (/^[0-9.?]/.test(n) || /^(R|C|RC)\d/i.test(n)) n = `_${n}`;
  const a1 = /^([A-Za-z]{1,3})(\d+)$/.exec(n);
  if (/^(R|C|RC|TRUE|FALSE)$/i.test(n) || (a1 && colIndex(a1[1]) < MAX_COLS && Number(a1[2]) >= 1 && Number(a1[2]) <= MAX_ROWS)) n = `${n}_`;
  n = n.slice(0, 255);
  if (/^(Print_Area|Print_Titles)$/i.test(n) || !validName(n)) return null;
  return n;
}

/** The text label at (r, c) of sheet si, or null (a text value, or a formula's text result). */
function labelAt(sheet, r, c) {
  const cell = sheet.cells.get(r)?.get(c);
  return cell && (cell.t === 's' || cell.t === 'str' || cell.t === 'inlineStr') && typeof cell.v === 'string' ? cell.v : null;
}

/** The names Create from Selection would make for `sel` ({ r0, c0, r1, c1 }) of sheet si with the sides `sides`
 *  ({ top, left, bottom, right }): [{ name, refersTo, r, c }] in the order made (a later one of the same name wins);
 *  null when the sides leave no cells to name. */
export function namesFromSelection(wb, si, sel, sides) {
  const { top, left, bottom, right } = sides;
  const body = { r0: sel.r0 + (top ? 1 : 0), r1: sel.r1 - (bottom ? 1 : 0), c0: sel.c0 + (left ? 1 : 0), c1: sel.c1 - (right ? 1 : 0) };
  if (!(top || left || bottom || right) || body.r0 > body.r1 || body.c0 > body.c1) return null;
  const sheet = wb.sheet(si);
  const prefix = sheetPrefix(wb.sheets[si].name);
  const abs = (r, c) => `$${colName(c)}$${r + 1}`;
  const refers = (r0, c0, r1, c1) => `=${prefix}${abs(r0, c0)}${r0 === r1 && c0 === c1 ? '' : `:${abs(r1, c1)}`}`;
  const out = [];
  const add = (r, c, area) => {
    const name = labelName(labelAt(sheet, r, c));
    if (name) out.push({ name, refersTo: refers(area.r0, area.c0, area.r1, area.c1), r, c });
  };
  for (let c = body.c0; c <= body.c1; c++) {
    if (top) add(sel.r0, c, { ...body, c0: c, c1: c });
    if (bottom) add(sel.r1, c, { ...body, c0: c, c1: c });
  }
  for (let r = body.r0; r <= body.r1; r++) {
    if (left) add(r, sel.c0, { ...body, r0: r, r1: r });
    if (right) add(r, sel.c1, { ...body, r0: r, r1: r });
  }
  // (The corners where two sides meet: all the cells inside.)
  if (top && left) add(sel.r0, sel.c0, body);
  if (top && right) add(sel.r0, sel.c1, body);
  if (bottom && left) add(sel.r1, sel.c0, body);
  if (bottom && right) add(sel.r1, sel.c1, body);
  return out;
}

/** The sides Excel's box starts with checked (not measured: COM doesn't reach the box): Top row when the selection
 *  has more than one row and its top row a text label past the first cell, Left column the same for its left column;
 *  neither of those: Top row for a column, Left column for a row. */
export function guessSides(wb, si, sel) {
  const sheet = wb.sheet(si);
  const any = (cells) => cells.some(([r, c]) => labelAt(sheet, r, c) != null);
  const across = [];
  for (let c = sel.c0 + 1; c <= sel.c1; c++) across.push([sel.r0, c]);
  const down = [];
  for (let r = sel.r0 + 1; r <= sel.r1; r++) down.push([r, sel.c0]);
  let top = sel.r1 > sel.r0 && (sel.c0 === sel.c1 ? labelAt(sheet, sel.r0, sel.c0) != null : any(across));
  let left = sel.c1 > sel.c0 && (sel.r0 === sel.r1 ? labelAt(sheet, sel.r0, sel.c0) != null : any(down));
  if (!top && !left) { top = sel.r1 > sel.r0 && sel.c0 === sel.c1; left = !top && sel.c1 > sel.c0; }
  return { top, left, bottom: false, right: false };
}

/** The names made (`names`: namesFromSelection's, or those of them kept): a new one defined, one there already
 *  (workbook-wide) given the new cells and the label's capitals. Returns how many. */
export function createNames(wb, names) {
  for (const { name, refersTo } of names) {
    const had = findName(wb, name, null);
    if (had) editName(wb, { name: had.name, sheet: null }, { name, refersTo });
    else defineName(wb, { name, refersTo });
  }
  return names.length;
}
