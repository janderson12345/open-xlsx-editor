// What the Name Box goes to when a reference or a name is typed into it, as Excel's Name Box takes them (Excel's
// documented behaviour: COM doesn't reach the Name Box): a cell or a range ("B5", "b5:d10", "$B$5"), whole columns or
// rows ("B:D", "5:5"), any of these on another sheet ("Sheet2!A1", "'My data'!B:B"), or a defined name (the sheet's
// own first, then the workbook's) that stands for a range, or a table: its name for its data rows, or a table
// reference ("Fruit[Qty]", "Fruit[#All]").
import { parseRange, MAX_ROWS, MAX_COLS, colIndex } from './refs.js';
import { readTables, tableArea } from './tables.js';

export const GOTO_MESSAGES = {
  invalid: 'You must enter a valid reference you want to go to, or type a valid name for the selection.',
  defining: "Naming the selection from the Name Box isn't available yet",
};

/** "B:D" or "5:7" (with or without $) as a range of whole columns or rows; null otherwise. */
function wholeLines(text) {
  let m = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(text);
  if (m) {
    const [a, b] = [colIndex(m[1]), colIndex(m[2])];
    if (a < MAX_COLS && b < MAX_COLS) return { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(a, b), c1: Math.max(a, b) };
  }
  m = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(text);
  if (m) {
    const [a, b] = [Number(m[1]) - 1, Number(m[2]) - 1];
    if (a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS) return { r0: Math.min(a, b), r1: Math.max(a, b), c0: 0, c1: MAX_COLS - 1 };
  }
  return null;
}

/** A name as Excel allows one: a letter, _ or \ first, then letters, digits, _, . and \ (not a cell reference). */
const NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

/**
 * Where `text` typed into the Name Box goes from sheet si: { si, sel: { r0, c0, r1, c1, ar, ac } } (the active cell
 * the range's first), or { error } with Excel's words, or { error, name: true } for a name that isn't there yet (Excel
 * would name the selection with it).
 */
export function goToTarget(wb, si, text, depth = 0) {
  const t = String(text ?? '').trim();
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(t);
  let sheet = si;
  let ref = t;
  if (m) {
    sheet = wb.sheetIndexOf((m[1] ?? m[2]).replace(/''/g, "'"));
    // (A hidden sheet can't be gone to, as Excel can't select on one.)
    if (sheet < 0 || !wb.canShow(sheet)) return { error: GOTO_MESSAGES.invalid };
    ref = m[3];
  }
  // (A table, by its name (its data rows) or a table reference.)
  const tm = !m && /^([A-Za-z_\\][A-Za-z0-9_.\\]*)(?:\[(.*)\])?$/.exec(t);
  if (tm) {
    const tables = readTables(wb);
    if (tables.some((x) => x.name.toLowerCase() === tm[1].toLowerCase())) {
      const a = tableArea(tables, tm[1], tm[2] ?? '#Data', { si, r: -1, c: -1 });
      if (a.error) return { error: GOTO_MESSAGES.invalid };
      if (!wb.canShow(a.si)) return { error: GOTO_MESSAGES.invalid };
      return { si: a.si, sel: { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1, ar: a.r0, ac: a.c0 } };
    }
  }
  // (parseRange reads a first and a second corner only: a third part is no reference here.)
  const area = ref.split(':').length <= 2 && (parseRange(ref) || wholeLines(ref));
  if (area) return { si: sheet, sel: { ...area, ar: area.r0, ac: area.c0 } };
  if (!m && depth === 0 && NAME.test(t)) {
    const named = wb.definedName(t, si);
    if (!named) return { error: GOTO_MESSAGES.defining, name: true };
    // (A name standing for one range: its text is a reference like Sheet1!$B$2:$B$9.)
    const target = goToTarget(wb, si, named.text, 1);
    return target.error ? { error: GOTO_MESSAGES.invalid } : target;
  }
  return { error: GOTO_MESSAGES.invalid };
}

/** What Excel's Go To box lists: the defined names that stand for ranges (not its own hidden ones, nor those of
 *  another sheet) and the tables, in order. */
export function goToNames(wb, si) {
  const names = wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.') && (d.sheet == null || d.sheet === si) && !goToTarget(wb, si, d.text, 1).error).map((d) => d.name);
  return [...new Set([...names, ...readTables(wb).map((t) => t.name)])].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
}
