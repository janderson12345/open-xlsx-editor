// What typing text into a cell changes: its content as Excel reads the entry (input.js parseInput), and its number
// format when Excel would give it one (input.js formatAfterEntry: a date typed into a General cell is shown as a date).
// Typing, pasting text and Replace all go through here.
import { parseInput, formatAfterEntry } from './input.js';
import { formatCodeOf } from './styles.js';
import { spanSheetRefs, mapRefs } from '../calc/tokens.js';
import { FUNCTIONS } from '../calc/functions.js';

/** A typed formula's 3D references as Excel takes them (measured): in the workbook's order, Apr:Jan!A1 as Jan:Apr!A1,
 *  and one sheet's, Jan:Jan!A1, as Jan!A1; the sheets named as the workbook names them. */
function inOrder(wb, formula) {
  if (!formula.includes(':')) return formula;
  return spanSheetRefs(formula, ([a, b]) => {
    const ia = wb.sheetIndexOf(a);
    const ib = wb.sheetIndexOf(b);
    if (ia < 0 || ib < 0) return null;
    const [first, last] = [wb.sheets[Math.min(ia, ib)].name, wb.sheets[Math.max(ia, ib)].name];
    if (ia === ib) return first;
    return first === a && last === b ? null : [first, last];
  });
}

/** A defined name called as a function (a LAMBDA kept in it) written as the name is, whatever case it was typed in:
 *  =twice(5) is Twice(5), as Excel writes it (measured); functions keep their capitals. */
function namesAsWritten(wb, si, formula) {
  if (!wb.definedNames?.length) return formula;
  try {
    return mapRefs(formula, (t) => {
      if (t.type !== 'func' || FUNCTIONS[t.text.toUpperCase()]) return null;
      const d = wb.definedName(t.text, si);
      return d && d.name !== t.text ? d.name : null;
    });
  } catch {
    return formula;
  }
}

/**
 * The changes for `text` entered into cell (r, c) of sheet si: { content: { r, c, value } or { r, c, formula }, style:
 * { r, c, s } or null }, or { error } (a formula Excel wouldn't take). `today` gives the year of a date typed without
 * one.
 */
export function entryChanges(wb, si, r, c, text, today = new Date()) {
  const s = wb.sheet(si).cells.get(r)?.get(c)?.s || 0;
  const code = formatCodeOf(wb.styles, s);
  const parsed = parseInput(text, { today, date1904: wb.date1904, asText: code === '@' });
  if (parsed.error) return { error: parsed.error };
  const content = parsed.formula != null ? { r, c, formula: namesAsWritten(wb, si, inOrder(wb, parsed.formula)), ...(parsed.legacy ? { legacy: true } : {}) } : { r, c, value: parsed.value };
  const format = formatAfterEntry(code, parsed);
  return { content, style: format ? { r, c, s: wb.styleWith(s, { numFmt: format }) } : null, parsed };
}
