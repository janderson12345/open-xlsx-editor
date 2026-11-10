// What typing text into a cell changes: its content as Excel reads the entry (input.js parseInput), and its number
// format when Excel would give it one (input.js formatAfterEntry: a date typed into a General cell is shown as a date).
// Typing, pasting text and Replace all go through here.
import { parseInput, formatAfterEntry } from './input.js';
import { formatCodeOf } from './styles.js';

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
  const content = parsed.formula != null ? { r, c, formula: parsed.formula } : { r, c, value: parsed.value };
  const format = formatAfterEntry(code, parsed);
  return { content, style: format ? { r, c, s: wb.styleWith(s, { numFmt: format }) } : null, parsed };
}
