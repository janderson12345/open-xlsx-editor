// What a person types into a cell, as the value it stores: Excel's reading of an entry, for the common cases. Dates,
// times, currency and fractions typed as text are read as text for now; formulas are not accepted until there is a
// calculation engine (they would need a value, and calcChain.xml, to be right).

/** A typed entry as { value } (a number, a boolean, a string, or null for an empty entry) or { error } when it can't be
 *  stored yet. Numbers: an optional sign, digits with or without thousands separators in groups of three, a decimal
 *  point, an exponent, a trailing % (divided by 100). TRUE and FALSE in any case are booleans. A leading apostrophe
 *  keeps the rest as text ('007), as in Excel. */
export function parseInput(text) {
  const s = String(text ?? '');
  if (s === '') return { value: null };
  if (s.startsWith("'")) return { value: s.slice(1) };
  if (s.startsWith('=')) return { error: 'Formulas are not supported yet' };
  const t = s.trim();
  if (/^(true|false)$/i.test(t)) return { value: t.toLowerCase() === 'true' };
  const m = /^([+-]?)(\d{1,3}(?:,\d{3})+|\d+)?(\.\d*)?(?:[eE]([+-]?\d+))?(%?)$/.exec(t);
  if (m && (m[2] || (m[3] && m[3].length > 1))) {
    let v = Number(`${m[1]}${(m[2] || '0').replace(/,/g, '')}${m[3] || ''}${m[4] ? `e${m[4]}` : ''}`);
    if (m[5]) v /= 100;
    if (Number.isFinite(v)) return { value: v };
  }
  return { value: s };
}

/** A cell's value as the text to edit it from: numbers in full (not as formatted), booleans as TRUE/FALSE, a formula
 *  with its = sign. */
export function editText(cell) {
  if (!cell) return '';
  if (cell.f) return `=${cell.f}`;
  if (cell.v == null) return '';
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (cell.t === 'n') return String(cell.v);
  return String(cell.v);
}
