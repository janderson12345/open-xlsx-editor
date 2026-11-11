// A criteria cell's test, as Advanced Filter (ooxml/advfilter.js) and the database functions (DSUM, DCOUNT...:
// functions-db.js) read it: =x the value x itself, x alone a value beginning with x (text; a number: equal), or >, >=,
// <, <=, <> before a value; * and ? are wildcards, ~ before one the character itself. Case aside.

const pattern = (v, whole) => new RegExp(`^${String(v).replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (m, lit, wild, re) => (lit ? `\\${lit}` : wild === '*' ? '[\\s\\S]*' : wild ? '[\\s\\S]' : `\\${re}`))}${whole ? '$' : ''}`, 'i');

/** One criteria cell's test of a value (and its shown text). */
export function cellTest(raw) {
  const s = String(raw);
  const m = /^(<>|>=|<=|=|>|<)?([\s\S]*)$/.exec(s);
  const op = m[1] || '';
  const val = m[2];
  const n = Number(val);
  const numeric = val.trim() !== '' && Number.isFinite(n);
  return (v, text) => {
    if (op === '' || op === '=') {
      if (op === '=' && val === '') return v == null || v === '';
      if (numeric && typeof v === 'number') return v === n;
      // (Alone: begins with; after =: the whole of it.)
      return pattern(val, op === '=').test(text);
    }
    if (op === '<>') return !(numeric && typeof v === 'number' ? v === n : pattern(val, true).test(text));
    if (v == null || v === '') return false;
    let d;
    if (numeric) { if (typeof v !== 'number') return false; d = v - n; } else { if (typeof v === 'number') return false; d = text.toLowerCase().localeCompare(val.toLowerCase()); }
    return op === '>' ? d > 0 : op === '>=' ? d >= 0 : op === '<' ? d < 0 : d <= 0;
  };
}
