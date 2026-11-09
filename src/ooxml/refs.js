// A1 references. Rows and columns are 0-based inside the code (row 0 is "1", column 0 is "A"); the 1-based numbers
// and letters appear only in the file and on screen.

/** The last row and column a sheet can have (1,048,576 rows by 16,384 columns, "XFD"). */
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** Column letters for a 0-based column: 0 is "A", 25 "Z", 26 "AA". */
export function colName(c) {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** The 0-based column of column letters ("A" is 0), case ignored. */
export function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** "B12" (or "$B$12") to { r: 11, c: 1 }; null if it isn't a cell reference. */
export function parseRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref || '');
  if (!m) return null;
  const r = Number(m[2]) - 1;
  const c = colIndex(m[1]);
  return r >= 0 && r < MAX_ROWS && c >= 0 && c < MAX_COLS ? { r, c } : null;
}

/** { r: 11, c: 1 } to "B12". */
export const refName = (r, c) => colName(c) + (r + 1);

/** "A1:C3" (or one cell, "B2") to { r0, c0, r1, c1 }, corners in order; null if it isn't a range. */
export function parseRange(range) {
  const [a, b = a] = String(range || '').split(':');
  const p = parseRef(a);
  const q = parseRef(b);
  if (!p || !q) return null;
  return { r0: Math.min(p.r, q.r), c0: Math.min(p.c, q.c), r1: Math.max(p.r, q.r), c1: Math.max(p.c, q.c) };
}

/** { r0, c0, r1, c1 } to "A1:C3" ("B2" for one cell). */
export const rangeName = ({ r0, c0, r1, c1 }) => (r0 === r1 && c0 === c1 ? refName(r0, c0) : `${refName(r0, c0)}:${refName(r1, c1)}`);
