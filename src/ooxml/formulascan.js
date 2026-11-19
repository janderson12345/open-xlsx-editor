// Whether a worksheet part has a formula saved without its result, found by scanning its bytes, without parsing the
// XML: Workbook.calculateMissing parses only such sheets on opening. A workbook Excel saved has every result saved
// (<c><f>..</f><v>..</v></c>), so its sheets aren't parsed until shown: a 25 MB workbook of two 100 MB sheets opened
// in 78 s, all of it parsing sheets whose 3,307 formulas all had results, and opens in a few seconds now.

const LT = 0x3c;
const GT = 0x3e;
const SLASH = 0x2f;
const COLON = 0x3a;
const F = 0x66;
const V = 0x76;
/** After a tag's name: a space, tab, line break, '>' or '/'. */
const ends = (b) => b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d || b === GT || b === SLASH;
const space = (b) => b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d;
const letter = (b) => (b >= 0x61 && b <= 0x7a) || (b >= 0x41 && b <= 0x5a) || (b >= 0x30 && b <= 0x39) || b === 0x5f || b === 0x2d || b === 0x2e;

/** The index of `pattern` (bytes) in `u8` from `from`, or -1. */
function find(u8, pattern, from = 0) {
  const n = pattern.length;
  const first = pattern[0];
  for (let i = u8.indexOf(first, from); i >= 0 && i + n <= u8.length; i = u8.indexOf(first, i + 1)) {
    let k = 1;
    while (k < n && u8[i + k] === pattern[k]) k++;
    if (k === n) return i;
  }
  return -1;
}
const bytes = (s) => Uint8Array.from(s, (ch) => ch.charCodeAt(0));
const SHEET_DATA = bytes('sheetData');

/** Where a tag named `name` (one byte, with or without a namespace prefix) starts at `i` ('<' there): the index after
 *  its name, or -1. */
function tagAt(u8, i, name) {
  let j = i + 1;
  if (u8[j] === SLASH) return -1;
  // (A prefix: letters then ':'.)
  let k = j;
  while (k < u8.length && letter(u8[k])) k++;
  if (u8[k] === COLON) j = k + 1;
  return u8[j] === name && ends(u8[j + 1]) ? j + 1 : -1;
}

/**
 * True when the worksheet XML (bytes) has, inside its sheetData, a formula element (<f>, any prefix) not followed by a
 * value (<v>) in its cell; false when every formula there has one, or there are none. Formulas outside sheetData (a
 * validation rule's xm:f in the extLst) don't count.
 */
export function hasFormulaWithoutResult(u8) {
  const open = find(u8, SHEET_DATA);
  if (open < 0) return false;
  // (Its closing tag, looked for from the part's end: the rows between aren't walked for it.)
  let close = -1;
  for (let i = u8.lastIndexOf(SLASH); i > open; i = u8.lastIndexOf(SLASH, i - 1)) {
    if (u8[i - 1] !== LT) continue;
    let k = i + 1;
    while (k < u8.length && letter(u8[k])) k++;
    const name = u8[k] === COLON ? k + 1 : i + 1;
    let n = 0;
    while (n < SHEET_DATA.length && u8[name + n] === SHEET_DATA[n]) n++;
    if (n === SHEET_DATA.length) { close = name; break; }
  }
  const end = close < 0 ? u8.length : close;
  for (let i = u8.indexOf(LT, open); i >= 0 && i < end; i = u8.indexOf(LT, i + 1)) {
    const after = tagAt(u8, i, F);
    if (after < 0) continue;
    // (The element's end: '/>' for an empty one, else past its closing tag.)
    let gt = u8.indexOf(GT, after);
    if (gt < 0) return true;
    let j;
    if (u8[gt - 1] === SLASH) j = gt + 1;
    else {
      j = gt + 1;
      for (;;) {
        const lt = u8.indexOf(LT, j);
        if (lt < 0) return true;
        if (u8[lt + 1] === SLASH) {
          let k = lt + 2;
          while (k < u8.length && letter(u8[k])) k++;
          const nameAt = u8[k] === COLON ? k + 1 : lt + 2;
          if (u8[nameAt] === F && (u8[nameAt + 1] === GT || space(u8[nameAt + 1]))) { gt = u8.indexOf(GT, nameAt); j = gt + 1; break; }
        }
        j = lt + 1;
      }
    }
    while (j < u8.length && space(u8[j])) j++;
    if (u8[j] !== LT || tagAt(u8, j, V) < 0) return true;
    i = j;
  }
  return false;
}
