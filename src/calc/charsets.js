// Characters as Excel 365 counts and codes them (measured over COM, docs/CALC.md): the text functions count code
// points, so an emoji (two UTF-16 units) is one character; CHAR and CODE are Windows-1252 (code page 1252), a
// character it hasn't coded 63, "?", with no nearest letter taken (Ā is 63, not A's 65); UNICHAR refuses the code
// points that aren't characters.

/** Whether `s` has a character beyond the first 65,536 (a surrogate pair, or half of one). */
const astral = (s) => /[\ud800-\udfff]/.test(s);

/** `s` as its characters (code points), or null when each is one UTF-16 unit (the common case: plain indexing). */
export const pointsOf = (s) => (astral(s) ? Array.from(s) : null);

/** The UTF-16 index where character `i` (0-based) of `s` starts; `s.length` past the end. */
export function unitAt(s, i) {
  if (!astral(s)) return Math.min(i, s.length);
  let u = 0;
  for (let k = 0; k < i && u < s.length; k++) u += s.codePointAt(u) > 0xffff ? 2 : 1;
  return u;
}

/** How many characters (code points) the first `u` UTF-16 units of `s` hold. */
export const pointsBefore = (s, u) => (astral(s) ? Array.from(s.slice(0, u)).length : u);

// (Code page 1252's 0x80 to 0x9F; 0x81, 0x8D, 0x8F, 0x90 and 0x9D are unassigned and stand for themselves.)
const HIGH = [0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f,
  0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178];
const BACK = new Map(HIGH.map((u, i) => [u, 0x80 + i]));

/** CHAR's character for code `n` (1 to 255). */
export const ansiChar = (n) => String.fromCharCode(n >= 0x80 && n <= 0x9f ? HIGH[n - 0x80] : n);

/** CODE's number for the first character of `s`: its code page 1252 code, 63 where it has none. */
export function ansiCode(s) {
  const u = s.codePointAt(0);
  // (The control characters U+0080 to U+009F keep their own numbers too, as Excel gives them.)
  if (u <= 0xff) return u;
  return BACK.get(u) ?? 63;
}

/** Whether UNICHAR refuses code point `n` (#N/A): a surrogate, or a noncharacter (U+FDD0 to U+FDEF, and the last two
 *  of every plane). */
export const notACharacter = (n) => (n >= 0xd800 && n <= 0xdfff) || (n >= 0xfdd0 && n <= 0xfdef) || (n & 0xfffe) === 0xfffe;
