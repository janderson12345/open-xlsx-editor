// Turning a cell's value into the text a sheet shows: Excel's General format here, format codes in format.js.
import { formatValue } from './format.js';
import { currentLocale, localCode, localDecimal } from './locale.js';

/** The most characters General uses, not counting a minus sign ("0.333333333", "-0.666666667"). */
const GENERAL_MAX = 11;

/** |n| as Excel holds it, 15 significant decimal digits: { digits: '15', exp: -10 } for 1.5E-10 (value 0.15 x 10^-9:
 *  the point goes after the first digit, times 10^exp). Rounding works on these digits, half up, as Excel rounds:
 *  1.5E-10 to no decimals is 2E-10, where binary rounding gives 1E-10. */
export function decimal(a) {
  const [m, e] = a.toExponential(14).split('e');
  return { digits: m.replace('.', '').replace(/0+$/, '') || '0', exp: Number(e) };
}

/** Round a digit string to `keep` digits, half up; { digits, carried } (carried: it became one digit longer). */
export function roundTo(digits, keep) {
  if (keep >= digits.length) return { digits, carried: false };
  if (keep < 0) return { digits: '', carried: false };
  const up = digits.charCodeAt(keep) >= 53; // ('5')
  let d = digits.slice(0, keep).split('').map(Number);
  if (up) {
    let i = d.length - 1;
    while (i >= 0 && d[i] === 9) d[i--] = 0;
    if (i < 0) return { digits: '1' + d.join(''), carried: true };
    d[i]++;
  }
  return { digits: d.join(''), carried: false };
}

/** Significant digits a fixed form with this many characters has room for (the digits after any leading zeros; 0 or
 *  less when the value is too small to show any), or null when its integer part doesn't fit. */
function fixedRoom(v, room) {
  const intDigits = v.exp >= 0 ? v.exp + 1 : 1;
  if (intDigits > room) return null;
  const decimals = Math.max(0, room - intDigits - 1);
  return v.exp + 1 + decimals;
}

/** Fixed notation with `keep` significant digits kept (from fixedRoom), trailing zeros dropped; null if rounding
 *  carried it past `room` characters (99999999999.5 in 11). */
function fixedText(v, keep, room) {
  let { digits, carried } = roundTo(v.digits, keep);
  let exp = v.exp + (carried ? 1 : 0);
  if (!digits.replace(/0/g, '')) return '0';
  // (Place the point: `digits` stand for 0.d1d2... x 10^(exp+1).)
  let s;
  if (exp < 0) s = '0.' + '0'.repeat(-exp - 1) + digits;
  else if (digits.length <= exp + 1) s = digits + '0'.repeat(exp + 1 - digits.length);
  else s = digits.slice(0, exp + 1) + '.' + digits.slice(exp + 1);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s.length > room ? null : s;
}

/** Scientific notation ("1.23E+08", "1E-300") with as many digits as fit in `room` characters; null if none fit. */
function sciText(v, room) {
  for (let keep = Math.max(1, room - 5); keep >= 1; keep--) {
    const { digits, carried } = roundTo(v.digits, keep);
    const exp = v.exp + (carried ? 1 : 0);
    const d = digits.replace(/0+$/, '') || '0';
    const s = `${d[0]}${d.length > 1 ? '.' + d.slice(1) : ''}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
    if (s.length <= room) return { s, keep };
  }
  return null;
}

/** A number in the General format, in a cell `width` characters wide (units.js colChars; Infinity for no limit).
 *  Fitted on Excel's PDF of test/fixtures/general.xlsx (43 values in columns of 30, 8.43, 6, 4 and 2.5 characters,
 *  all 215 as Excel shows them): at most 11 characters plus a minus sign, and no more than the cell's width. Fixed
 *  notation when it shows the value exactly; otherwise whichever form has room for more significant digits, fixed on
 *  a tie (0.000012345 stays fixed, 0.0000123456789 becomes 1.23457E-05; in 6 characters 0.00012345678 is 0.0001 but
 *  0.00005 is 5E-05). When neither fits, the cell is filled with #. Rounding keeps the sign: -0.12 in two characters
 *  is "-0", as in Excel. */
export function formatGeneral(n, width = Infinity) {
  if (!Number.isFinite(n)) return '#NUM!';
  if (n === 0) return width >= 1 ? '0' : '';
  const sign = n < 0 ? '-' : '';
  const room = Math.min(GENERAL_MAX, width - sign.length);
  const v = decimal(Math.abs(n));
  const fRoom = room >= 1 ? fixedRoom(v, room) : null;
  const f = fRoom != null ? fixedText(v, fRoom, room) : null;
  const e = room >= 1 ? sciText(v, room) : null;
  const exact = f != null && v.digits.length <= fRoom;
  if (f != null && (exact || !e || fRoom >= e.keep)) return sign + f;
  if (e) return sign + e.s;
  return Number.isFinite(width) ? '#'.repeat(Math.max(0, width)) : sign + Math.abs(n);
}

/** What a cell shows, from what the reader stored ({ t, v }), in a cell `width` characters wide, in format `code`
 *  (format.js; General when none): { text, parts, color, fit }. `fit`: a formatted number, which the view replaces
 *  with #s when it is wider than its cell (General fits itself to the width). `locale` (locale.js; the page's unless
 *  given) gives the separators a number is shown with, and the short date built-in format 14 shows. */
export function cellDisplay(cell, width = Infinity, code = 'General', date1904 = false, locale = currentLocale()) {
  if (!cell || cell.v == null) return { text: '', parts: [], color: null, fit: false };
  const plain = (text) => ({ text, parts: [text], color: null, fit: false });
  if (cell.t === 'b') return plain(cell.v ? 'TRUE' : 'FALSE');
  if (cell.t === 'e') return plain(String(cell.v));
  if (cell.t === 'n') {
    if (!code || code === 'General') return plain(localDecimal(formatGeneral(cell.v, width), locale));
    // A number in a text format (@) is shown as General, and to the left as text is (Excel's PDF of numfmt.xlsx).
    if (code === '@') return { ...plain(localDecimal(formatGeneral(cell.v, width), locale)), textFormat: true };
    const r = formatValue(cell.v, localCode(code, locale), { date1904, width, locale });
    if (r.hashes) return { text: '#'.repeat(Number.isFinite(width) ? width : 1), parts: ['#'.repeat(Number.isFinite(width) ? width : 1)], color: r.color, fit: false, hashes: true };
    return { ...r, fit: true };
  }
  if (!code || code === 'General') return plain(String(cell.v));
  // (A text format, @, leaves text as it is.)
  return { ...formatValue(String(cell.v), code, { date1904 }), fit: false };
}

/** The text a cell shows (cellDisplay's, padding as spaces). */
export const cellText = (cell, width = Infinity, code = 'General', date1904 = false, locale = currentLocale()) => cellDisplay(cell, width, code, date1904, locale).text;
