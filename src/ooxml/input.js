// What a person types into a cell, as Excel (United States settings) reads it: measured over COM on 136 entries
// (docs/ENTRY.md). Numbers with thousands separators, currency, percentages, exponents and fractions; dates, times and
// both; each with the number format Excel gives the cell when it was General. An entry starting with = is a formula, one
// starting with an apostrophe text.
import { normalizeFormula, displayFormula } from '../calc/fileform.js';
import { parseFormula } from '../calc/parse.js';
import { BUILTIN_FORMATS as B, formatValue, isDateFormat } from './format.js';

const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

/** A month's number from its name: the full name or its first three letters (or more), and "Sept"; else 0. */
function monthOf(word) {
  const w = word.toLowerCase();
  if (w === 'sept') return 9;
  if (w.length < 3) return 0;
  const i = MONTHS.findIndex((m) => m.startsWith(w));
  return i + 1;
}

/** Days in a month (1900 is a leap year in Excel's calendar, its old Lotus mistake). */
const daysIn = (y, m) => (m === 2 && y === 1900 ? 29 : new Date(Date.UTC(y, m, 0)).getUTCDate());
/** A two-digit year as Excel reads it: 00 to 29 are 2000 to 2029, 30 to 99 are 1930 to 1999. */
export const fullYear = (text) => (text.length <= 2 ? (Number(text) < 30 ? 2000 : 1900) + Number(text) : text.length === 4 ? Number(text) : NaN);

/** A date's serial number (1900 system: 1 is 1/1/1900, and the 29th of February 1900 that never was is 60; 1904
 *  system: 0 is 1/1/1904). Null when it isn't a date Excel has (1900 or 1904 to 9999). */
export function dateSerial(y, m, d, date1904 = false) {
  if (!(y >= (date1904 ? 1904 : 1900) && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysIn(y, m))) return null;
  if (date1904) return (Date.UTC(y, m - 1, d) - Date.UTC(1904, 0, 1)) / 864e5;
  if (y === 1900 && m === 2 && d === 29) return 60;
  const days = (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 31)) / 864e5;
  return y > 1900 || m > 2 ? days + 1 : days;
}

/** A day of a month, or else (as Excel reads "2/29" in a year without one, or "Jan-32") the 1st of that month in a
 *  year: [y, m, d, format]. */
function dayOrYear(m, text, thisYear) {
  const n = Number(text);
  if (text.length <= 2 && n >= 1 && n <= daysIn(thisYear, m)) return [thisYear, m, n, B[16]];
  const y = fullYear(text);
  return Number.isNaN(y) ? null : [y, m, 1, B[17]];
}

/** A date typed alone: [y, m, d, format] or null. */
function readDate(s, thisYear) {
  let m = /^(\d{1,4})([/-])(\d{1,4})(?:\2(\d{1,4}))?$/.exec(s);
  if (m) {
    if (m[4] != null) {
      if (m[3].length > 2) return null;
      const [y, mo, d] = m[1].length === 4 ? [Number(m[1]), Number(m[3]), Number(m[4])] : [fullYear(m[4]), Number(m[1]), Number(m[3])];
      if (m[1].length === 4 && m[4].length > 2) return null;
      return Number.isNaN(y) ? null : [y, mo, d, B[14]];
    }
    if (m[1].length > 2) return null;
    const mo = Number(m[1]);
    return mo >= 1 && mo <= 12 ? dayOrYear(mo, m[3], thisYear) : null;
  }
  const word = '([A-Za-z]{3,9})';
  if ((m = new RegExp(`^${word}[\\s-]+(\\d{1,4})$`).exec(s)) && monthOf(m[1])) return dayOrYear(monthOf(m[1]), m[2], thisYear);
  if ((m = new RegExp(`^(\\d{1,2})[\\s-]+${word}$`).exec(s)) && monthOf(m[2])) return [thisYear, monthOf(m[2]), Number(m[1]), B[16]];
  if ((m = new RegExp(`^${word}\\s+(\\d{1,2}),\\s*(\\d{2}|\\d{4})$`).exec(s)) && monthOf(m[1])) return [fullYear(m[3]), monthOf(m[1]), Number(m[2]), B[15]];
  if ((m = new RegExp(`^(\\d{1,2})[\\s-]+${word}[\\s-]+(\\d{2}|\\d{4})$`).exec(s)) && monthOf(m[2])) return [fullYear(m[3]), monthOf(m[2]), Number(m[1]), B[15]];
  return null;
}

/** A time typed alone: { value (a fraction of a day, more past 24 hours), format } or null. Minutes or seconds past
 *  59 still make a number, but General, as Excel leaves it; AM and PM want a space before them. */
function readTime(s) {
  const ap = /^(.*?)\s+(am|pm)$/i.exec(s);
  const body = ap ? ap[1] : s;
  if (ap && /^\d{1,2}$/.test(body)) {
    const h = Number(body);
    if (h > 12) return null;
    return { value: ((h % 12) + (/pm/i.test(ap[2]) ? 12 : 0)) / 24, format: B[18] };
  }
  const m = /^(\d+):(\d+)(?::(\d+))?(\.\d+)?$/.exec(body);
  if (!m) return null;
  const [a, b] = [Number(m[1]), Number(m[2])];
  const c = m[3] != null ? Number(m[3]) : null;
  const frac = m[4] ? Number(m[4]) : 0;
  if (m[4] && c == null) {
    // ("10:30.5": minutes and seconds with a fraction.)
    return { value: (a * 60 + b + frac) / 86400, format: b > 59 ? 'General' : B[47] };
  }
  if (ap) {
    if (a > 12 || b > 59 || (c ?? 0) > 59) return null;
    const h = (a % 12) + (/pm/i.test(ap[2]) ? 12 : 0);
    return { value: (h * 3600 + b * 60 + (c ?? 0) + frac) / 86400, format: c == null ? B[18] : B[19] };
  }
  const value = (a * 3600 + b * 60 + (c ?? 0) + frac) / 86400;
  if (b > 59 || (c ?? 0) > 59) return { value, format: 'General' };
  if (m[4]) return { value, format: B[47] };
  if (a >= 24) return { value, format: B[46] };
  return { value, format: c == null ? B[20] : B[21] };
}

/** A number typed with its decoration: { value, format } or null. */
function readNumber(s) {
  let t = s;
  let neg = false;
  let paren = false;
  let p = /^\((.*)\)$/.exec(t);
  if (p) { paren = true; neg = true; t = p[1].trim(); }
  let money = '';
  let pct = false;
  const sign = () => { const m = /^([+-])\s*/.exec(t); if (!m) return; if (paren) throw new Error('two signs'); if (m[1] === '-') neg = !neg; t = t.slice(m[0].length); };
  try {
    sign();
    if ((p = /^([$€])\s*/.exec(t))) { money = p[1]; t = t.slice(p[0].length); sign(); }
  } catch {
    return null;
  }
  if (t.startsWith('%')) { pct = true; t = t.slice(1); } else if (t.endsWith('%')) { pct = true; t = t.slice(0, -1); }
  // (A comma wants at least three digits after it: "1,2345" is 12345, "1,23" isn't a number.)
  const m = /^(\d+(?:,\d{3,})*)?(\.\d*)?(?:[eE]([+-]?\d+))?$/.exec(t);
  if (!m || !(m[1] || (m[2] && m[2].length > 1))) return null;
  let value = Number(`${(m[1] || '0').replace(/,/g, '')}${m[2] || ''}${m[3] ? `e${m[3]}` : ''}`);
  if (!Number.isFinite(value)) return null;
  if (pct) value /= 100;
  if (neg) value = -value;
  const decimals = !!(m[2] && m[2].length > 1);
  let format = 'General';
  if (money === '$') format = decimals ? B[8] : B[6];
  else if (pct) format = decimals ? B[10] : B[9];
  else if (m[3]) format = B[11];
  else if (m[1] && m[1].includes(',')) format = decimals ? B[4] : B[3];
  return { value, format };
}

/** The error values Excel reads when typed (test/entry-excel.json). */
const TYPED_ERRORS = /^(#NULL!|#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A|#GETTING_DATA)$/i;

/**
 * A typed entry as Excel reads it: { value, format?, kind? } (a number, boolean, string, an error as { code }, or null
 * for an empty entry; `format` the number format Excel gives a General cell, `kind` 'date', 'time', 'datetime',
 * 'number' or 'error'), { formula }
 * (as the file has it, without "=": normalizeFormula) or { error } when it isn't a formula Excel would take. `today`
 * gives the year of a date typed without one; `date1904`, the workbook's date system; `asText`, a cell formatted as
 * Text (@), where Excel keeps whatever is typed as it is (a formula too).
 */
export function parseInput(text, { today = new Date(), date1904 = false, asText = false } = {}) {
  const s = String(text ?? '');
  if (s === '') return { value: null };
  if (asText) return { value: s };
  if (s.startsWith("'")) return { value: s.slice(1) };
  if (s.startsWith('=')) {
    try {
      const formula = normalizeFormula(s.slice(1));
      parseFormula(formula);
      // (=@A1:A3, all @, is kept as the old kind of formula, which doesn't spill: as Excel saves it.)
      return /^\s*@/.test(s.slice(1)) && !/^_xlfn\.SINGLE\(/i.test(formula) ? { formula, legacy: true } : { formula };
    } catch {
      return { error: "There's a problem with this formula" };
    }
  }
  // (An error value typed as Excel names it, in any case, is that error; with a space around it, or #SPILL! and
  // #CALC!, it stays text, as Excel was seen to keep it.)
  if (TYPED_ERRORS.test(s)) return { value: { code: s.toUpperCase() }, kind: 'error' };
  const t = s.trim();
  if (/^(true|false)$/i.test(t)) return { value: t.toLowerCase() === 'true' };
  const n = readNumber(t);
  if (n) return { value: n.value, format: n.format === 'General' ? undefined : n.format, kind: 'number' };
  const frac = /^([+-]?)(\d+)\s+(\d+)\/(\d+)$/.exec(t);
  if (frac && Number(frac[4]) > 0) {
    const v = Number(frac[2]) + Number(frac[3]) / Number(frac[4]);
    return { value: frac[1] === '-' ? -v : v, format: Number(frac[4]) < 10 ? B[12] : B[13], kind: 'number' };
  }
  const thisYear = today.getFullYear();
  const date = readDate(t, thisYear);
  if (date) {
    const serial = dateSerial(date[0], date[1], date[2], date1904);
    return serial == null ? { value: s } : { value: serial, format: date[3], kind: 'date' };
  }
  const time = readTime(t);
  if (time) return { value: time.value, format: time.format === 'General' ? undefined : time.format, kind: 'time' };
  // A date and a time: m/d/yyyy h:mm, whatever the time's own form (General when the time is past 24 hours).
  const both = /^(\S+)\s+(.+)$/.exec(t);
  if (both) {
    const d = readDate(both[1], thisYear);
    const tm = d && (/:/.test(both[2]) || /\s(am|pm)$/i.test(both[2])) && readTime(both[2]);
    const serial = tm && dateSerial(d[0], d[1], d[2], date1904);
    if (serial != null && tm) return { value: serial + tm.value, format: tm.value >= 1 || tm.format === 'General' ? undefined : B[22], kind: 'datetime' };
  }
  return { value: s };
}

const PERCENT = /%(?=(?:[^"]*"[^"]*")*[^"]*$)/;
/** What a format code shows of a date and a time (outside quotes and brackets). */
function parts(code) {
  const bare = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return { date: /[yd]|mmm/i.test(bare), time: /[hs]|AM\/PM/i.test(bare) };
}

/**
 * The format a cell has after an entry: the entry's own (parseInput's `format`) when the cell was General, and in the
 * cases Excel was seen to change it otherwise (a date into a percentage or a time-only format, a time into a
 * date-only one); else null (the cell keeps its format: a date typed into 0.00 stays 0.00).
 */
export function formatAfterEntry(current, parsed) {
  if (!parsed.format) return null;
  if (!current || current === 'General') return parsed.format;
  const pct = PERCENT.test(current) && !isDateFormat(current);
  const shows = isDateFormat(current) ? parts(current) : null;
  if (parsed.kind === 'date' || parsed.kind === 'datetime') return pct || (shows && !shows.date) ? parsed.format : null;
  if (parsed.kind === 'time') return pct || (shows && !shows.time) ? parsed.format : null;
  return null;
}

/**
 * A cell's value as the text to edit it from, as Excel's formula bar has it: a formula with its "=", a date as m/d/yyyy
 * (with its time as h:mm:ss AM/PM), a time alone as h:mm:ss AM/PM, a percentage as a percentage, other numbers in full,
 * booleans as TRUE and FALSE. Typed back, each gives the same value.
 */
export function editText(cell, code = 'General', date1904 = false, at = null) {
  if (!cell) return '';
  if (cell.f) {
    // (A formula of the old kind works by implicit intersection: Excel 365 shows an @ where it cuts a range to one value.
    // `at` puts them in (calc/legacyat.js, which needs the cell's place); without it, a formula that is a whole range,
    // A1:A3, gets its @.)
    const old = !cell.dynamic && !cell.array;
    if (old && at) return `=${at(displayFormula(cell.f))}`;
    const whole = old && /^('[^']+'|[^!(]+!)?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(cell.f);
    return `=${whole ? '@' : ''}${displayFormula(cell.f)}`;
  }
  if (cell.v == null) return '';
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (cell.t !== 'n') return String(cell.v);
  if (code && code !== 'General' && isDateFormat(code) && cell.v >= 0) {
    const whole = Math.floor(cell.v);
    const frac = cell.v - whole;
    const time = formatValue(cell.v, 'h:mm:ss AM/PM', { date1904 }).text;
    if (whole === 0 && !date1904) return time;
    const date = formatValue(whole, 'm/d/yyyy', { date1904 }).text;
    return frac ? `${date} ${time}` : date;
  }
  if (code && PERCENT.test(code)) return `${Number((cell.v * 100).toPrecision(15))}%`;
  return String(cell.v);
}
