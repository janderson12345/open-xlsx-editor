// The reader's country settings, as Excel takes them from the computer's (Windows' Region settings): the decimal and
// thousands separators for typing and showing numbers, the order of day, month and year for typing dates and the short
// date that built-in format 14 shows, the list separator for formula arguments and CSV fields. The file is the same
// whatever they are: numbers, dates and formulas are written as en-US has them (ECMA-376 Part 1, 18.8.30: built-in
// format 14 is "the system's short date"; formulas are stored in the invariant form). Like Excel's, the setting belongs
// to the page, not to a workbook: setLocale() makes it the one every editor on the page shows and reads typing with;
// the calculation engine (TEXT, VALUE) keeps en-US, as the file's own values do.

const NBSP = String.fromCharCode(0xa0);
const NNBSP = String.fromCharCode(0x202f);

/** en-US, Excel's settings for the United States, and the default. */
export const EN_US = Object.freeze({
  tag: 'en-US', decimal: '.', group: ',', dateOrder: 'MDY', dateSep: '/', shortDate: 'm/d/yyyy', hour12: true,
  listSep: ',', arrayCol: ',', arrayRow: ';',
});

// Windows 11's defaults for some countries (Region > Regional format): decimal, thousands, date order, date separator,
// short date (as a format code), 12-hour clock. Others come from the browser's Intl (localeFrom).
const KNOWN = {
  'en-US': ['.', ',', 'MDY', '/', 'm/d/yyyy', true],
  'en-GB': ['.', ',', 'DMY', '/', 'dd/mm/yyyy', false],
  'en-AU': ['.', ',', 'DMY', '/', 'd/mm/yyyy', true],
  'en-CA': ['.', ',', 'YMD', '-', 'yyyy-mm-dd', true],
  'en-IN': ['.', ',', 'DMY', '-', 'dd-mm-yyyy', true],
  'en-IE': ['.', ',', 'DMY', '/', 'dd/mm/yyyy', false],
  'en-NZ': ['.', ',', 'DMY', '/', 'd/mm/yyyy', true],
  'en-ZA': ['.', ' ', 'YMD', '/', 'yyyy/mm/dd', false],
  'de-DE': [',', '.', 'DMY', '.', 'dd.mm.yyyy', false],
  'de-AT': [',', NBSP, 'DMY', '.', 'dd.mm.yyyy', false],
  'de-CH': ['.', String.fromCharCode(0x2019), 'DMY', '.', 'dd.mm.yyyy', false],
  'fr-FR': [',', NNBSP, 'DMY', '/', 'dd/mm/yyyy', false],
  'fr-CA': [',', NBSP, 'YMD', '-', 'yyyy-mm-dd', false],
  'fr-BE': [',', NNBSP, 'DMY', '/', 'dd-mm-yy', false],
  'es-ES': [',', '.', 'DMY', '/', 'dd/mm/yyyy', false],
  'es-MX': ['.', ',', 'DMY', '/', 'dd/mm/yyyy', true],
  'it-IT': [',', '.', 'DMY', '/', 'dd/mm/yyyy', false],
  'pt-BR': [',', '.', 'DMY', '/', 'dd/mm/yyyy', false],
  'pt-PT': [',', NBSP, 'DMY', '/', 'dd/mm/yyyy', false],
  'nl-NL': [',', '.', 'DMY', '-', 'd-m-yyyy', false],
  'nl-BE': [',', '.', 'DMY', '/', 'd/mm/yyyy', false],
  'sv-SE': [',', NBSP, 'YMD', '-', 'yyyy-mm-dd', false],
  'da-DK': [',', '.', 'DMY', '.', 'dd.mm.yyyy', false],
  'nb-NO': [',', NBSP, 'DMY', '.', 'dd.mm.yyyy', false],
  'fi-FI': [',', NBSP, 'DMY', '.', 'd.m.yyyy', false],
  'pl-PL': [',', NBSP, 'DMY', '.', 'dd.mm.yyyy', false],
  'cs-CZ': [',', NBSP, 'DMY', '.', 'dd.mm.yyyy', false],
  'ru-RU': [',', NBSP, 'DMY', '.', 'dd.mm.yyyy', false],
  'tr-TR': [',', '.', 'DMY', '.', 'd.mm.yyyy', false],
  'ja-JP': ['.', ',', 'YMD', '/', 'yyyy/mm/dd', false],
  'zh-CN': ['.', ',', 'YMD', '/', 'yyyy/m/d', false],
  'zh-TW': ['.', ',', 'YMD', '/', 'yyyy/m/d', true],
  'ko-KR': ['.', ',', 'YMD', '-', 'yyyy-mm-dd', true],
};

/** A short date code for a date order and separator, as Windows makes one (two-digit day and month but en-US's). */
export function shortDateOf(order, sep) {
  if (order === 'YMD') return `yyyy${sep}mm${sep}dd`;
  if (order === 'DMY') return `dd${sep}mm${sep}yyyy`;
  return `m${sep}d${sep}yyyy`;
}

/** The settings of a locale tag the table above lacks, from the browser's Intl (null if Intl doesn't know it). */
function fromIntl(tag) {
  try {
    const num = new Intl.NumberFormat(tag, { useGrouping: true }).formatToParts(1234567.5);
    const decimal = num.find((p) => p.type === 'decimal')?.value ?? '.';
    const group = num.find((p) => p.type === 'group')?.value ?? (decimal === ',' ? '.' : ',');
    const parts = new Intl.DateTimeFormat(tag, { year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date(2024, 10, 22));
    const order = parts.filter((p) => ['year', 'month', 'day'].includes(p.type)).map((p) => p.type[0].toUpperCase()).join('');
    const dateSep = (parts.find((p) => p.type === 'literal')?.value || '/').trim()[0] || '/';
    const hour12 = new Intl.DateTimeFormat(tag, { hour: 'numeric' }).resolvedOptions().hour12 ?? false;
    const dateOrder = ['MDY', 'DMY', 'YMD'].includes(order) ? order : 'DMY';
    return [decimal, group, dateOrder, dateSep, shortDateOf(dateOrder, dateSep), hour12];
  } catch {
    return null;
  }
}

/** The table's row for a tag: itself, else the first one of its language (fr-LU as fr-FR), else Intl's. */
function rowOf(tag) {
  if (KNOWN[tag]) return [tag, KNOWN[tag]];
  const lang = tag.split('-')[0].toLowerCase();
  const near = Object.keys(KNOWN).find((k) => k.toLowerCase() === tag.toLowerCase());
  if (near) return [near, KNOWN[near]];
  const viaIntl = fromIntl(tag);
  if (viaIntl) return [tag, viaIntl];
  const first = Object.keys(KNOWN).find((k) => k.startsWith(`${lang}-`));
  return first ? [tag, KNOWN[first]] : ['en-US', KNOWN['en-US']];
}

/**
 * A locale's settings from a tag ('de-DE'), or an object with a tag and any settings changed ({ tag: 'en-US',
 * dateOrder: 'DMY' }, { decimal: ',', group: '.' }), as Excel's own Options > Advanced > "Use system separators" and
 * Windows' Region settings let a person change them: { tag, decimal, group, dateOrder ('MDY', 'DMY' or 'YMD'), dateSep,
 * shortDate (built-in format 14's code), hour12, listSep, arrayCol, arrayRow }. The list separator is ; where the
 * decimal is a comma (as Windows sets it), and an array constant's columns are then split by a dot, as German Excel
 * writes {1.2;3.4} for {1,2;3,4} (not measured: Excel here has United States settings only).
 */
export function localeOf(spec) {
  if (spec && typeof spec === 'object' && spec.decimal && spec.listSep && spec.shortDate && spec.tag && Object.isFrozen(spec)) return spec;
  const given = typeof spec === 'string' ? { tag: spec } : { ...(spec || {}) };
  const [tag, row] = rowOf(String(given.tag || 'en-US'));
  const [decimal, group, dateOrder, dateSep, shortDate, hour12] = row;
  const out = { tag, decimal, group, dateOrder, dateSep, shortDate, hour12 };
  for (const k of ['decimal', 'group', 'dateOrder', 'dateSep', 'hour12', 'listSep', 'shortDate']) if (given[k] != null) out[k] = given[k];
  if (!['MDY', 'DMY', 'YMD'].includes(out.dateOrder)) out.dateOrder = 'MDY';
  // (A date order or separator changed without a short date: one made from them.)
  if (given.shortDate == null && (given.dateOrder != null || given.dateSep != null)) out.shortDate = shortDateOf(out.dateOrder, out.dateSep);
  if (out.group === out.decimal) out.group = out.decimal === ',' ? '.' : ',';
  out.listSep = given.listSep ?? (out.decimal === ',' ? ';' : ',');
  out.arrayRow = ';';
  out.arrayCol = out.decimal === ',' ? '.' : ',';
  // (en-US unchanged is EN_US itself, which the readers' quick paths look for.)
  if (Object.keys(EN_US).every((k) => EN_US[k] === out[k])) return EN_US;
  return Object.freeze(out);
}

let active = EN_US;

/** The page's locale now (en-US until setLocale is called). */
export const currentLocale = () => active;

/** Make `spec` (localeOf) the page's locale, as a change of the computer's settings is to Excel; returns it. */
export function setLocale(spec) {
  active = localeOf(spec);
  return active;
}

/** The browser's language, or en-US where there is none (Node). */
export const browserLocale = () => (typeof navigator !== 'undefined' && navigator.language) || 'en-US';

/** Whether a character is a space as a thousands separator can be (a space, no-break space or narrow no-break space). */
export const isSpaceGroup = (ch) => ch === ' ' || ch === NBSP || ch === NNBSP;

/** A format code as the locale shows it: built-in format 14 (m/d/yyyy) and 22 (m/d/yyyy h:mm), "the system's short
 *  date", as the locale's short date; others as they are (their separators are changed as the number is drawn). */
export function localCode(code, locale = active) {
  if (locale === EN_US || locale.shortDate === 'm/d/yyyy') return code;
  if (code === 'm/d/yyyy') return locale.shortDate;
  if (code === 'm/d/yyyy h:mm') return `${locale.shortDate} h:mm`;
  return code;
}

/** A number's text (en-US digits: "1234.5", "1.5E+20") with the locale's decimal separator. */
export const localDecimal = (text, locale = active) => (locale.decimal === '.' ? text : text.replace('.', locale.decimal));
