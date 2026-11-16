// Text files into cells, as Excel's Text Import Wizard reads them (measured over COM: Workbooks.OpenText of eleven
// made-up files, test/textimport-excel.json, docs/CSV.md): the encoding (a byte order mark, else UTF-8 when the bytes
// are valid UTF-8, else Windows-1252), delimited fields (any of several delimiter characters, a text qualifier around a
// field holding them, qualifiers doubled inside it, consecutive delimiters as one) or fixed-width columns, the first
// row to take, and each column's type: General (as if typed, with the decimal and thousands separators given), Text
// (as it is, formatted @), a date in a day, month and year order (anything else in it stays text), or skipped.
import { entryChanges } from './entry.js';
import { dateSerial, fullYear } from './input.js';
import { BUILTIN_FORMATS as B } from './format.js';
import { currentLocale, localeOf } from './locale.js';

const TAB = String.fromCharCode(9);

/** Bytes as text: { text, encoding } ('utf-8', 'utf-16le', 'utf-16be' or 'windows-1252'). `encoding` 'auto' reads a
 *  byte order mark, then takes UTF-16 when every other byte is 0, UTF-8 when the bytes are valid UTF-8, and
 *  Windows-1252 (Excel's code page for a file without a mark here) otherwise. A byte order mark isn't part of the text. */
export function decodeText(bytes, encoding = 'auto') {
  let enc = encoding;
  let skip = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) { skip = enc === 'auto' || enc === 'utf-8' ? 3 : 0; if (enc === 'auto') enc = 'utf-8'; }
  else if (bytes[0] === 0xff && bytes[1] === 0xfe) { skip = enc === 'auto' || enc === 'utf-16le' ? 2 : 0; if (enc === 'auto') enc = 'utf-16le'; }
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) { skip = enc === 'auto' || enc === 'utf-16be' ? 2 : 0; if (enc === 'auto') enc = 'utf-16be'; }
  if (enc === 'auto') {
    const n = Math.min(bytes.length, 400) & ~1;
    let odd = 0;
    let even = 0;
    for (let i = 0; i < n; i += 2) { if (bytes[i] === 0) even++; if (bytes[i + 1] === 0) odd++; }
    if (n >= 4 && odd >= n / 4 && even === 0) enc = 'utf-16le';
    else if (n >= 4 && even >= n / 4 && odd === 0) enc = 'utf-16be';
    else {
      try { new TextDecoder('utf-8', { fatal: true }).decode(bytes); enc = 'utf-8'; } catch { enc = 'windows-1252'; }
    }
  }
  return { text: new TextDecoder(enc).decode(bytes.subarray(skip)), encoding: enc };
}

/**
 * The fields of delimited text, row by row. `seps`: the delimiter characters (',' for a CSV; ';', TAB, '|', ' ' or
 * several at once); `quote`: the text qualifier ('"', "'" or '' for none): a field in it may hold delimiters and line
 * breaks, the qualifier doubled for one, as Excel reads them ('it''s' is it's; with no qualifier "b,c" is two fields,
 * "b and c", each keeping its quote); `consecutive`: delimiters side by side are one (a delimiter at the start of a line
 * still makes an empty first field, as Excel's wizard was seen to keep). A last empty line isn't a row.
 */
export function splitDelimited(text, { seps = ',', quote = '"', consecutive = false } = {}) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let lastSep = false;
  let i = 0;
  const end = () => { row.push(field); field = ''; };
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === quote && text[i + 1] === quote) { field += quote; i += 2; continue; }
      if (ch === quote) { quoted = false; i++; continue; }
      field += ch;
      i++;
      continue;
    }
    if (quote && ch === quote) { quoted = true; lastSep = false; i++; continue; }
    if (seps.includes(ch)) {
      if (!(consecutive && lastSep)) end();
      lastSep = true;
      i++;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      end();
      rows.push(row);
      row = [];
      lastSep = false;
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    lastSep = false;
    field += ch;
    i++;
  }
  if (field !== '' || row.length) { end(); rows.push(row); }
  return rows;
}

/** The lines of a text (CR LF, LF or CR), a last empty one left out. */
export const linesOf = (text) => {
  const lines = text.split(/\r\n|\n|\r/);
  if (lines.length && lines[lines.length - 1] === '') lines.pop();
  return lines;
};

/** Fixed-width fields: each line cut where `breaks` say a column starts (character positions, in order), each field
 *  without the spaces around it, as Excel's wizard trims them. */
export function splitFixed(text, breaks) {
  const at = [0, ...[...breaks].sort((a, b) => a - b)];
  return linesOf(text).map((line) => at.map((p, k) => line.slice(p, k + 1 < at.length ? at[k + 1] : undefined).trim()));
}

/** Where columns of fixed-width text seem to start: after a run of positions that are a space (or past the end) on
 *  every line, as Excel's wizard first places its break lines. */
export function guessBreaks(text, maxLines = 20) {
  const lines = linesOf(text).slice(0, maxLines).filter((l) => l.trim());
  if (!lines.length) return [];
  const width = Math.max(...lines.map((l) => l.length));
  const blank = (i) => lines.every((l) => i >= l.length || l[i] === ' ');
  const out = [];
  for (let i = 1; i < width; i++) if (blank(i - 1) && !blank(i)) out.push(i);
  return out;
}

/** Which of comma, semicolon, tab and pipe splits most lines into the same number of fields (more than one), the
 *  first of `prefer` on a tie: the delimiter a CSV or text file was most likely written with. */
export function detectSeparator(text, { quote = '"', prefer = ',' } = {}) {
  const sample = text.slice(0, 64 * 1024);
  const order = [...new Set([prefer, ',', ';', TAB, '|'])];
  let best = prefer;
  let score = 0;
  for (const sep of order) {
    const rows = splitDelimited(sample, { seps: sep, quote }).slice(0, 50);
    if (!rows.length) continue;
    const counts = new Map();
    for (const r of rows) if (r.length > 1) counts.set(r.length, (counts.get(r.length) || 0) + 1);
    const top = Math.max(0, ...counts.values());
    // (Lines agreeing on a field count, more fields breaking ties.)
    const s = top * 1000 + Math.min(999, [...counts].find(([, v]) => v === top)?.[0] || 0);
    if (s > score) { score = s; best = sep; }
  }
  return best;
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

/** A field of a date column in `order` ('MDY', 'DMY' or 'YMD'): { value, format } or null (it stays text). Numbers
 *  split by /, - or . (2024-03-05, 24/3/5 and 2024.3.5 in YMD; 05.03.2024, 5-3-24 in DMY); a month by name too; without
 *  a year, this year, shown d-mmm. A four-digit year first is year, month, day in any order (2024-03-05 in MDY); a day
 *  that month hasn't is text (31/2/2024), as Excel leaves it. */
export function dateField(field, order, { today = new Date(), date1904 = false } = {}) {
  const t = field.trim();
  const m = /^(\d{1,4}|[A-Za-z]{3,9})[/.\s-]+(\d{1,4}|[A-Za-z]{3,9})(?:[/.\s-]+(\d{1,4}))?$/.exec(t);
  if (!m) return null;
  const month = (s) => (/^\d+$/.test(s) ? Number(s) : MONTHS.indexOf(s.slice(0, 3).toLowerCase()) + 1);
  let y; let mo; let d;
  const thisYear = today.getFullYear();
  if (m[3] == null) {
    // (A day and a month, this year: the order without its year.)
    const [a, b] = order === 'DMY' ? [m[2], m[1]] : [m[1], m[2]];
    [y, mo, d] = [thisYear, month(a), Number(b)];
  } else if (m[1].length === 4 || order === 'YMD') [y, mo, d] = [fullYear(m[1]), month(m[2]), Number(m[3])];
  else if (order === 'DMY') [y, mo, d] = [fullYear(m[3]), month(m[2]), Number(m[1])];
  else [y, mo, d] = [fullYear(m[3]), month(m[1]), Number(m[2])];
  if (!Number.isInteger(d) || Number.isNaN(y)) return null;
  const serial = dateSerial(y, mo, d, date1904);
  return serial == null ? null : { value: serial, format: m[3] == null ? B[16] : B[14] };
}

/** "€5" read in a General column: Excel's wizard gives it a euro format (where typing it gives a plain 5). */
const EURO = /^(-?)\s*€\s*(-?)(\d[\d,]*)(\.\d+)?$/;

/**
 * The cells for imported rows put at (r0, c0) of sheet si: { contents, styles, rows, cols } for Workbook.apply and
 * setStyles (rows and cols: how much was filled). `columns`: each column's type ('general', 'text', 'MDY', 'DMY',
 * 'YMD', 'skip'; general past the list); `startRow`: the first row taken (1 for all); `locale`: what a General column
 * is read in (locale.js: its decimal and thousands separators, its date order), with `decimal` and `group` over it,
 * as the wizard's Advanced box sets them.
 */
export function importCells(wb, si, r0, c0, rows, { columns = [], startRow = 1, locale = currentLocale(), decimal, group, today = new Date() } = {}) {
  const loc = decimal || group ? localeOf({ ...locale, tag: locale.tag, decimal: decimal ?? locale.decimal, group: group ?? locale.group }) : locale;
  const contents = [];
  const styles = [];
  let cols = 0;
  const taken = rows.slice(Math.max(0, startRow - 1));
  taken.forEach((fields, dr) => {
    const r = r0 + dr;
    let c = c0;
    fields.forEach((field, k) => {
      const type = columns[k] || 'general';
      if (type === 'skip') return;
      const at = c++;
      cols = Math.max(cols, at - c0 + 1);
      if (field === '') return;
      const s = wb.sheet(si).cells.get(r)?.get(at)?.s || 0;
      if (type === 'text') {
        contents.push({ r, c: at, value: field });
        styles.push({ r, c: at, s: wb.styleWith(s, { numFmt: '@' }) });
        return;
      }
      if (type === 'MDY' || type === 'DMY' || type === 'YMD') {
        const d = dateField(field, type, { today, date1904: wb.date1904 });
        if (!d) { contents.push({ r, c: at, value: field }); return; }
        contents.push({ r, c: at, value: d.value });
        styles.push({ r, c: at, s: wb.styleWith(s, { numFmt: d.format }) });
        return;
      }
      const euro = loc.decimal === '.' && EURO.exec(field.trim());
      if (euro && !(euro[1] && euro[2])) {
        const v = Number(`${euro[3].replace(/,/g, '')}${euro[4] || ''}`) * (euro[1] || euro[2] ? -1 : 1);
        const z = euro[4] ? '#,##0.00' : '#,##0';
        contents.push({ r, c: at, value: v });
        styles.push({ r, c: at, s: wb.styleWith(s, { numFmt: `[$€-2] ${z};[Red]-[$€-2] ${z}` }) });
        return;
      }
      const ch = entryChanges(wb, si, r, at, field, today, loc);
      // (What Excel wouldn't take as a formula is the text itself.)
      if (ch.error) { contents.push({ r, c: at, value: field }); return; }
      contents.push(ch.content);
      if (ch.style) styles.push(ch.style);
    });
  });
  return { contents, styles, rows: taken.length, cols };
}

/** The settings a text file most likely wants, as the Text Import box first shows them: its encoding, delimited by
 *  the delimiter detected (a .tsv or .txt by tab when it has tabs), qualifier ", from the first row, General columns. */
export function guessSettings(bytes, fileName = '', locale = currentLocale()) {
  const { text, encoding } = decodeText(bytes);
  const tabbed = /\.(txt|tsv|tab)$/i.test(fileName) && text.includes(TAB);
  const seps = tabbed ? TAB : detectSeparator(text, { prefer: locale.listSep });
  return { encoding, kind: 'delimited', seps, quote: '"', consecutive: false, breaks: guessBreaks(text), startRow: 1, columns: [], locale: locale.tag, decimal: locale.decimal, group: locale.group };
}

/** The rows a file gives with `settings` (guessSettings' shape), before column types: { text, rows }. */
export function splitText(bytes, settings) {
  const { text } = decodeText(bytes, settings.encoding || 'auto');
  const rows = settings.kind === 'fixed' ? splitFixed(text, settings.breaks || []) : splitDelimited(text, settings);
  return { text, rows };
}
