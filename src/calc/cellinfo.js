// What CELL tells of a cell besides its value, as Excel 365 (English) gives it (measured over COM, docs/CALC.md): the
// format code as Excel's letters ("G", "F2", ",0", "C2-", "P0", "S2", "D1" to "D9", "()" after one with parentheses),
// whether negative numbers are in colour or the positive in parentheses, the text alignment's prefix, the protection and
// the column's width.
import { xfOf, formatCodeOf } from '../ooxml/styles.js';
import { colWidthToChars, defaultColPx, pxToChars } from '../ooxml/units.js';

/** The first section of a format code with quoted text, escaped and padded characters taken out. */
function bare(section) {
  return section.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/[_*]./g, '');
}

/** A format code's sections, split at the semicolons outside quotes. */
function sections(code) {
  const out = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') quoted = !quoted;
    if (ch === '\\' && !quoted) { cur += ch + (code[++i] ?? ''); continue; }
    if (ch === ';' && !quoted) { out.push(cur); cur = ''; } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** A date or time code's letters as Excel's CELL("format") reads them, or null when it isn't one. The order of the
 *  first date parts decides (day-month-year D1, day-month D2, month-year D3, month-day-year D4, month-day D5); a time
 *  is h:mm (D9), h:mm:ss (D8), with AM/PM D7 and D6. A locale prefix ([$-409]) counts as a day, as measured. */
function dateLetters(first) {
  let s = first.replace(/\[\$-[0-9A-Fa-f]+\]/g, '§');
  if (/\[(h+|m+|s+)\]/i.test(s)) return 'G';
  s = s.replace(/\[[^\]]*\]/g, '');
  const ampm = /AM\/PM|A\/P/i.test(s);
  s = s.replace(/AM\/PM|A\/P/gi, '');
  const runs = s.match(/§|d+|m+|y+|h+|s+/gi);
  if (!runs) return null;
  // (m after h or before s is minutes.)
  const kinds = runs.map((r, i) => {
    const k = r === '§' ? 'd' : r[0].toLowerCase();
    if (k !== 'm') return k;
    const prev = runs[i - 1]?.[0].toLowerCase();
    const next = runs[i + 1]?.[0].toLowerCase();
    return prev === 'h' || next === 's' ? 'n' : 'm';
  });
  if (kinds[0] === 'h') {
    const parts = kinds.join('');
    if (parts.startsWith('hns')) return ampm ? 'D6' : 'D8';
    if (parts.startsWith('hn')) return ampm ? 'D7' : 'D9';
    return 'G';
  }
  const [a, b, c] = kinds;
  if (a === 'd' && b === 'm') return c === 'y' ? 'D1' : 'D2';
  if (a === 'm' && b === 'd') return c === 'y' ? 'D4' : 'D5';
  if (a === 'm' && b === 'y') return 'D3';
  return 'G';
}

/** CELL("format"), CELL("color") and CELL("parentheses") of format code `code`: { format, color, parentheses }. */
export function formatLetters(code) {
  const parts = sections(code || 'General');
  const first = bare(parts[0]);
  const color = parts.length > 1 && /\[(red|blue|green|yellow|magenta|cyan|white|black|color\s*\d+)\]/i.test(parts[1]) ? 1 : 0;
  const parentheses = /\(/.test(first.replace(/\[[^\]]*\]/g, '')) ? 1 : 0;
  const tail = (color ? '-' : '') + (parentheses ? '()' : '');
  const plain = first.replace(/\[[^\]]*\]/g, '');
  if (/^\s*general\s*$/i.test(plain) || /^\s*@\s*$/.test(plain)) return { format: 'G' + tail, color, parentheses };
  const date = dateLetters(first);
  if (date) return { format: date === 'G' ? 'G' : date + tail, color, parentheses };
  if (/[?#0]\s*\/\s*[?#0]/.test(plain)) return { format: 'G', color, parentheses };
  const decimals = (plain.match(/\.([0#?]+)/)?.[1] || '').length;
  let letter;
  if (/%/.test(plain)) letter = 'P';
  else if (/E[+-]/i.test(plain)) letter = 'S';
  else if (/\$/.test(first) || /\[\$[^\]-]/.test(parts[0])) letter = 'C';
  else if (/[0#?],[0#?]/.test(plain)) letter = ',';
  else letter = 'F';
  return { format: letter + decimals + tail, color, parentheses };
}

const PREFIX = { left: "'", general: "'", justify: "'", distributed: "'", right: '"', center: '^', centerContinuous: '^', fill: '\\' };

/** What CELL needs of cell (r, c) of sheet `si` of workbook `book`: { code, prefix, locked, width, defaultWidth }. */
export function cellInfo(book, si, r, c) {
  const sheet = book.sheet(si);
  const cell = sheet.cells.get(r)?.get(c);
  const col = sheet.cols.find((x) => c + 1 >= x.min && c + 1 <= x.max);
  const s = cell?.s || col?.style || 0;
  const xf = xfOf(book.styles, s);
  const typedText = typeof cell?.v === 'string' && !cell.f;
  const ownWidth = col && col.width != null && !col.hidden;
  const fmt = sheet.format || {};
  const mdw = fmt.mdw ?? 7;
  const chars = col?.hidden ? 0 : ownWidth ? colWidthToChars(col.width, mdw) : pxToChars(defaultColPx(fmt, mdw), mdw);
  return {
    code: formatCodeOf(book.styles, s),
    prefix: typedText ? PREFIX[xf.align.horizontal] ?? "'" : '',
    locked: xf.protection.locked,
    width: Math.round(chars),
    defaultWidth: !ownWidth,
  };
}
