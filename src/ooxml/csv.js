// A .csv file opened as Excel opens one (measured over COM, Workbooks.Open of a CSV, each cell's value, format and
// formula read back: docs/CSV.md): one sheet named after the file, each field taken as if it were typed into its
// cell (entry.js: 007 a number, 50% a percentage, $1,234.50 currency, 1/2 a date, =1+2 a formula, TRUE a boolean,
// #N/A an error), quotes letting a field hold commas, line breaks and "" for a quote. And a sheet saved as a CSV as
// Excel's Save As CSV (UTF-8) and CSV write one.
import { zipSync, strToU8 } from 'fflate';
import { Workbook } from './workbook.js';
import { decodeText, splitDelimited, detectSeparator, importCells } from './textimport.js';
import { currentLocale } from './locale.js';
import { cellText } from './numfmt.js';
import { formatCodeOf } from './styles.js';
import { ansiCode } from '../calc/charsets.js';

const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The fields of a CSV text, row by row: a field in quotes may hold the separator, line breaks (kept as they are,
 *  CR LF in a file from Windows) and "" for a quote; outside quotes a line ends at CR LF, LF or CR. A last empty line
 *  isn't a row. (textimport.js splitDelimited, with one separator.) */
export const parseCsv = (text, sep = ',') => splitDelimited(text, { seps: sep, quote: '"' });

/** Excel's sheet name for a file: its name without the extension, cut to 31 characters, the characters a sheet's
 *  name can't hold (\ / ? * [ ] :) made _. */
export function sheetNameOf(fileName) {
  const base = String(fileName || 'Sheet1').replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '') || 'Sheet1';
  return base.replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
}

/** An empty workbook of one sheet named `name`, as the parts every workbook needs. */
function blankBytes(name) {
  const files = {
    '[Content_Types].xml': `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `${DECL}<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><bookViews><workbookView/></bookViews><sheets><sheet name="${esc(name)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `${DECL}<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': `${DECL}<styleSheet xmlns="${X}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  };
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
}

/** An empty workbook of one sheet named `name`. */
export const blankWorkbook = (name = 'Sheet1') => Workbook.fromBytes(blankBytes(name));

/** A workbook from a CSV file's bytes and its name, as Excel opens it. `encoding`: textimport.js decodeText's ('auto':
 *  a byte order mark, else UTF-8 when the bytes are, else Windows-1252); `sep`: the field separator, else the one the
 *  file seems written with (the locale's list separator on a tie: Excel takes the computer's); `locale` (locale.js):
 *  the numbers' separators and the dates' order, the page's unless given. */
export function csvWorkbook(bytes, fileName, today = new Date(), { locale = currentLocale(), sep = null, encoding = 'auto', quote = '"' } = {}) {
  const { text } = decodeText(bytes, encoding);
  const wb = blankWorkbook(sheetNameOf(fileName));
  const rows = splitDelimited(text, { seps: sep ?? detectSeparator(text, { quote, prefer: locale.listSep }), quote });
  const { contents, styles } = importCells(wb, 0, 0, 0, rows, { locale, today });
  if (contents.length) wb.apply(0, contents);
  if (styles.length) wb.setStyles(0, styles);
  wb.edited = false;
  return wb;
}

/** Whether bytes are a CSV rather than a workbook: not a zip (PK) or an encrypted workbook's compound file. */
export const isCsv = (bytes, fileName) => /\.csv$/i.test(fileName || '') && !(bytes[0] === 0x50 && bytes[1] === 0x4b) && !(bytes[0] === 0xd0 && bytes[1] === 0xcf);

const TAB = String.fromCharCode(9);

/** A field as Excel writes it: in quotes when it holds a comma, a quote, a tab, a line break or the separator, its
 *  quotes doubled (a tab cell quoted in a CSV, a comma one in a tab-delimited file: Excel's saves of
 *  test/savetext-sheet.js). */
const field = (text, sep) => (/[",\t\r\n]/.test(text) || text.includes(sep) ? `"${text.replace(/"/g, '""')}"` : text);

/** Text in an encoding: 'utf-8' (with a byte order mark when `bom`), 'utf-16le' (with one when `bom`) or
 *  'windows-1252' (? for a character it hasn't, as Excel writes one). */
export function encodeText(text, encoding = 'utf-8', bom = true) {
  if (encoding === 'utf-16le') {
    const out = new Uint8Array((bom ? 2 : 0) + text.length * 2);
    let k = 0;
    if (bom) { out[k++] = 0xff; out[k++] = 0xfe; }
    for (let i = 0; i < text.length; i++) { const u = text.charCodeAt(i); out[k++] = u & 255; out[k++] = u >> 8; }
    return out;
  }
  if (encoding === 'windows-1252') return Uint8Array.from([...text].map((ch) => ansiCode(ch)));
  const body = new TextEncoder().encode(text);
  return bom ? new Uint8Array([0xef, 0xbb, 0xbf, ...body]) : body;
}

/** The kinds of text file Save As writes (Excel's names): separator and encoding. */
export const TEXT_KINDS = {
  'csv-utf8': { sep: ',', encoding: 'utf-8', bom: true },
  csv: { sep: ',', encoding: 'windows-1252', bom: false },
  txt: { sep: TAB, encoding: 'windows-1252', bom: false },
  unicode: { sep: TAB, encoding: 'utf-16le', bom: true },
};

/** Sheet si as Excel's Save As writes a CSV (measured over COM, SaveAs CSV and CSV UTF-8): each cell as it shows
 *  ("1,234.50", "25%", "3/5/2024", "(5.00)"; a General number in 11 characters, 1/3 as 0.333333333, whatever its
 *  column's width; '007 as 007), the used range's rows from its first column, each as wide as the used range, rows
 *  ending in CR LF (a line break in a cell kept as it is). `utf8`: UTF-8 with a byte order mark, as CSV UTF-8; else
 *  code page 1252, ? for what it hasn't, as CSV. `kind` (TEXT_KINDS: 'csv-utf8', 'csv', 'txt' as Text (Tab
 *  delimited), 'unicode' as Unicode Text), or `sep`, `encoding` and `bom` given, over it. Cells show as the page's
 *  locale shows them (locale.js), as Excel's CSV has the computer's separators. */
export function csvBytes(wb, si, { utf8 = true, kind = null, sep = null, encoding = null, bom = null } = {}) {
  const base = TEXT_KINDS[kind || (utf8 ? 'csv-utf8' : 'csv')];
  const how = { sep: sep ?? base.sep, encoding: encoding ?? base.encoding, bom: bom ?? base.bom };
  return encodeText(sheetText(wb, si, how.sep), how.encoding, how.bom);
}

/** Sheet si's used range as delimited text (csvBytes' rules), rows ending in CR LF. */
export function sheetText(wb, si, sep = ',') {
  const sh = wb.sheet(si);
  let c0 = Infinity;
  for (const line of sh.cells.values()) for (const [c, cell] of line) if (cell.v != null && c < c0) c0 = c;
  const lines = [];
  if (Number.isFinite(c0)) {
    let r0 = Infinity;
    for (const [r, line] of sh.cells) if ([...line.values()].some((x) => x.v != null)) r0 = Math.min(r0, r);
    for (let r = r0; r <= sh.maxRow; r++) {
      const out = [];
      for (let c = c0; c <= sh.maxCol; c++) {
        const cell = sh.cells.get(r)?.get(c);
        const code = formatCodeOf(wb.styles, cell?.s || 0);
        out.push(field(cell ? cellText(cell, !code || code === 'General' ? 11 : Infinity, code, wb.date1904) : '', sep));
      }
      lines.push(out.join(sep));
    }
  }
  return lines.map((x) => `${x}\r\n`).join('');
}
