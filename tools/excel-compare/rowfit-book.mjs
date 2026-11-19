// A workbook for the row heights Excel works out itself (src/editor/rowheights.js), the same every time: Calibri 11,
// no customHeight anywhere. Row 1 a 16pt bold title merged across A1:H1 with a thick bottom border (thickBot); row 2
// wrapped bold headings, one with a line break; rows 3-62 a wrapped formula's text in column B (19.11 wide: one long
// "word" of 14 to 26 characters, broken where it reaches the edge), a wrapped sentence of 4 to 16 words in column H
// (45 wide), and in rows 55-62 the long words as strings instead. Excel's heights for it: excel-rowfit.ps1, giving
// test/rowfit-excel.json; the browser check is test/rowfit.mjs.
// usage: node tools/excel-compare/rowfit-book.mjs <out.xlsx>
import { zipSync, strToU8 } from 'fflate';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const WORDS = ['analysis', 'of', 'operational', 'aspects', 'for', 'the', 'and', 'design', 'support', 'flow', 'model',
  'systems', 'research', 'testing', 'advanced', 'structural', 'weight', 'estimation', 'program', 'study', 'data',
  'methods', 'vehicle', 'operations', 'integration', 'development', 'safety', 'technology', 'flight', 'wing', 'power',
  'electric', 'materials', 'damage', 'through', 'small', 'large', 'review', 'computational', 'experimental', 'a', 'to',
  'in', 'with', 'on', 'new', 'field', 'sensor', 'performance', 'evaluation'];
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** The workbook's rows: [{ r, b (text), bFormula (true: a formula's text result), h (text) }] for rows 3-62. */
export function rowfitRows() {
  let seed = 20261118;
  const rand = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  const out = [];
  for (let k = 0; k < 60; k++) {
    const r = k + 3;
    const idLen = 14 + (k % 13);
    let id = 'TDN.C2.';
    while (id.length < idLen) id += String(Math.floor(rand() * 10));
    const n = 4 + Math.floor(rand() * 13);
    const words = Array.from({ length: n }, () => WORDS[Math.floor(rand() * WORDS.length)]);
    words[0] = words[0][0].toUpperCase() + words[0].slice(1);
    out.push({ r, b: id, bFormula: r < 55, h: words.join(' ') });
  }
  return out;
}

/** Row 63's text: the word list 40 times over, about 4,700 characters. */
export const LONG_TEXT = Array.from({ length: 40 }, () => WORDS.join(' ')).join(' ');

const SENTENCE = 'Review of advanced structural weight estimation methods for small electric flight vehicle operations and state-of-the-art follow-up testing';
/** Rows 64 on, a rule of rowheights.js each, the text in column H (45 wide) and "Row n" in column A in the Normal
 *  font: { font: [name, size, 'b' | 'i' | 'bi'], wrap, text (SENTENCE when left out), valign, indent, beside (an empty
 *  cell of that font in column G), thickTop, thickBot (the row's attribute and the cell's border), runs (a rich
 *  text: [text, font or null]), bare (no cell at all) }. */
export const MORE_ROWS = [
  { font: ['Calibri', 10], wrap: true }, // 64: a font smaller than the Normal one, a pixel over its lines
  { font: ['Calibri', 10], wrap: true, valign: 'top' }, // 65: not when aligned to the top
  { font: ['Arial', 8], wrap: true },
  { font: ['Arial', 12], text: 'One line of Arial 12' }, // 67: a pixel over its line, not wrapped either
  { font: ['Arial', 12], text: '' }, // 68: an empty formatted cell too
  { font: ['Calibri', 10], wrap: true, beside: ['Calibri', 20] }, // 69: under a deeper font's baseline
  { font: ['Calibri', 11], text: 'Thick top border', thickTop: true },
  { font: ['Calibri', 11], text: 'Thick top and bottom borders', thickTop: true, thickBot: true },
  { font: ['Calibri', 11, 'i'], wrap: true }, // 72: an italic's widths and allowance
  { font: ['Book Antiqua', 10, 'i'], wrap: true },
  { font: ['Arial', 7], wrap: true }, // 74: Small Fonts
  { font: ['Arial', 7, 'b'], wrap: true },
  { font: ['Calibri', 11], wrap: true, text: 'well-known, state-of-the-art e-mail follow-up on the 12/31/2024-01/15/2025 non-recurring re-test of every sub-system' },
  { font: ['Calibri', 11], wrap: true, text: `${SENTENCE}\n \n ` }, // 77: the last line after a line of one space
  { font: ['Calibri', 12, 'b'], wrap: true, runs: [['Quarterly Review Totals ', null], ['(filled in by the sender each month, and checked by the reviewer before it is sent on)', ['Calibri', 10, 'i']]] },
  { font: ['Calibri', 11], wrap: true, indent: 2 },
  { font: ['Helvetica', 12], wrap: true }, // 80: a font Windows hasn't, laid out in Arial
  { font: ['Verdana', 9], wrap: true },
  { font: ['Times New Roman', 10, 'b'], wrap: true, valign: 'center' },
  { bare: true, thickBot: true }, // 83: a thick bottom border and no cell
  { font: ['Calibri', 11], text: 'The last row' },
];

/** The .xlsx's bytes. */
export function rowfitBook() {
  const cells = [];
  const fonts = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>', '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>', '<font><b/><sz val="16"/><name val="Calibri"/><family val="2"/></font>'];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>'];
  const fontXml = ([name, size, style = '']) => `${style.includes('b') ? '<b/>' : ''}${style.includes('i') ? '<i/>' : ''}<sz val="${size}"/>`;
  const add = (list, xml) => (list.includes(xml) ? list.indexOf(xml) : list.push(xml) - 1);
  const fontId = (f) => add(fonts, `<font>${fontXml(f)}<name val="${f[0]}"/><family val="2"/></font>`);
  // (Borders: 0 none, 1 a thick bottom, 2 a thick top, 3 both.)
  const xfOf = (x, font) => add(xfs, `<xf numFmtId="0" fontId="${fontId(font)}" fillId="0" borderId="${(x.thickTop ? 2 : 0) + (x.thickBot ? 1 : 0)}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">`
    + `<alignment${x.valign ? ` vertical="${x.valign}"` : ''}${x.wrap ? ' wrapText="1"' : ''}${x.indent ? ` indent="${x.indent}"` : ''}/></xf>`);
  cells.push('<row r="1" thickBot="1"><c r="A1" s="3" t="inlineStr"><is><t>Row heights as Excel works them out</t></is></c>'
    + ['B', 'C', 'D', 'E', 'F', 'G', 'H'].map((c) => `<c r="${c}1" s="3"/>`).join('') + '</row>');
  cells.push('<row r="2"><c r="A2" s="2" t="inlineStr"><is><t>Name</t></is></c><c r="B2" s="2" t="inlineStr"><is><t>Identifier</t></is></c>'
    + '<c r="C2" s="2" t="inlineStr"><is><t xml:space="preserve">Subtask &#10;(Enter 001, if no subtask)</t></is></c>'
    + '<c r="H2" s="2" t="inlineStr"><is><t>Title</t></is></c></row>');
  for (const x of rowfitRows()) {
    const b = x.bFormula
      ? `<c r="B${x.r}" s="1" t="str"><f>"${esc(x.b.slice(0, 7))}"&amp;"${esc(x.b.slice(7))}"</f><v>${esc(x.b)}</v></c>`
      : `<c r="B${x.r}" s="1" t="inlineStr"><is><t>${esc(x.b)}</t></is></c>`;
    cells.push(`<row r="${x.r}" ht="28.8"><c r="A${x.r}" t="inlineStr"><is><t>Row ${x.r}</t></is></c>${b}<c r="H${x.r}" s="1" t="inlineStr"><is><t>${esc(x.h)}</t></is></c></row>`);
  }
  // (Row 63: a wrapped text of some hundred lines, taller than Excel lets a row be: it stops at 409.5pt.)
  cells.push(`<row r="63"><c r="A63" t="inlineStr"><is><t>Row 63</t></is></c><c r="H63" s="1" t="inlineStr"><is><t>${esc(LONG_TEXT)}</t></is></c></row>`);
  MORE_ROWS.forEach((x, k) => {
    const r = 64 + k;
    const attrs = `${x.thickTop ? ' thickTop="1"' : ''}${x.thickBot ? ' thickBot="1"' : ''}`;
    if (x.bare) { cells.push(`<row r="${r}"${attrs}/>`); return; }
    const text = x.text ?? SENTENCE;
    const is = x.runs
      ? x.runs.map(([t, f]) => `<r>${f ? `<rPr>${fontXml(f)}<rFont val="${f[0]}"/><family val="2"/></rPr>` : ''}<t xml:space="preserve">${esc(t)}</t></r>`).join('')
      : `<t xml:space="preserve">${esc(text).replace(/\n/g, '&#10;')}</t>`;
    const h = text || x.runs ? `<c r="H${r}" s="${xfOf(x, x.font)}" t="inlineStr"><is>${is}</is></c>` : `<c r="H${r}" s="${xfOf(x, x.font)}"/>`;
    const g = x.beside ? `<c r="G${r}" s="${xfOf({}, x.beside)}"/>` : '';
    cells.push(`<row r="${r}"${attrs}><c r="A${r}" t="inlineStr"><is><t>Row ${r}</t></is></c>${g}${h}</row>`);
  });
  const cols = '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="19.109375" customWidth="1"/>'
    + '<col min="3" max="3" width="8.88671875" customWidth="1"/><col min="8" max="8" width="45" customWidth="1"/></cols>';
  const files = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Rows" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + `<fonts count="${fonts.length}">${fonts.join('')}</fonts>`
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
      + '<borders count="4"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color auto="1"/></bottom><diagonal/></border>'
      + '<border><left/><right/><top style="thick"><color auto="1"/></top><bottom/><diagonal/></border><border><left/><right/><top style="thick"><color auto="1"/></top><bottom style="thick"><color auto="1"/></bottom><diagonal/></border></borders>'
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`
      + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="14.4"/>${cols}<sheetData>${cells.join('')}</sheetData><mergeCells count="1"><mergeCell ref="A1:H1"/></mergeCells></worksheet>`,
  };
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  writeFileSync(process.argv[2], rowfitBook());
  console.log(`written ${process.argv[2]}`);
}
