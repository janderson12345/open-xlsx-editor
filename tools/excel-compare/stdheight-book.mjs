// Workbooks for a sheet's standard row height (src/editor/rowheights.js standardRowHeight), the same every time: one
// sheet each, Normal style Calibri 11 (Arial 10 in `arial`), each varying one thing. Rows 1-12 the same in all: 1
// empty, 2 a number in the Normal style, 3 in Calibri 6, 4 in Calibri 20, 5 in Calibri 9 bold, 6 an empty Calibri 6
// cell, 7 a row element with ht="30" and nothing else, 8 a row made Calibri 6 with nothing in it, 9 made Calibri 20,
// 10 made Calibri 6 with a Calibri 6 number, 11 made Calibri 6 with a Normal number, 12 made Calibri 9 bold with a
// Calibri 6 number. Excel's heights for them: excel-stdheight.ps1, giving test/stdheight-excel.json; the browser
// check is test/stdheight.mjs.
// usage: node tools/excel-compare/stdheight-book.mjs <outDir>
import { zipSync, strToU8 } from 'fflate';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const col = (min, max, s, extra = '') => `<col min="${min}" max="${max}" width="9.140625" style="${s}"${extra}/>`;

/** The cases: { name, fp (sheetFormatPr's attributes), cols (<col> elements), arial (Normal Arial 10) }. */
export const STDHEIGHT_CASES = [
  // (Stored by Excel on a screen at 125% and 150%: 14.4 and 14.5 for Calibri 11's 15.)
  { name: 'stored-14.4', fp: 'defaultRowHeight="14.4"' },
  { name: 'stored-14.5', fp: 'defaultRowHeight="14.5"' },
  // (Any stored number without customHeight: Excel works it out.)
  { name: 'stored-12', fp: 'defaultRowHeight="12"' },
  { name: 'stored-20', fp: 'defaultRowHeight="20"' },
  { name: 'arial-stored-13.2', fp: 'defaultRowHeight="13.2"', arial: true },
  { name: 'none', fp: '' },
  // (With customHeight the stored number is the standard height, to the quarter point.)
  { name: 'custom-20', fp: 'defaultRowHeight="20" customHeight="1"' },
  { name: 'custom-10.3', fp: 'defaultRowHeight="10.3" customHeight="1"' },
  // (Column styles: every column, all but one, a few, mixed.)
  { name: 'cols-all-6', fp: 'defaultRowHeight="15"', cols: col(1, 16384, 1) },
  { name: 'cols-all-but-last-6', fp: 'defaultRowHeight="15"', cols: col(1, 16383, 1) },
  { name: 'cols-first-five-20', fp: 'defaultRowHeight="15"', cols: col(1, 5, 2) },
  { name: 'cols-9b-and-6', fp: 'defaultRowHeight="15"', cols: col(1, 3, 3) + col(4, 16384, 1) },
  { name: 'cols-hidden-20', fp: 'defaultRowHeight="15"', cols: col(1, 99, 1) + col(100, 100, 2, ' hidden="1"') + col(101, 16384, 1) },
  { name: 'cols-all-6-custom-20', fp: 'defaultRowHeight="20" customHeight="1"', cols: col(1, 16384, 1) },
];

const ROWS = [
  '<row r="2"><c r="A2"><v>1</v></c></row>',
  '<row r="3"><c r="A3" s="1"><v>1</v></c></row>',
  '<row r="4"><c r="A4" s="2"><v>1</v></c></row>',
  '<row r="5"><c r="A5" s="3"><v>1</v></c></row>',
  '<row r="6"><c r="A6" s="1"/></row>',
  '<row r="7" ht="30"/>',
  '<row r="8" s="1" customFormat="1"/>',
  '<row r="9" s="2" customFormat="1"/>',
  '<row r="10" s="1" customFormat="1"><c r="A10" s="1"><v>1</v></c></row>',
  '<row r="11" s="1" customFormat="1"><c r="A11"><v>1</v></c></row>',
  '<row r="12" s="3" customFormat="1"><c r="A12" s="1"><v>1</v></c></row>',
].join('');

/** One case's .xlsx bytes. */
export function stdheightBook(c) {
  const normal = c.arial ? '<sz val="10"/><name val="Arial"/><family val="2"/>' : '<sz val="11"/><name val="Calibri"/><family val="2"/>';
  const files = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Rows" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + `<fonts count="4"><font>${normal}</font><font><sz val="6"/><name val="Calibri"/><family val="2"/></font><font><sz val="20"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="9"/><name val="Calibri"/><family val="2"/></font></fonts>`
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
      + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
      + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
      + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    'xl/worksheets/sheet1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + `${c.fp ? `<sheetFormatPr ${c.fp}/>` : ''}${c.cols ? `<cols>${c.cols}</cols>` : ''}<sheetData>${ROWS}</sheetData></worksheet>`,
  };
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  mkdirSync(process.argv[2], { recursive: true });
  for (const c of STDHEIGHT_CASES) writeFileSync(join(process.argv[2], `${c.name}.xlsx`), stdheightBook(c));
  console.log(`written ${STDHEIGHT_CASES.length} workbooks to ${process.argv[2]}`);
}
