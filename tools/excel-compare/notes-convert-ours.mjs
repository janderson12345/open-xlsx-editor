// A workbook with notes written here, for Excel's Convert to Comments (excel-notes-visible.ps1 <ours>): notes.xlsx
// (Avery Chen's notes on B2 and D4, a thread on C3) with Sam Rivera's note on A6, a note on F8 whose text has no
// author's name in bold (a plain run), notes in column G whose first line is bold but not the author's, the author's
// name not bold, another's name bold, or with no author, and a second sheet with Sam Rivera's note on B3.
// usage: node tools/excel-compare/notes-convert-ours.mjs <out.xlsx> (or convertInput(), for notesvisible.test.js)
import fs from 'node:fs';
import { pathToFileURL } from 'node:url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { setNote } from '../../src/ooxml/noteedit.js';

/** The workbook (a Workbook), made from test/fixtures/notes.xlsx. */
export function convertInput() {
  setXmlImpl({ DOMParser, XMLSerializer });
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/notes.xlsx')));
  setNote(wb, 0, 5, 0, 'A new one', 'Sam Rivera');
  setNote(wb, 0, 7, 5, 'No name here', 'Sam Rivera');
  // (F8's text a plain run, without the bold name.)
  const c = wb.pkg.text('xl/comments1.xml');
  wb.pkg.setText('xl/comments1.xml', c.replace(/(<comment ref="F8"[^>]*><text>)[\s\S]*?(<\/text>)/, '$1<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>No name here</t></r>$2'));
  // (The first line bold but not the author's name; the author's name in a plain run; another person's name in bold; no
  // author at all.)
  setNote(wb, 0, 1, 6, 'x', 'Avery Chen');
  setNote(wb, 0, 3, 6, 'x', 'Avery Chen');
  setNote(wb, 0, 5, 6, 'x', 'Avery Chen');
  setNote(wb, 0, 7, 6, 'x', '');
  const runs = { G2: [['Important:', 1], ['\nbold first line', 0]], G4: [['Avery Chen:\nplain name', 0]], G6: [['Sam Rivera:', 1], ['\nanother name', 0]], G8: [['no author', 0]] };
  let t = wb.pkg.text('xl/comments1.xml');
  for (const [ref, rs] of Object.entries(runs)) {
    const body = rs.map(([s, b]) => `<r><rPr>${b ? '<b/>' : ''}<sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${s}</t></r>`).join('');
    t = t.replace(new RegExp(`(<comment ref="${ref}"[^>]*><text>)[\\s\\S]*?(</text>)`), (m, a, z) => a + body + z);
  }
  wb.pkg.setText('xl/comments1.xml', t);
  wb.addSheet(1, 'Second');
  setNote(wb, 1, 2, 1, 'On the second sheet', 'Sam Rivera');
  return wb;
}

if (import.meta.url === pathToFileURL(process.argv[1] || '').href) {
  fs.writeFileSync(process.argv[2], convertInput().toBytes());
  console.log('saved');
}
