// Notes written, edited and deleted by the editor (noteedit.js), saved for Excel to open (excel-notes-check.ps1):
// new.xlsx, two new notes on widths.xlsx (which has none); edit.xlsx, notes.xlsx with B2's note's text replaced, one
// added at A6, D4's deleted and C3's threaded comment deleted. expected.json: what each should hold.
// usage: node tools/excel-compare/notes-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { setNote, deleteNote } from '../../src/ooxml/noteedit.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const open = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`)));
const a = open('widths.xlsx');
setNote(a, 0, 4, 2, 'Check this width', 'Avery Chen');
setNote(a, 0, 0, 0, 'First line\nSecond line', 'Sam Rivera');
fs.writeFileSync(`${dir}/new.xlsx`, a.toBytes());
const b = open('notes.xlsx');
setNote(b, 0, 1, 1, 'Totals checked', 'Someone Else');
setNote(b, 0, 5, 0, 'A new one', 'Sam Rivera');
deleteNote(b, 0, 3, 3);
deleteNote(b, 0, 2, 2);
fs.writeFileSync(`${dir}/edit.xlsx`, b.toBytes());
fs.writeFileSync(`${dir}/expected.json`, JSON.stringify({
  'new.xlsx': [['A1', 'Sam Rivera', 'Sam Rivera:\nFirst line\nSecond line'], ['C5', 'Avery Chen', 'Avery Chen:\nCheck this width']],
  'edit.xlsx': [['A6', 'Sam Rivera', 'Sam Rivera:\nA new one'], ['B2', 'Avery Chen', 'Avery Chen:\nTotals checked']],
}));
console.log('saved');
