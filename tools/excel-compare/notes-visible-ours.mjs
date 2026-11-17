// Show/Hide Note, Show All Notes and Convert to Comments done here, saved for Excel to read back
// (excel-notes-readback.ps1): shown.xlsx, notes.xlsx with B2's note shown and D4's shown and hidden again;
// showall.xlsx, notes.xlsx with every note shown; converted.xlsx, the workbook notes-convert-ours.mjs writes (notes by
// made-up people on two sheets, a thread) with its notes converted to comments.
// usage: node tools/excel-compare/notes-visible-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { setNoteVisible, setAllNotesVisible } from '../../src/ooxml/noteedit.js';
import { convertNotes } from '../../src/ooxml/noteconvert.js';
import { convertInput } from './notes-convert-ours.mjs';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const open = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(f)));
const a = open('test/fixtures/notes.xlsx');
setNoteVisible(a, 0, 1, 1, true);
setNoteVisible(a, 0, 3, 3, true);
setNoteVisible(a, 0, 3, 3, false);
fs.writeFileSync(`${dir}/shown.xlsx`, a.toBytes());
const b = open('test/fixtures/notes.xlsx');
setAllNotesVisible(b, true);
fs.writeFileSync(`${dir}/showall.xlsx`, b.toBytes());
const c = convertInput();
convertNotes(c);
fs.writeFileSync(`${dir}/converted.xlsx`, c.toBytes());
console.log('saved');
