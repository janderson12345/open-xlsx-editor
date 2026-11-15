import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
import { setNote } from '../src/ooxml/noteedit.js';
import { addThread } from '../src/ooxml/threadedit.js';
import { readNotes } from '../src/ooxml/notes.js';
import { clearChanges } from '../src/ooxml/clearops.js';
import { hyperlinksIn, partsClear } from '../src/ooxml/clearparts.js';

// (Measured over COM in a hidden Excel, the saved files read back: A1 a hyperlink to an address, B1 a note, C1 a
// threaded comment, D1:E2 one hyperlink, A1 bold.)
const LINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
function book() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Clear');
  ['site', 'noted', 'threaded', 'wide'].forEach((v, c) => wb.setCell(si, { r: 0, c }, v));
  wb.setCell(si, 'F1', 'bold');
  const bold = wb.styleWith(0, { font: { bold: true } });
  wb.setStyles(si, [{ r: 0, c: 0, s: bold }, { r: 0, c: 3, s: bold }, { r: 1, c: 4, s: bold }, { r: 0, c: 5, s: bold }]);
  const id = wb.pkg.addRel(LINK, 'https://example.com/', true, wb.sheets[si].path);
  const { doc } = wb._part(si);
  const ns = doc.documentElement.namespaceURI;
  const list = doc.createElementNS(ns, 'hyperlinks');
  const a1 = doc.createElementNS(ns, 'hyperlink');
  a1.setAttribute('ref', 'A1');
  a1.setAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r:id', id);
  const wide = doc.createElementNS(ns, 'hyperlink');
  wide.setAttribute('ref', 'D1:E2');
  wide.setAttribute('location', 'Clear!A10');
  list.appendChild(a1);
  list.appendChild(wide);
  doc.documentElement.insertBefore(list, doc.getElementsByTagNameNS(ns, 'pageMargins')[0] || null);
  wb.dirty.add(wb.sheets[si].path);
  setNote(wb, si, 0, 1, 'a note');
  addThread(wb, si, 0, 2, 'a comment');
  return { wb, si };
}
const area = (text) => { const [a, b = a] = text.split(':').map(parseRef); return { r0: a.r, c0: a.c, r1: b.r, c1: b.c }; };
const links = (wb, si) => hyperlinksIn(wb, si, area('A1:Z99')).map((h) => h.el.getAttribute('ref'));
const notes = (wb, si) => [...readNotes(wb, si).keys()];
const style = (wb, si, ref) => { const { r, c } = parseRef(ref); return wb.sheet(si).cells.get(r)?.get(c)?.s || 0; };
const run = (wb, si, what, sel, withCells = false) => {
  const step = withCells ? clearChanges(wb, si, area(sel), what) : null;
  partsClear(wb, si, [area(sel)], what, step)(wb, si);
};

test('Clear Hyperlinks: a hyperlink meeting the cells goes whole, with its relationship; the formats stay (as Excel)', () => {
  const { wb, si } = book();
  assert.deepEqual(links(wb, si), ['A1', 'D1:E2']);
  run(wb, si, 'hyperlinks', 'D1');
  assert.deepEqual(links(wb, si), ['A1']);
  assert.notEqual(style(wb, si, 'E2'), 0);
  run(wb, si, 'hyperlinks', 'A1:E2');
  assert.deepEqual(links(wb, si), []);
  assert.ok(!/relationships\/hyperlink"/.test(wb.pkg.text(wb.sheets[si].path.replace(/worksheets\//, 'worksheets/_rels/') + '.rels') || ''));
  assert.ok(!wb.pkg.text(wb.sheets[si].path).includes('<hyperlinks'));
});

test('Remove Hyperlinks: the formats of the hyperlinks\' cells go too, theirs only (F1 stays bold, as Excel)', () => {
  const { wb, si } = book();
  run(wb, si, 'removeHyperlinks', 'A1:F1');
  assert.deepEqual(links(wb, si), []);
  assert.deepEqual([style(wb, si, 'A1'), style(wb, si, 'D1'), style(wb, si, 'E2')], [0, 0, 0]);
  assert.notEqual(style(wb, si, 'F1'), 0);
});

test('Clear Comments and Notes: the notes and threaded comments of the cells, nothing else (as Excel)', () => {
  let { wb, si } = book();
  assert.deepEqual(notes(wb, si), ['0,2', '0,1']);
  run(wb, si, 'comments', 'B1');
  assert.deepEqual(notes(wb, si), ['0,2']);
  ({ wb, si } = book());
  run(wb, si, 'comments', 'A1:E2');
  assert.deepEqual([notes(wb, si), links(wb, si)], [[], ['A1', 'D1:E2']]);
});

test('Clear All takes the hyperlinks, notes and comments with the cells; Clear Contents the hyperlinks (as Excel)', () => {
  let { wb, si } = book();
  run(wb, si, 'all', 'A1:E2', true);
  assert.deepEqual([links(wb, si), notes(wb, si), wb.sheet(si).cells.get(0)?.get(0)?.v ?? null], [[], [], null]);
  ({ wb, si } = book());
  run(wb, si, 'contents', 'D1', true);
  assert.deepEqual([links(wb, si), notes(wb, si).length, style(wb, si, 'D1') !== 0], [['A1'], 2, true]);
  // (Clear Formats leaves them all.)
  assert.equal(partsClear(wb, si, [area('A1:E2')], 'formats'), null);
});
