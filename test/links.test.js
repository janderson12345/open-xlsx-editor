import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { addHyperlink, readHyperlinks, hyperlinkAt, screenTip, followOf, encodeAddress, decodeAddress } from '../src/ooxml/links.js';
import { addressOf, mailtoOf, locationOf, newDocumentName } from '../src/editor/linkbox.js';

// (Excel 365 over COM: Hyperlinks.Add on A1:A7 of a new sheet, and the file it saved; the xr:uid Excel adds left out.)
const book = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'L');
  return { wb, si };
};
const value = (wb, si, r, c) => wb.sheet(si).cells.get(r)?.get(c)?.v ?? null;

test('hyperlinks as Excel writes them: an address as an external relationship, a place as location, display and tooltip as Excel gives them', () => {
  const { wb, si } = book();
  addHyperlink(wb, si, 0, 0, { address: 'https://example.com/page' });
  addHyperlink(wb, si, 1, 0, { address: 'https://example.com/x', tooltip: 'Tip here', text: 'Shown text' });
  addHyperlink(wb, si, 2, 0, { location: "'Sheet2'!B5", text: 'Go there' });
  addHyperlink(wb, si, 3, 0, { address: 'mailto:someone@example.com?subject=Hi' });
  wb.setCell(si, 'A5', 'kept text');
  addHyperlink(wb, si, 4, 0, { address: 'https://example.com/' });
  wb.setCell(si, 'A6', 42);
  addHyperlink(wb, si, 5, 0, { address: 'https://example.com/n' });
  const xml = serializeXml(wb._part(si).doc);
  const tags = xml.match(/<hyperlinks>.*<\/hyperlinks>/)[0].replace(/ r:id="rId\d+"/g, ' r:id');
  assert.equal(tags, '<hyperlinks><hyperlink ref="A1" r:id/><hyperlink ref="A2" r:id tooltip="Tip here"/><hyperlink ref="A3" location="\'Sheet2\'!B5" display="Go there"/>'
    + '<hyperlink ref="A4" r:id/><hyperlink ref="A5" r:id display="https://example.com/"/><hyperlink ref="A6" r:id display="https://example.com/n"/></hyperlinks>');
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((r) => value(wb, si, r, 0)), ['https://example.com/page', 'Shown text', 'Go there', 'mailto:someone@example.com?subject=Hi', 'kept text', 42]);
  const back = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(readHyperlinks(back, si).map((h) => [h.address, h.location]), [['https://example.com/page', null], ['https://example.com/x', null], [null, "'Sheet2'!B5"],
    ['mailto:someone@example.com?subject=Hi', null], ['https://example.com/', null], ['https://example.com/n', null]]);
  // (The Hyperlink cell style, as Excel adds it: underlined, the theme's colour 10, nothing marked applied.)
  const styles = serializeXml(back.pkg.xml(back.stylesPath));
  assert.match(styles, /<cellStyle name="Hyperlink" xfId="(\d+)" builtinId="8"\/>/);
  assert.match(styles, /<font><u\/><sz val="11"\/><color theme="10"\/>/);
  const s = back.sheet(si).cells.get(0).get(0).s;
  assert.match(styles.match(/<cellXfs[^>]*>(.*)<\/cellXfs>/)[1].match(/<xf [^>]*\/>|<xf [^>]*>.*?<\/xf>/g)[s], /^<xf numFmtId="0" fontId="\d+" fillId="0" borderId="0" xfId="\d+"\/>$/);
  // (A cell's own hyperlink goes when it gets another.)
  addHyperlink(wb, si, 0, 0, { address: 'https://example.com/other' });
  assert.equal(readHyperlinks(wb, si).filter((h) => h.range.r0 === 0).length, 1);
  assert.equal(hyperlinkAt(wb, si, 0, 0).address, 'https://example.com/other');
});

test("following a hyperlink: http, https, mailto and ftp opened; a place in the workbook gone to; anything else refused in Excel's words; Excel's ScreenTip", () => {
  assert.deepEqual(followOf({ address: 'https://example.com/a', location: null }), { url: 'https://example.com/a' });
  assert.deepEqual(followOf({ address: 'https://example.com/a', location: 'top' }), { url: 'https://example.com/a#top' });
  assert.deepEqual(followOf({ address: 'mailto:someone@example.com', location: null }), { url: 'mailto:someone@example.com' });
  // (An address read back unescaped, opened escaped again, as its Target has it.)
  assert.deepEqual(followOf({ address: 'https://example.com/a b?q=100%', location: null }), { url: 'https://example.com/a%20b?q=100%25' });
  assert.deepEqual(followOf({ address: null, location: "'Sheet2'!B5" }), { place: "'Sheet2'!B5" });
  assert.deepEqual(followOf({ address: 'example.com', location: null }), { refused: 'Cannot open the specified file.' });
  assert.deepEqual(followOf({ address: 'javascript:alert(1)', location: null }), { refused: 'Cannot open the specified file.' });
  assert.equal(screenTip({ address: 'https://example.com/page', location: null, tooltip: null }), 'https://example.com/page - Click once to follow. Click and hold to select this cell.');
  assert.equal(screenTip({ address: 'https://example.com/x', location: null, tooltip: 'Tip here' }), 'Tip here - Click once to follow. Click and hold to select this cell.');
});

test("Insert Hyperlink's fields as the address and place: http:// before an address starting www., mailto: with its subject, the sheet quoted as a formula quotes it", () => {
  assert.equal(addressOf('example.com/page'), 'example.com/page');
  assert.equal(addressOf('www.example.com'), 'http://www.example.com');
  assert.equal(addressOf('https://example.com/'), 'https://example.com/');
  assert.equal(addressOf('ftp://files.example.com/a'), 'ftp://files.example.com/a');
  assert.equal(addressOf('C:\\Docs\\a.xlsx'), 'C:\\Docs\\a.xlsx');
  assert.equal(addressOf('report.xlsx'), 'report.xlsx');
  assert.equal(addressOf('notes'), 'notes');
  assert.equal(mailtoOf('someone@example.com', 'Hi there'), 'mailto:someone@example.com?subject=Hi%20there');
  assert.equal(mailtoOf('mailto:someone@example.com', ''), 'mailto:someone@example.com');
  assert.equal(locationOf('Notes & more', 'C30'), "'Notes & more'!C30");
  assert.equal(locationOf('Data', 'A1'), 'Data!A1');
  assert.equal(locationOf(null, 'MyName'), 'MyName');
});

test("HYPERLINK's cells as Excel's: the Hyperlink style when typed at the formula's top level; a click follows its link_location", async () => {
  const { topLevelHyperlink, formulaLink, cellLink } = await import('../src/calc/hyperlinkcell.js');
  const { entryChanges } = await import('../src/ooxml/entry.js');
  assert.equal(topLevelHyperlink('HYPERLINK("https://example.com/a","Shown")'), true);
  assert.equal(topLevelHyperlink('1+HYPERLINK("https://example.com/c",5)'), true);
  assert.equal(topLevelHyperlink('IF(TRUE,HYPERLINK("https://example.com/d","In if"))'), false);
  assert.equal(topLevelHyperlink('SUM(1,2)'), false);
  const { wb, si } = book();
  wb.setCell(si, 'B1', 'page');
  wb.setFormula(si, 'A1', 'HYPERLINK("https://example.com/"&B1,"Shown")');
  wb.setFormula(si, 'A2', 'HYPERLINK("#Sheet2!B5","Go")');
  wb.setFormula(si, 'A3', 'IF(TRUE,HYPERLINK("https://example.com/d","In if"))');
  wb.setFormula(si, 'A4', 'SUM(1,2)');
  assert.deepEqual(formulaLink(wb, si, 0, 0), { address: 'https://example.com/page', location: null, display: null, tooltip: null });
  assert.equal(formulaLink(wb, si, 1, 0).location, 'Sheet2!B5');
  assert.equal(formulaLink(wb, si, 2, 0).address, 'https://example.com/d');
  assert.equal(formulaLink(wb, si, 3, 0), null);
  addHyperlink(wb, si, 5, 0, { address: 'https://example.com/own' });
  assert.equal(cellLink(wb, si, 5, 0).address, 'https://example.com/own');
  const typed = (text) => entryChanges(wb, si, 9, 0, text);
  const style = (ch) => ch.style && wb.styles.cellXfs[ch.style.s].xfId;
  assert.ok(style(typed('=HYPERLINK("https://example.com/a","Shown")')) > 0);
  assert.ok(style(typed('=1+HYPERLINK("https://example.com/c",5)')) > 0);
  assert.equal(typed('=IF(TRUE,HYPERLINK("https://example.com/d","In if"))').style, null);
});

test('a link to a file as Excel writes it: the name relative to the workbook as its Target, escaped as Excel escapes it, TargetMode External, a place in it as location', () => {
  // (Excel 365 over COM: tools/excel-compare/excel-filelinks.ps1, Hyperlinks.Add of file names and the file saved;
  // test/filelinks-excel.json. A full path in the workbook's own folder was saved as the name alone: Browse and
  // Create New Document here give the name.)
  const cases = JSON.parse(fs.readFileSync('test/filelinks-excel.json', 'utf8'));
  const { wb, si } = book();
  cases.forEach((x, r) => addHyperlink(wb, si, r, 0, { address: x.given.replace(/^<folder>./, ''), location: x.subAddress }));
  const rels = wb.pkg.readRels(wb.sheets[si].path);
  const els = [...wb._part(si).doc.getElementsByTagNameNS('*', 'hyperlink')];
  cases.forEach((x, r) => {
    const rel = rels.find((y) => y.id === els[r].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
    assert.equal(rel.rawTarget, x.target, x.ref);
    assert.equal(rel.external ? 'External' : null, x.mode, x.ref);
    assert.equal(els[r].getAttribute('location') || null, x.location, x.ref);
    assert.equal(readHyperlinks(wb, si)[r].address, x.address, x.ref);
  });
  assert.equal(encodeAddress('a b%20c%zz'), 'a%20b%20c%25zz');
  assert.equal(decodeAddress('My%20File%5B1%5D.xlsx'), 'My File[1].xlsx');
});

test("Create New Document's name as the link's address: .xlsx added when it has no extension; a name a file can't have refused", () => {
  assert.equal(newDocumentName('Budget'), 'Budget.xlsx');
  assert.equal(newDocumentName('  Notes.docx '), 'Notes.docx');
  assert.equal(newDocumentName('Plan 2027.xlsm'), 'Plan 2027.xlsm');
  for (const bad of ['', '   ', 'a/b', 'a:b', 'what?', '..']) assert.equal(newDocumentName(bad), null, bad);
});
