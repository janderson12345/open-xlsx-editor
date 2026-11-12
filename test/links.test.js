import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { addHyperlink, readHyperlinks, hyperlinkAt, screenTip, followOf } from '../src/ooxml/links.js';
import { addressOf, mailtoOf, locationOf } from '../src/editor/linkbox.js';

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
