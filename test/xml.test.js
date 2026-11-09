import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, kids, kid, attr, onOff, NS, decodeXstring, outerXml, nsDecls } from '../src/ooxml/xml.js';

const SHEET = `<worksheet xmlns="${NS.x}" xmlns:r="${NS.r}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><x:foo xmlns:x="urn:other"/></worksheet>`;

test('lookups default to SpreadsheetML and skip other namespaces', () => {
  const root = parseXml(SHEET).documentElement;
  assert.equal(kids(root).length, 2);
  assert.equal(kids(root, 'sheetData').length, 1);
  assert.equal(kid(root, 'foo'), null);
  const c = kid(kid(kid(root, 'sheetData'), 'row'), 'c');
  assert.equal(attr(c, 'r'), 'A1');
  assert.equal(attr(c, 't'), 's');
});

test('a prefixed SpreadsheetML part reads the same as an unprefixed one', () => {
  const root = parseXml(`<x:worksheet xmlns:x="${NS.x}"><x:sheetData><x:row r="2"/></x:sheetData></x:worksheet>`).documentElement;
  assert.equal(attr(kid(kid(root, 'sheetData'), 'row'), 'r'), '2');
});

test('on/off properties: <b/> on, <b val="0"/> and <b val="false"/> off', () => {
  const root = parseXml(`<font xmlns="${NS.x}"><b/><i val="0"/><u val="false"/><strike val="1"/></font>`).documentElement;
  assert.deepEqual(['b', 'i', 'u', 'strike'].map((n) => onOff(kid(root, n))), [true, false, false, true]);
  assert.equal(onOff(kid(root, 'outline')), null);
});

test('ST_Xstring escapes decode, and _x005F_ keeps a literal escape as text', () => {
  assert.equal(decodeXstring('a_x000D_b'), 'a\rb');
  assert.equal(decodeXstring('_x005F_x000D_'), '_x000D_');
  assert.equal(decodeXstring('plain'), 'plain');
});

test('outerXml drops namespace declarations the root already makes', () => {
  const doc = parseXml(SHEET);
  const sd = kid(doc.documentElement, 'sheetData');
  assert.ok(!outerXml(sd, nsDecls(doc.documentElement)).includes('xmlns='));
});
