import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OpcPackage } from '../src/ooxml/package.js';
import { validationXml } from '../src/ooxml/dvedit.js';
import { parseDecls } from '../src/ooxml/htmltable.js';
import { Pattern, STEPS, TOTAL_STEPS } from '../src/calc/pcre-match.js';
import { E } from '../src/calc/values.js';
import fs from 'node:fs';
import { zipSync, strToU8 } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { readObjects, GROUP_DEPTH } from '../src/ooxml/drawobjects.js';
import { customPaths } from '../src/ooxml/presets.js';
import { readMashup } from '../src/ooxml/queries.js';
import { fontStack } from '../src/editor/fonts.js';

// Text a workbook, the clipboard or a pattern brings, handled so it can't change what the code does: the security
// review of the November merges.

test('removeRel takes a relationship id as text, not as a pattern', () => {
  const rels = '<Relationships><Relationship Id="rId1" Target="a.xml"/><Relationship Id="(a+)+" Target="b.xml"/></Relationships>';
  const files = { 'xl/drawings/_rels/drawing1.xml.rels': rels };
  const pkg = { text: (p) => files[p] ?? null, setText: (p, s) => { files[p] = s; } };
  OpcPackage.prototype.removeRel.call(pkg, 'xl/drawings/drawing1.xml', '(a+)+');
  assert.equal(files['xl/drawings/_rels/drawing1.xml.rels'], '<Relationships><Relationship Id="rId1" Target="a.xml"/></Relationships>');
  OpcPackage.prototype.removeRel.call(pkg, 'xl/drawings/drawing1.xml', 'rId.');
  assert.match(files['xl/drawings/_rels/drawing1.xml.rels'], /Id="rId1"/);
});

test('a validation rule read from a file writes only the error styles and types Excel has', () => {
  const ranges = [{ r0: 0, c0: 0, r1: 0, c1: 0 }];
  const xml = validationXml({ type: 'list', f1: 'a,b', error: { style: 'warning" evil="1' } }, ranges);
  assert.doesNotMatch(xml, /evil/);
  assert.doesNotMatch(xml, /errorStyle/);
  assert.match(validationXml({ type: 'whole', op: 'greater', f1: '1', error: { style: 'information' } }, ranges), /errorStyle="information"/);
  const odd = validationXml({ type: 'constructor', op: 'toString', f1: '1' }, ranges);
  assert.doesNotMatch(odd, /function|native/);
});

test('clipboard CSS with a long run of spaces is read at once, !important still cut off', () => {
  const started = Date.now();
  const decls = parseDecls(`color: red${' '.repeat(200000)}x; font-weight: 700 !important`);
  assert.ok(Date.now() - started < 500);
  assert.equal(decls['font-weight'], '700');
  assert.ok(decls.color.startsWith('red'));
});

test('all the searches of one REGEXREPLACE share a limit', () => {
  const p = new Pattern('a', false);
  let calls = 0;
  p.exec = (s, from) => { calls++; p.S.steps = STEPS; return { index: from, start: from, end: from + 1, caps: [] }; };
  assert.throws(() => p.all('a'.repeat(1000)), (e) => e === E.value);
  assert.equal(calls, TOTAL_STEPS / STEPS + 1);
});

// (The security review of 0.8.1: the drawings, the query package and the font names a workbook brings.)

test('a font name is quoted so that nothing in it can end the CSS string', () => {
  const css = fontStack('x\n;}body{background:url(//e.example)}\\\'"');
  const name = css.slice(0, css.indexOf(','));
  assert.match(name, /^'[^'"\\\n\r]*'$/);
  assert.doesNotMatch(css, /[\n\r\\"]/);
  assert.equal(fontStack('Arial').split(',')[0], 'Arial');
  assert.equal(fontStack('Times New Roman').split(',')[0], "'Times New Roman'");
});

test('a custom shape\'s arc with a huge sweep draws at most a turn, at once', () => {
  const started = Date.now();
  const paths = customPaths([{ w: 10, h: 10, cmds: [{ op: 'moveTo', pts: [[10, 5]] }, { op: 'arcTo', pts: [], wR: 5, hR: 5, stAng: 0, swAng: 1e300 }, { op: 'arcTo', pts: [], wR: 5, hR: 5, stAng: 0, swAng: Infinity }] }], 10, 10);
  assert.ok(Date.now() - started < 500);
  assert.ok(paths[0].cmds.length <= 1 + 4 + 4);
});

test('drawing parts nested thousands deep are read as far as Excel would nest them, not until the stack runs out', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/drawings.xlsx')));
  const part = wb.pkg.readRels(wb.sheets[0].path).find((r) => /drawing/.test(r.type)).target;
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const corner = (tag) => `<xdr:${tag}><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
  const group = (inner) => `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="2" name="G"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>${inner}</xdr:grpSp>`;
  let deep = '';
  for (let k = 0; k < 5000; k++) deep = group(deep);
  wb.pkg.setText(part, `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor>${corner('from')}${corner('to')}${deep}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
  const [obj] = readObjects(wb, 0);
  assert.equal(obj.kind, 'group');
  let levels = 0;
  for (let g = obj.item; g; g = g.children?.[0]) levels++;
  assert.equal(levels, GROUP_DEPTH);
  // (An anchor buried under thousands of other elements isn't looked for.)
  let buried = `<xdr:twoCellAnchor>${corner('from')}${corner('to')}<xdr:clientData/></xdr:twoCellAnchor>`;
  for (let k = 0; k < 5000; k++) buried = `<x>${buried}</x>`;
  wb.pkg.setText(part, `<xdr:wsDr xmlns:xdr="${XDR}">${buried}</xdr:wsDr>`);
  assert.deepEqual(readObjects(wb, 0), []);
});

test('a query package that says it unpacks past 64 MB is not unpacked', () => {
  const m = 'section Section1;\nshared Q = 1;';
  const zip = zipSync({ 'Formulas/Section1.m': strToU8(m) });
  // (The central directory's uncompressed size, at 24 past its signature, made 100 MB.)
  const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  let cd = zip.length - 4;
  while (cd >= 0 && dv.getUint32(cd, true) !== 0x02014b50) cd--;
  const mashup = () => {
    const bin = new Uint8Array(4 + 4 + zip.length + 12);
    new DataView(bin.buffer).setUint32(4, zip.length, true);
    bin.set(zip, 8);
    return strToU8(`<?xml version="1.0" encoding="utf-8"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${Buffer.from(bin).toString('base64')}</DataMashup>`);
  };
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  dv.setUint32(cd + 24, 100 * 1024 * 1024, true);
  wb.pkg.files['customXml/item1.xml'] = mashup();
  assert.equal(readMashup(wb), null);
  // (The same package with its true size is read.)
  dv.setUint32(cd + 24, strToU8(m).length, true);
  wb.pkg.files['customXml/item1.xml'] = mashup();
  assert.deepEqual(readMashup(wb)?.queries.map((q) => q.name), ['Q']);
});
