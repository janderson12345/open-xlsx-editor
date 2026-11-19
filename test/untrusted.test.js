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
import { SheetCells } from '../src/ooxml/sheetcells.js';
import { autoRowHeights } from '../src/editor/rowheights.js';

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

test('a workbook whose parts say they unpack past 1 GB is refused before anything is unpacked', () => {
  const zip = zipSync({ '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'), 'xl/a.xml': strToU8('<a/>') });
  const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  // (Each central directory entry's uncompressed size, at 24 past its signature, made 600 MB: 1.2 GB in all.)
  for (let o = 0; o + 4 <= zip.length; o++) if (dv.getUint32(o, true) === 0x02014b50) dv.setUint32(o + 24, 600 * 1024 * 1024, true);
  const started = Date.now();
  assert.throws(() => OpcPackage.fromBytes(zip), /damaged, or it unpacks to more than 1 GB/);
  assert.ok(Date.now() - started < 500);
  // (A real workbook opens as before.)
  assert.ok(OpcPackage.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'))).has('[Content_Types].xml'));
});

test('Convert to Comments writes a person id from the file escaped, as one attribute', async () => {
  const { readNotes } = await import('../src/ooxml/notes.js');
  const { convertNotes } = await import('../src/ooxml/noteconvert.js');
  const { personsPart, threadsPart } = await import('../src/ooxml/threadedit.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/notes.xlsx')));
  const author = [...readNotes(wb, 0).values()].find((n) => n.kind === 'note' && n.author)?.author;
  assert.ok(author, 'the fixture has a note with an author');
  // (The workbook's own person for that author, its id crafted to end the attribute and add one.)
  const path = personsPart(wb);
  const name = author.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  wb.pkg.setText(path, wb.pkg.text(path).replace(/<\/personList>/, `<person displayName="${name}" id="x&quot; evil=&quot;1" userId="${name}" providerId="None"/></personList>`));
  assert.ok(convertNotes(wb) > 0);
  const xml = wb.pkg.text(threadsPart(wb, 0));
  const doc = parseXml(xml);
  const threads = [...doc.getElementsByTagName('threadedComment')];
  assert.ok(threads.length > 0);
  for (const t of threads) assert.equal(t.hasAttribute('evil'), false);
  assert.ok(threads.some((t) => t.getAttribute('personId') === 'x" evil="1'));
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

test("a font named constructor or __proto__ is a font Windows hasn't, not one of Object's own properties", async () => {
  // (The font tables were looked up by the name in the file as it stood: fontStack threw for these names, and so did
  // opening a workbook that used one.)
  const { windowsHasFont, fontLinePx, fontRowPx } = await import('../src/editor/fonts.js');
  const { excelWidths } = await import('../src/ooxml/excelwidths.js');
  const { excelFontMetric, excelFontLine } = await import('../src/ooxml/excelfonts.js');
  const unknown = 'Nofont Xyz';
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Names');
  wb.setCell(si, 'A1', 'a wrapped line of text, long enough to take a few lines in a narrow column');
  const sheet = wb.sheet(si);
  const rowOf = (family) => autoRowHeights(sheet, () => ({ font: { family, sizePx: 44 / 3, bold: false, italic: true, css: '' }, align: { wrap: true } }), { size: () => 64 }).get(0);
  for (const name of ['constructor', '__proto__', 'toString', 'hasOwnProperty', 'valueOf']) {
    assert.equal(fontStack(name), `${name}, sans-serif`);
    assert.equal(windowsHasFont(name), false);
    assert.deepEqual([fontRowPx(name, 15), fontLinePx(name, 15, true, false, true)], [fontRowPx(unknown, 15), fontLinePx(unknown, 15, true, false, true)]);
    assert.deepEqual([excelWidths(name, 15, true, true), excelFontMetric(name, 11, 'row'), excelFontLine(name, 11, false, false)], [null, null, null]);
    assert.equal(rowOf(name), rowOf(unknown));
  }
});

test('wrapped cells of crafted long words and long runs of spaces are sized at once', async () => {
  // (A word too wide for its column was cut a character at a time, the line measured again for each, and the spaces
  // after it walked over again at each cut: a million characters, which a file zips to a few KB, took minutes; and a
  // rich text of 300,000 runs overflowed the stack working out their italic allowance.)
  const { excelLineCount, wrapLines } = await import('../src/editor/rowheights.js');
  const font = { family: 'Calibri', sizePx: 44 / 3, bold: false, italic: false, roman: false, css: '11pt Calibri' };
  const started = Date.now();
  // (Past the 32,767 characters a cell holds in Excel nothing is measured.)
  assert.equal(excelLineCount(`${'W'.repeat(500000)}${' '.repeat(500000)}`, font, 64), excelLineCount('W'.repeat(32767), font, 64));
  assert.equal(excelLineCount(`${'W'.repeat(16000)}${' '.repeat(16000)}x`, font, 64), excelLineCount(`${'W'.repeat(16000)} x`, font, 64));
  assert.equal(excelLineCount('i'.repeat(1000000), font, 1790), 77);
  assert.equal(wrapLines(`${'W'.repeat(500000)} x`, '11pt Calibri', 64).length, 1);
  const runs = Array.from({ length: 300000 }, (_, k) => ({ text: 'ab', font: { name: 'Calibri', size: 11, b: false, i: k % 2 === 0 } }));
  assert.ok(excelLineCount('ab'.repeat(300000).slice(0, 32767), font, 64, 0, runs) > 1);
  assert.ok(excelLineCount('ab'.repeat(16000), font, 64, 0, runs.slice(0, 16000)) > 1);
  assert.ok(Date.now() - started < 3000, `took ${Date.now() - started} ms`);
});

test('rows of formatted empty cells alternating styles across every column are sized at once', () => {
  // (200 rows, each 16,384 empty cells of two styles in turn, 8,192 runs, every row a little different so none are
  // one band: counting each band's empty columns run by run took 200 x 16,384 x 8,192 steps.)
  const cells = new SheetCells();
  const runs = [];
  for (let c = 0; c < 16384; c++) runs.push(c, c, c % 2);
  for (let r = 0; r < 200; r++) { runs[2] = r % 2 ? 1 : 0; runs[5] = r % 3 ? 1 : 0; runs[8] = r; cells.addBlankRow(r, runs, runs.length); }
  const sheet = { cells, rows: new Map(), cols: [], merges: [], format: { defaultRowHeight: 15 } };
  const deco = () => ({ font: { family: 'Calibri', sizePx: 44 / 3, bold: false, css: '11pt Calibri' }, align: {} });
  const started = Date.now();
  const out = autoRowHeights(sheet, deco, { size: () => 64 });
  assert.ok(Date.now() - started < 3000, `took ${Date.now() - started} ms`);
  assert.equal(out.size, 200);
});
