import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import { OpcPackage, resolvePath } from '../src/ooxml/package.js';
import { REL_TYPES } from '../src/ooxml/xml.js';

const tiny = () => zipSync({
  '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>'),
  '_rels/.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_TYPES.officeDocument}" Target="xl/workbook.xml"/></Relationships>`),
  'xl/workbook.xml': strToU8('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'),
  'xl/_rels/workbook.xml.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_TYPES.worksheet}" Target="worksheets/sheet1.xml"/></Relationships>`),
  'xl/worksheets/sheet1.xml': strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'),
});

test('the main part is the workbook, found by its officeDocument relationship; targets resolve to package paths', () => {
  const pkg = OpcPackage.fromBytes(tiny());
  assert.equal(pkg.mainPath, 'xl/workbook.xml');
  assert.equal(pkg.mainDir, 'xl/');
  assert.equal(pkg.relById('rId1').target, 'xl/worksheets/sheet1.xml');
  assert.equal(pkg.partByType(REL_TYPES.worksheet), 'xl/worksheets/sheet1.xml');
  assert.equal(resolvePath('xl/worksheets/', '../drawings/drawing1.xml'), 'xl/drawings/drawing1.xml');
  assert.equal(resolvePath('xl/', '/xl/styles.xml'), 'xl/styles.xml');
});

test('saving without changes gives every part back byte for byte, [Content_Types].xml first', () => {
  const src = tiny();
  const before = unzipSync(src);
  const after = unzipSync(OpcPackage.fromBytes(src).toBytes());
  assert.equal(Object.keys(after)[0], '[Content_Types].xml');
  assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort());
  for (const k of Object.keys(before)) assert.ok(Buffer.from(after[k]).equals(Buffer.from(before[k])), k);
});

test('new relationships and content-type overrides; a prefix of its own keeps co-editors\' ids apart', () => {
  const pkg = OpcPackage.fromBytes(tiny());
  const id = pkg.addRel(REL_TYPES.styles, 'styles.xml');
  assert.equal(id, 'rId2');
  assert.equal(pkg.relById('rId2').target, 'xl/styles.xml');
  pkg.addOverride('xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
  pkg.addOverride('xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
  assert.equal(strFromU8(pkg.files['[Content_Types].xml']).split('PartName="/xl/styles.xml"').length, 2);
  pkg.idPrefix = 'rIdc11n';
  assert.match(pkg.addRel(REL_TYPES.theme, 'theme/theme1.xml'), /^rIdc11n\d+$/);
});

test('not a package: a clear error', () => {
  assert.throws(() => OpcPackage.fromBytes(new Uint8Array([1, 2, 3])), /could not unzip/);
  assert.throws(() => OpcPackage.fromBytes(zipSync({ 'a.txt': strToU8('x') })), /Content_Types/);
});
