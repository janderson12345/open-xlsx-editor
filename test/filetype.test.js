import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { FILE_TYPES } from '../src/ooxml/filetype.js';

// Excel's SaveAs of one workbook as .xlsx, .xltx, .xltm and .xlsm (a hidden Excel; docs/FIDELITY.md): the files
// differed only in the workbook part's content type in [Content_Types].xml, each the one in FILE_TYPES; the editor's
// own conversion of Excel's .xlsx gave Excel's [Content_Types].xml in all three, byte for byte (checked against
// Excel's files at the time; they hold the user's name, so the fixtures here are made in code).
const basic = () => new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'));
const same = (a, b) => Buffer.from(a).equals(Buffer.from(b));

test('a workbook saved as a template and back: only the workbook part\'s content type changes', () => {
  const src = unzipSync(basic());
  const wb = Workbook.fromBytes(basic());
  assert.equal(wb.fileType, 'xlsx');
  wb.setFileType('xltx');
  const out = unzipSync(wb.toBytes());
  assert.deepEqual(Object.keys(out).sort(), Object.keys(src).sort());
  for (const p of Object.keys(src)) if (p !== '[Content_Types].xml') assert.ok(same(out[p], src[p]), `${p} byte for byte`);
  assert.equal(strFromU8(out['[Content_Types].xml']), strFromU8(src['[Content_Types].xml']).replace(FILE_TYPES.xlsx, FILE_TYPES.xltx));
  // (A template opens as itself, and saved as a workbook is the workbook again, as Excel's new workbook from it is.)
  const t = Workbook.fromBytes(zipSync(out));
  assert.equal(t.fileType, 'xltx');
  t.setFileType('xlsx');
  assert.equal(strFromU8(unzipSync(t.toBytes())['[Content_Types].xml']), strFromU8(src['[Content_Types].xml']));
});

test('a workbook with macros only becomes a macro kind; one without, any kind', () => {
  const parts = unzipSync(basic());
  parts['[Content_Types].xml'] = strToU8(strFromU8(parts['[Content_Types].xml']).replace(FILE_TYPES.xlsx, FILE_TYPES.xlsm).replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'));
  parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
  parts['xl/vbaProject.bin'] = Uint8Array.from({ length: 64 }, (_, i) => i);
  const wb = Workbook.fromBytes(zipSync(parts));
  assert.equal(wb.fileType, 'xlsm');
  assert.throws(() => wb.setFileType('xlsx'), /macros/);
  assert.throws(() => wb.setFileType('xls'), /Not a workbook type/);
  wb.setFileType('xltm');
  const out = unzipSync(wb.toBytes());
  assert.equal(Workbook.fromBytes(zipSync(out)).fileType, 'xltm');
  assert.ok(same(out['xl/vbaProject.bin'], parts['xl/vbaProject.bin']));
  // (Without macros, a macro kind and back.)
  const plain = Workbook.fromBytes(basic());
  plain.setFileType('xlsm');
  plain.setFileType('xlsx');
  assert.equal(plain.fileType, 'xlsx');
});
