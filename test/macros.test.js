import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { cellAt } from '../src/ooxml/sheet.js';
import { cellText } from '../src/ooxml/numfmt.js';

// A macro workbook (.xlsm): what makes it one is the workbook part's macro-enabled content type, a VBA project
// (xl/vbaProject.bin, related from the workbook) and often a custom ribbon (customUI/customUI14.xml, related from the
// package). The editor models none of these; an edit, a new sheet and a save must leave them as they were, byte for
// byte, so Excel still opens the file as a macro workbook with its macros and ribbon.
const MACRO_MAIN = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const VBA = Uint8Array.from({ length: 512 }, (_, i) => (i * 37 + 11) & 255);
const RIBBON = '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="t1" label="Tools"><group id="g1" label="Run"><button id="b1" label="Go" onAction="Module1.Go"/></group></tab></tabs></ribbon></customUI>';

function macroBook() {
  const parts = unzipSync(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const edit = (path, from, to) => {
    const s = strFromU8(parts[path]);
    assert.ok(s.includes(from), `${path} has ${from}`);
    parts[path] = strToU8(s.replace(from, to));
  };
  edit('[Content_Types].xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', MACRO_MAIN);
  edit('[Content_Types].xml', '</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/></Types>');
  edit('xl/_rels/workbook.xml.rels', '</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
  edit('_rels/.rels', '</Relationships>', '<Relationship Id="rIdUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>');
  parts['xl/vbaProject.bin'] = VBA;
  parts['customUI/customUI14.xml'] = strToU8(RIBBON);
  return { parts, bytes: zipSync(parts) };
}

test('a macro workbook keeps its VBA project, custom ribbon and macro content type through an edit, a new sheet and a save', () => {
  const { parts: src, bytes } = macroBook();
  const wb = Workbook.fromBytes(bytes);
  wb.apply(wb.activeSheet, [{ r: 0, c: 0, value: 'edited' }]);
  wb.addSheet(wb.sheets.length, 'Added');
  const out = unzipSync(wb.toBytes());
  const same = (a, b) => Buffer.from(a).equals(Buffer.from(b));
  assert.ok(same(out['xl/vbaProject.bin'], VBA), 'the VBA project, byte for byte');
  assert.ok(same(out['customUI/customUI14.xml'], src['customUI/customUI14.xml']), 'the ribbon part, byte for byte');
  assert.ok(same(out['_rels/.rels'], src['_rels/.rels']), 'the package relationships, with the ribbon one');
  const types = strFromU8(out['[Content_Types].xml']);
  assert.ok(types.includes(`ContentType="${MACRO_MAIN}"`), 'the workbook part stays macro-enabled');
  assert.ok(!types.includes('spreadsheetml.sheet.main+xml'), 'and is not turned back into an .xlsx');
  assert.ok(types.includes('Extension="bin" ContentType="application/vnd.ms-office.vbaProject"'), 'the VBA project content type');
  assert.ok(types.includes('PartName="/customUI/customUI14.xml"'), 'the ribbon content type');
  const rels = strFromU8(out['xl/_rels/workbook.xml.rels']);
  assert.ok(/Type="http:\/\/schemas\.microsoft\.com\/office\/2006\/relationships\/vbaProject" Target="vbaProject\.bin"/.test(rels), 'the workbook still relates its VBA project');
  // (Opened again: the edit and the new sheet are there.)
  const again = Workbook.fromBytes(zipSync(out));
  assert.deepEqual(again.sheets.map((s) => s.name).slice(-1), ['Added']);
  assert.equal(cellText(cellAt(again.sheet(again.activeSheet), 0, 0)), 'edited');
});
