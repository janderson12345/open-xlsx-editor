import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { protectSheet, unprotectSheet, sheetProtection, hashPassword, legacyHash, isLocked, protectionRefusal, ALLOWS, PROTECTED_MESSAGE } from '../src/ooxml/protection.js';

// (Excel 365 over COM, Worksheet.Protect, and the files it saved.)
const book = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'P');
  wb.setCell(si, 'A1', 1);
  return { wb, si };
};
const xml = (wb, si) => serializeXml(wb._part(si).doc);
const tag = (wb, si) => xml(wb, si).match(/<sheetProtection[^>]*\/>/)?.[0] ?? null;

test('Protect Sheet as Excel writes it: the defaults, each option allowed as "0", the selection the other way round', () => {
  const { wb, si } = book();
  protectSheet(wb, si);
  assert.equal(tag(wb, si), '<sheetProtection sheet="1" objects="1" scenarios="1"/>');
  assert.match(xml(wb, si), /<\/sheetData><sheetProtection /);
  protectSheet(wb, si, { allow: Object.fromEntries(ALLOWS.map((k) => [k, true])) });
  assert.equal(tag(wb, si), '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0" pivotTables="0"/>');
  protectSheet(wb, si, { selectLocked: false, selectUnlocked: false });
  assert.equal(tag(wb, si), '<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1" selectUnlockedCells="1"/>');
  protectSheet(wb, si, { objects: false, scenarios: false });
  assert.equal(tag(wb, si), '<sheetProtection sheet="1"/>');
  assert.deepEqual(sheetProtection(Workbook.fromBytes(wb.toBytes()), si).allow.sort, false);
});

test("a password as Excel 2013 and later hash it (Excel's own salt and hash for Test-pass 1); unprotecting needs it", () => {
  const salt = Uint8Array.from(atob('elg2Pkx22aHeJJEnE+FfVA=='), (c) => c.charCodeAt(0));
  assert.equal(hashPassword('Test-pass 1', salt), 'Q9iq78G1lcl9fqiuh7fbxUaTXPIc6nVZX0tcNKns35DJrLmJib7hYXbh6UPaaEMQq7U4RMCojlzcIHfhS0UJQQ==');
  const { wb, si } = book();
  protectSheet(wb, si, { password: 'Test-pass 1', salt });
  assert.equal(tag(wb, si), '<sheetProtection algorithmName="SHA-512" hashValue="Q9iq78G1lcl9fqiuh7fbxUaTXPIc6nVZX0tcNKns35DJrLmJib7hYXbh6UPaaEMQq7U4RMCojlzcIHfhS0UJQQ==" saltValue="elg2Pkx22aHeJJEnE+FfVA==" spinCount="100000" sheet="1" objects="1" scenarios="1"/>');
  assert.equal(sheetProtection(wb, si).password, true);
  assert.equal(unprotectSheet(wb, si, 'test-pass 1'), false);
  assert.ok(sheetProtection(wb, si));
  assert.equal(unprotectSheet(wb, si, 'Test-pass 1'), true);
  assert.equal(tag(wb, si), null);
  // (The old 16-bit hash, of files from before Excel 2013: "test" is CBEB.)
  assert.equal(legacyHash('test'), 'CBEB');
});

test('a protected sheet refuses changes to locked cells in Excel\'s words; unlocked cells and allowed formatting go through', () => {
  const { wb, si } = book();
  const open = wb.styleWith(0, { protection: { locked: false } });
  wb.setStyles(si, [{ r: 1, c: 1, s: open }]);
  assert.equal(protectionRefusal(wb, si, 'contents', [{ r: 0, c: 0 }]), null);
  protectSheet(wb, si, { allow: { formatColumns: true } });
  assert.equal(isLocked(wb, si, 0, 0), true);
  assert.equal(isLocked(wb, si, 1, 1), false);
  assert.equal(protectionRefusal(wb, si, 'contents', [{ r: 0, c: 0 }]), PROTECTED_MESSAGE);
  assert.equal(protectionRefusal(wb, si, 'contents', [{ r: 1, c: 1 }]), null);
  assert.equal(protectionRefusal(wb, si, 'formatCells'), PROTECTED_MESSAGE);
  assert.equal(protectionRefusal(wb, si, 'formatColumns'), null);
});
