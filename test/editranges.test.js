import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { protectSheet, protectionRefusal, PROTECTED_MESSAGE } from '../src/ooxml/protection.js';
import { editRanges, addEditRange, modifyEditRange, deleteEditRange, unlockEditRange, lockedRangeAt } from '../src/ooxml/editranges.js';
import { refersOf, sqrefOfRefers } from '../src/editor/rangesbox.js';

// (Excel 365 over COM: Protection.AllowEditRanges.Add, the sheet protected, cells changed, and the file it saved.)
const book = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'R');
  wb.setCell(si, 'A1', 1);
  return { wb, si };
};
const xml = (wb, si) => serializeXml(wb._part(si).doc);
const salt = Uint8Array.from(atob('zi7x8iS5YOE5vBWLzsx14g=='), (c) => c.charCodeAt(0));

test("Allow Edit Ranges as Excel writes them: after sheetProtection, areas apart by spaces, Excel's own hash", () => {
  const { wb, si } = book();
  addEditRange(wb, si, { name: 'Open', sqref: 'B2:C3' });
  addEditRange(wb, si, { name: 'Keyed', sqref: 'E5', password: 'Test-pass 1', salt });
  addEditRange(wb, si, { name: 'Two', sqref: 'G1:G2 H4' });
  protectSheet(wb, si);
  assert.match(xml(wb, si), /<sheetProtection sheet="1" objects="1" scenarios="1"\/><protectedRanges><protectedRange sqref="B2:C3" name="Open"\/><protectedRange algorithmName="SHA-512" hashValue="eI61rordeGuZxhj3Wige3XOdZD9qyfYlSkjlwtk23mNEYMcoK2EAXIptPGJ3Nq\/voVTHGbTdFnltfCZFxO6pXg==" saltValue="zi7x8iS5YOE5vBWLzsx14g==" spinCount="100000" sqref="E5" name="Keyed"\/><protectedRange sqref="G1:G2 H4" name="Two"\/><\/protectedRanges><pageMargins /);
  const back = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(editRanges(back, si).map((x) => [x.name, x.areas.length, x.password]), [['Open', 1, false], ['Keyed', 1, true], ['Two', 2, false]]);
  assert.equal(deleteEditRange(wb, si, 'Two'), true);
  assert.equal(deleteEditRange(wb, si, 'Open'), true);
  assert.equal(deleteEditRange(wb, si, 'Keyed'), true);
  assert.doesNotMatch(xml(wb, si), /protectedRanges/);
});

test('a protected sheet: a range without a password changes though locked; a change reaching past it, or a range with a password, refused until unlocked', () => {
  const { wb, si } = book();
  addEditRange(wb, si, { name: 'Open', sqref: 'B2:C3' });
  addEditRange(wb, si, { name: 'Keyed', sqref: 'E5', password: 'Test-pass 1', salt });
  addEditRange(wb, si, { name: 'Two', sqref: 'G1:G2 H4' });
  protectSheet(wb, si);
  const refused = (cells) => protectionRefusal(wb, si, 'contents', cells);
  assert.equal(refused([{ r: 1, c: 1 }]), null); // B2
  assert.equal(refused([{ r: 3, c: 7 }]), null); // H4
  assert.equal(refused([{ r: 0, c: 0 }]), PROTECTED_MESSAGE); // A1
  assert.equal(refused([{ r: 1, c: 1 }, { r: 1, c: 2 }, { r: 1, c: 3 }]), PROTECTED_MESSAGE); // B2:D2
  assert.equal(refused([{ r: 4, c: 4 }]), PROTECTED_MESSAGE); // E5
  assert.equal(lockedRangeAt(wb, si, 4, 4), 'Keyed');
  assert.equal(unlockEditRange(wb, si, 'Keyed', 'test-pass 1'), false);
  assert.equal(unlockEditRange(wb, si, 'Keyed', 'Test-pass 1'), true);
  assert.equal(refused([{ r: 4, c: 4 }]), null);
  assert.equal(lockedRangeAt(wb, si, 4, 4), null);
  // (Opened again, locked again: Excel doesn't keep the unlocking.)
  const back = Workbook.fromBytes(wb.toBytes());
  assert.equal(protectionRefusal(back, si, 'contents', [{ r: 4, c: 4 }]), PROTECTED_MESSAGE);
  assert.equal(protectionRefusal(back, si, 'formatCells'), PROTECTED_MESSAGE);
});

test('Modify Range: a new title and cells in its place; a blank password keeps the one it had', () => {
  const { wb, si } = book();
  addEditRange(wb, si, { name: 'A', sqref: 'B2' });
  addEditRange(wb, si, { name: 'Keyed', sqref: 'E5', password: 'Test-pass 1', salt });
  addEditRange(wb, si, { name: 'Z', sqref: 'H4' });
  assert.equal(modifyEditRange(wb, si, 'Keyed', { name: 'Kept', sqref: 'E5:F6', password: null }), true);
  assert.match(xml(wb, si), /<protectedRange sqref="B2" name="A"\/><protectedRange algorithmName="SHA-512" hashValue="eI61ror[^"]*" saltValue="zi7x8iS5YOE5vBWLzsx14g==" spinCount="100000" sqref="E5:F6" name="Kept"\/><protectedRange sqref="H4" name="Z"\/>/);
  assert.equal(unlockEditRange(wb, si, 'Kept', 'Test-pass 1'), true);
  modifyEditRange(wb, si, 'Kept', { name: 'Kept', sqref: 'E5', password: '' });
  assert.deepEqual(editRanges(wb, si).map((x) => [x.name, x.password]), [['A', false], ['Kept', false], ['Z', false]]);
});

test("the box's Refers to cells: =$B$2:$C$3,$H$4 for B2:C3 H4, and back; anything else refused", () => {
  assert.equal(refersOf([{ r0: 1, c0: 1, r1: 2, c1: 2 }, { r0: 3, c0: 7, r1: 3, c1: 7 }]), '=$B$2:$C$3,$H$4');
  assert.equal(sqrefOfRefers('=$B$2:$C$3,$H$4'), 'B2:C3 H4');
  assert.equal(sqrefOfRefers('b2'), 'B2');
  assert.equal(sqrefOfRefers('=Sheet1!B2'), null);
  assert.equal(sqrefOfRefers(''), null);
});
