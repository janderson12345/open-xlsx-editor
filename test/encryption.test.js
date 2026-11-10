import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { readCompoundFile, writeCompoundFile, isCompoundFile } from '../src/ooxml/cfb.js';
import { decryptWorkbook, encryptWorkbook, dataSpaces, PasswordError } from '../src/ooxml/encryption.js';
import { Workbook } from '../src/ooxml/workbook.js';
import { cellAt } from '../src/ooxml/sheet.js';

const basic = () => new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'));
const same = (a, b) => Buffer.from(a).equals(Buffer.from(b));
const reason = (fn) => { try { fn(); return null; } catch (e) { assert.ok(e instanceof PasswordError, e.message); return e.reason; } };

test('compound files: streams small and large, in storages, written and read back the same', () => {
  const big = Uint8Array.from({ length: 9 * 1024 * 1024 }, (_, i) => (i * 7919) & 255); // (past 109 FAT sectors: DIFAT sectors)
  const streams = new Map([
    ['EncryptionInfo', Uint8Array.from({ length: 1289 }, (_, i) => i & 255)],
    ['\u0006DataSpaces/Version', Uint8Array.from([1, 2, 3])],
    ['\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', new Uint8Array(200).fill(9)],
    ['Empty', new Uint8Array(0)],
    ['EncryptedPackage', big],
  ]);
  const file = writeCompoundFile(streams);
  assert.ok(isCompoundFile(file));
  const back = readCompoundFile(file);
  assert.deepEqual([...back.keys()].sort(), [...streams.keys()].sort());
  for (const [k, v] of streams) assert.ok(same(back.get(k), v), k);
  assert.ok(!isCompoundFile(basic()), 'an .xlsx is a zip, not a compound file');
});

test('the data spaces beside an encrypted package are the bytes Excel 365 writes', () => {
  // (Read from a workbook a hidden Excel saved with a password: these streams name the transform, nothing else.)
  const excel = {
    '\u0006DataSpaces/Version': '3c0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e004400610074006100530070006100630065007300010000000100000001000000',
    '\u0006DataSpaces/DataSpaceMap': '08000000010000006800000001000000000000002000000045006e0063007200790070007400650064005000610063006b00610067006500320000005300740072006f006e00670045006e006300720079007000740069006f006e004400610074006100530070006100630065000000',
    '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace': '0800000001000000320000005300740072006f006e00670045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d000000',
    '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary': '58000000010000004c0000007b00460046003900410033004600300033002d0035003600450046002d0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200340036007d004e0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00000001000000010000000100000000000000000000000000000004000000',
  };
  const mine = dataSpaces();
  assert.deepEqual([...mine.keys()], Object.keys(excel));
  for (const [k, hex] of Object.entries(excel)) assert.equal(Buffer.from(mine.get(k)).toString('hex'), hex, k);
});

test('a workbook encrypted with a password opens with it, and is refused without it, with a wrong one, or changed', () => {
  const plain = basic();
  const file = encryptWorkbook(plain, 'Test-pass 1', null, { spins: 1000 });
  assert.ok(isCompoundFile(file));
  assert.ok(same(decryptWorkbook(file, 'Test-pass 1').bytes, plain), 'the same bytes back');
  assert.equal(reason(() => decryptWorkbook(file)), 'needed');
  assert.equal(reason(() => decryptWorkbook(file, 'test-pass 1')), 'wrong', 'passwords match case');
  const streams = readCompoundFile(file);
  const pkg = streams.get('EncryptedPackage');
  pkg[pkg.length - 20] ^= 1;
  assert.equal(reason(() => decryptWorkbook(writeCompoundFile(streams), 'Test-pass 1')), 'damaged', 'the integrity check');
  const info = streams.get('EncryptionInfo');
  info[0] = 3; // (version 3.x: the older "standard" encryption)
  assert.equal(reason(() => decryptWorkbook(writeCompoundFile(streams), 'Test-pass 1')), 'unsupported');
});

test('an encrypted workbook opens, is edited and saves encrypted with its password; setPassword changes it', () => {
  const file = encryptWorkbook(basic(), 'Test-pass 1', null, { spins: 1000 });
  assert.throws(() => Workbook.fromBytes(file), PasswordError);
  const wb = Workbook.fromBytes(file, { password: 'Test-pass 1' });
  assert.ok(wb.encrypted);
  wb.setCell(0, 'H20', 'edited');
  const saved = wb.toBytes();
  assert.ok(isCompoundFile(saved), 'saved encrypted again');
  const again = Workbook.fromBytes(saved, { password: 'Test-pass 1' });
  assert.equal(cellAt(again.sheet(0), 19, 7).v, 'edited');
  again.setPassword('New pass 2');
  assert.equal(reason(() => Workbook.fromBytes(again.toBytes(), { password: 'Test-pass 1' })), 'wrong');
  again.setPassword(null);
  const open = again.toBytes();
  assert.ok(!isCompoundFile(open) && Workbook.fromBytes(open).sheets.length === 3, 'no password: an .xlsx again');
});

test('a password with characters past ASCII works both ways (UTF-16, as the format hashes it)', () => {
  const file = encryptWorkbook(basic(), 'Pässwört ✓', null, { spins: 500 });
  assert.equal(decryptWorkbook(file, 'Pässwört ✓').bytes.length, basic().length);
  assert.equal(reason(() => decryptWorkbook(file, 'Passwort ✓')), 'wrong');
});
