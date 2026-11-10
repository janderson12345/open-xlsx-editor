// Workbooks saved with a password (File > Info > Protect Workbook > Encrypt with Password): the .xlsx's bytes
// encrypted inside a compound file (cfb.js), as MS-OFFCRYPTO's agile encryption describes it. The password's key is a
// salted hash spun spinCount times (100,000 in Excel's), which opens the package's own random key; the package is
// AES-CBC in 4096-byte segments, each with its own IV, and an HMAC over it shows it hasn't been changed. Read: Excel
// 2010 and later's agile encryption, SHA-1/256/384/512 and AES-128/192/256. Written: the same as Excel 365 writes
// (AES-256, SHA-512, 100,000 spins), every salt and key new each save.
import { sha1 } from '@noble/hashes/legacy';
import { sha256, sha384, sha512 } from '@noble/hashes/sha2';
import { hmac } from '@noble/hashes/hmac';
import { randomBytes } from '@noble/hashes/utils';
import { cbc } from '@noble/ciphers/aes';
import { readCompoundFile, writeCompoundFile, isCompoundFile } from './cfb.js';
import { parseXml, attr, kid, kids } from './xml.js';

export const ENCRYPTION_MESSAGES = {
  needed: 'This workbook is protected with a password.',
  wrong: "The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.",
  unsupported: "This workbook is encrypted in a way that can't be opened here (only Excel 2010 and later's encryption can).",
  damaged: 'The encrypted workbook is damaged: its contents do not match its integrity check.',
};

/** A refusal to open an encrypted workbook: `reason` 'needed', 'wrong', 'unsupported' or 'damaged'. */
export class PasswordError extends Error {
  constructor(reason) {
    super(ENCRYPTION_MESSAGES[reason]);
    this.name = 'PasswordError';
    this.reason = reason;
  }
}

const HASHES = { SHA1: sha1, SHA256: sha256, SHA384: sha384, SHA512: sha512 };
const BLOCK = {
  verifierInput: [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79],
  verifierHash: [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e],
  key: [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6],
  hmacKey: [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6],
  hmacValue: [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33],
};
const SEGMENT = 4096;
const ENCRYPTION_NS = 'http://schemas.microsoft.com/office/2006/encryption';
const PASSWORD_NS = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';

const concat = (...parts) => { const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
const u32 = (n) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
const utf16 = (s) => { const b = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { b[i * 2] = s.charCodeAt(i) & 255; b[i * 2 + 1] = s.charCodeAt(i) >> 8; } return b; };
const fromBase64 = (s) => Uint8Array.from(atob(s), (ch) => ch.charCodeAt(0));
const toBase64 = (b) => { let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); };
/** `b` cut to n bytes, or padded to them with `pad`. */
const fit = (b, n, pad = 0x36) => (b.length >= n ? b.slice(0, n) : concat(b, new Uint8Array(n - b.length).fill(pad)));
/** `b` padded with zeros to a whole number of cipher blocks. */
const blocks = (b, size) => fit(b, Math.ceil(b.length / size) * size, 0);
const equal = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
const decrypt = (key, iv, data) => cbc(key, iv, { disablePadding: true }).decrypt(data);
const encrypt = (key, iv, data) => cbc(key, iv, { disablePadding: true }).encrypt(data);

/** The key from a password for one purpose (`blockKey`), as MS-OFFCRYPTO 2.3.4.11 derives it. */
function passwordKey(hash, password, salt, spins, blockKey, bytes) {
  let h = hash(concat(salt, utf16(password)));
  const buf = new Uint8Array(4 + h.length);
  const dv = new DataView(buf.buffer);
  for (let i = 0; i < spins; i++) {
    dv.setUint32(0, i, true);
    buf.set(h, 4);
    h = hash(buf);
  }
  return fit(hash(concat(h, Uint8Array.from(blockKey))), bytes);
}

/** The settings and values of an EncryptionInfo stream (agile only). */
function readInfo(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (dv.getUint16(0, true) !== 4 || dv.getUint16(2, true) !== 4) throw new PasswordError('unsupported');
  const root = parseXml(new TextDecoder().decode(bytes.subarray(8))).documentElement;
  const keyData = kid(root, 'keyData', ENCRYPTION_NS);
  const integrity = kid(root, 'dataIntegrity', ENCRYPTION_NS);
  const encryptor = kids(kid(root, 'keyEncryptors', ENCRYPTION_NS), 'keyEncryptor', ENCRYPTION_NS).find((e) => attr(e, 'uri') === PASSWORD_NS);
  const key = encryptor && kid(encryptor, 'encryptedKey', PASSWORD_NS);
  if (!keyData || !key) throw new PasswordError('unsupported');
  const settings = (el) => {
    const hash = HASHES[attr(el, 'hashAlgorithm')];
    if (!hash || attr(el, 'cipherAlgorithm') !== 'AES' || attr(el, 'cipherChaining') !== 'ChainingModeCBC') throw new PasswordError('unsupported');
    return { hash, keyBytes: Number(attr(el, 'keyBits')) / 8, blockSize: Number(attr(el, 'blockSize')), hashSize: Number(attr(el, 'hashSize')), salt: fromBase64(attr(el, 'saltValue')) };
  };
  return {
    keyData: settings(keyData),
    key: { ...settings(key), spins: Number(attr(key, 'spinCount')), verifierInput: fromBase64(attr(key, 'encryptedVerifierHashInput')), verifierHash: fromBase64(attr(key, 'encryptedVerifierHashValue')), keyValue: fromBase64(attr(key, 'encryptedKeyValue')) },
    integrity: integrity && { key: fromBase64(attr(integrity, 'encryptedHmacKey')), value: fromBase64(attr(integrity, 'encryptedHmacValue')) },
  };
}

/** The IV of segment i, or of an HMAC value (`blockKey`): the keyData salt and the index hashed, cut to a block. */
const ivFor = (kd, extra) => fit(kd.hash(concat(kd.salt, extra)), kd.blockSize);

/**
 * The .xlsx inside an encrypted workbook: { bytes, streams } (`streams` the compound file's others, kept to save it
 * again). Throws a PasswordError: 'needed' without a password, 'wrong', 'unsupported' or 'damaged'.
 */
export function decryptWorkbook(file, password) {
  const streams = readCompoundFile(file);
  const infoBytes = streams.get('EncryptionInfo');
  const pkg = streams.get('EncryptedPackage');
  if (!infoBytes || !pkg) throw new PasswordError('unsupported');
  const info = readInfo(infoBytes);
  if (password == null) throw new PasswordError('needed');
  const k = info.key;
  const at = (blockKey) => passwordKey(k.hash, password, k.salt, k.spins, blockKey, k.keyBytes);
  const input = decrypt(at(BLOCK.verifierInput), k.salt, k.verifierInput).slice(0, k.salt.length);
  const verifier = decrypt(at(BLOCK.verifierHash), k.salt, k.verifierHash).slice(0, k.hashSize);
  if (!equal(k.hash(input), verifier)) throw new PasswordError('wrong');
  const secret = decrypt(at(BLOCK.key), k.salt, k.keyValue).slice(0, info.keyData.keyBytes);
  const kd = info.keyData;
  if (info.integrity) {
    const hmacKey = decrypt(secret, ivFor(kd, Uint8Array.from(BLOCK.hmacKey)), info.integrity.key).slice(0, kd.hashSize);
    const expected = decrypt(secret, ivFor(kd, Uint8Array.from(BLOCK.hmacValue)), info.integrity.value).slice(0, kd.hashSize);
    if (!equal(hmac(kd.hash, hmacKey, pkg), expected)) throw new PasswordError('damaged');
  }
  const size = Number(new DataView(pkg.buffer, pkg.byteOffset, 8).getBigUint64(0, true));
  const out = new Uint8Array(Math.ceil((pkg.length - 8) / kd.blockSize) * kd.blockSize);
  for (let i = 0, o = 8; o < pkg.length; i++, o += SEGMENT) {
    const seg = pkg.subarray(o, Math.min(o + SEGMENT, pkg.length));
    out.set(decrypt(secret, ivFor(kd, u32(i)), blocks(seg, kd.blockSize)), o - 8);
  }
  streams.delete('EncryptionInfo');
  streams.delete('EncryptedPackage');
  return { bytes: out.slice(0, size), streams };
}

/** The data spaces streams Excel writes beside an encrypted package (MS-OFFCRYPTO 2.1, as Excel 365 writes them). */
function dataSpaces() {
  const b = [];
  const u32s = (...ns) => ns.forEach((n) => b.push(...u32(n)));
  const str = (s) => { u32s(s.length * 2); b.push(...utf16(s)); while (b.length % 4) b.push(0); };
  const take = () => Uint8Array.from(b.splice(0));
  // Version: "Microsoft.Container.DataSpaces", reader, updater and writer 1.0.
  str('Microsoft.Container.DataSpaces'); u32s(1, 1, 1);
  const version = take();
  // DataSpaceMap: one entry, the EncryptedPackage stream in the StrongEncryptionDataSpace.
  u32s(8, 1);
  const name = 'EncryptedPackage';
  const space = 'StrongEncryptionDataSpace';
  const entry = [];
  const push = (arr) => b.push(...arr);
  b.splice(0);
  str(name);
  const ref = Uint8Array.from(b.splice(0));
  str(space);
  const sp = Uint8Array.from(b.splice(0));
  entry.push(...u32(4 + 4 + 4 + ref.length + sp.length), ...u32(1), ...u32(0), ...ref, ...sp);
  push(u32(8)); push(u32(1)); push(entry);
  const map = take();
  // DataSpaceInfo/StrongEncryptionDataSpace: one transform, StrongEncryptionTransform.
  u32s(8, 1); str('StrongEncryptionTransform');
  const definition = take();
  // TransformInfo/StrongEncryptionTransform/\x06Primary.
  const id = '{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}';
  const header = (() => { str(id); const s = Uint8Array.from(b.splice(0)); str('Microsoft.Container.EncryptionTransform'); const n = Uint8Array.from(b.splice(0)); return concat(u32(8 + s.length), u32(1), s, n, u32(1), u32(1), u32(1)); })();
  // (Then the encryption name, empty, the block size and cipher mode, 0, and 4.)
  const primary = concat(header, u32(0), u32(0), u32(0), u32(4));
  return new Map([
    ['\u0006DataSpaces/Version', version],
    ['\u0006DataSpaces/DataSpaceMap', map],
    ['\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', definition],
    ['\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', primary],
  ]);
}

/**
 * `bytes` (an .xlsx) encrypted with `password` as Excel 365 encrypts: a compound file. `streams`: the other streams to
 * keep (the data spaces of the file it was opened from); new ones are written without them.
 */
export function encryptWorkbook(bytes, password, streams = null, { spins = 100000 } = {}) {
  const hash = sha512;
  const kd = { hash, salt: randomBytes(16), blockSize: 16, hashSize: 64 };
  const secret = randomBytes(32);
  const salt = randomBytes(16);
  const at = (blockKey) => passwordKey(hash, password, salt, spins, blockKey, 32);
  const input = randomBytes(16);
  const verifierInput = encrypt(at(BLOCK.verifierInput), salt, input);
  const verifierHash = encrypt(at(BLOCK.verifierHash), salt, blocks(hash(input), 16));
  const keyValue = encrypt(at(BLOCK.key), salt, secret);
  const pkg = new Uint8Array(8 + Math.ceil(bytes.length / 16) * 16);
  new DataView(pkg.buffer).setBigUint64(0, BigInt(bytes.length), true);
  for (let i = 0, o = 0; o < bytes.length; i++, o += SEGMENT) {
    pkg.set(encrypt(secret, ivFor(kd, u32(i)), blocks(bytes.subarray(o, Math.min(o + SEGMENT, bytes.length)), 16)), 8 + o);
  }
  const hmacKey = randomBytes(64);
  const hmacKeyOut = encrypt(secret, ivFor(kd, Uint8Array.from(BLOCK.hmacKey)), hmacKey);
  const hmacValue = encrypt(secret, ivFor(kd, Uint8Array.from(BLOCK.hmacValue)), hmac(hash, hmacKey, pkg));
  const common = 'blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"';
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<encryption xmlns="${ENCRYPTION_NS}" xmlns:p="${PASSWORD_NS}" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">`
    + `<keyData saltSize="16" ${common} saltValue="${toBase64(kd.salt)}"/>`
    + `<dataIntegrity encryptedHmacKey="${toBase64(hmacKeyOut)}" encryptedHmacValue="${toBase64(hmacValue)}"/>`
    + `<keyEncryptors><keyEncryptor uri="${PASSWORD_NS}"><p:encryptedKey spinCount="${spins}" saltSize="16" ${common} saltValue="${toBase64(salt)}"`
    + ` encryptedVerifierHashInput="${toBase64(verifierInput)}" encryptedVerifierHashValue="${toBase64(verifierHash)}" encryptedKeyValue="${toBase64(keyValue)}"/>`
    + '</keyEncryptor></keyEncryptors></encryption>';
  const info = concat(Uint8Array.from([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml));
  const out = new Map(streams && streams.size ? streams : dataSpaces());
  out.set('EncryptionInfo', info);
  out.set('EncryptedPackage', pkg);
  return writeCompoundFile(out);
}

export { isCompoundFile, dataSpaces };
