// Excel's password hashes, for Protect Sheet, Protect Workbook and Allow Edit Ranges (protection.js, editranges.js):
// Excel 2013 and later's SHA-512 over a 16-byte salt and the password (UTF-16LE), hashed again 100,000 times with the
// round's number after it, in base64; and the old 16-bit hash files from before Excel 2013 keep.
import { sha512 } from '@noble/hashes/sha2';
import { attr } from './xml.js';

export const b64 = (bytes) => btoa(String.fromCharCode(...bytes));
export const unb64 = (s) => Uint8Array.from(atob(s), (ch) => ch.charCodeAt(0));
const utf16 = (s) => { const out = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 255; out[i * 2 + 1] = s.charCodeAt(i) >> 8; } return out; };

/** A password's hash as Excel keeps a sheet's: base64 of SHA-512 over the salt and the password, `spin` times again. */
export function hashPassword(password, salt, spin = 100000) {
  let h = sha512(new Uint8Array([...salt, ...utf16(password)]));
  const buf = new Uint8Array(h.length + 4);
  for (let i = 0; i < spin; i++) {
    buf.set(h);
    buf[h.length] = i & 255; buf[h.length + 1] = (i >>> 8) & 255; buf[h.length + 2] = (i >>> 16) & 255; buf[h.length + 3] = i >>> 24;
    h = sha512(buf);
  }
  return b64(h);
}

/** Excel's old 16-bit password hash (the `password` attribute of files from before Excel 2013), 4 hex digits. */
export function legacyHash(password) {
  let h = 0;
  for (let i = password.length - 1; i >= 0; i--) { h = ((h >> 14) & 1) | ((h << 1) & 0x7fff); h ^= password.charCodeAt(i); }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= password.length ^ 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}

/** Whether `password` opens a <sheetProtection> (or with `pre` 'workbook', a <workbookProtection>): none set, Excel
 *  2013's SHA hash, or the old 16-bit one. */
export function passwordFits(el, password, pre = '') {
  const a = (name) => attr(el, pre ? pre + name[0].toUpperCase() + name.slice(1) : name);
  const hash = a('hashValue');
  if (hash) {
    if ((a('algorithmName') || '').toUpperCase() !== 'SHA-512') return false;
    return hashPassword(password, unb64(a('saltValue') || ''), Number(a('spinCount') || 100000)) === hash;
  }
  const legacy = a('password');
  if (legacy) return legacyHash(password) === legacy.toUpperCase().padStart(4, '0');
  return true;
}
