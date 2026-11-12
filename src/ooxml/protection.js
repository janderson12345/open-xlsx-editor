// Review > Protect Sheet as Excel writes it (measured over COM, Worksheet.Protect with each option and a password,
// and the files Excel saved: docs/PROTECT.md): <sheetProtection> after sheetData, sheet="1", objects="1" and
// scenarios="1" while those are protected, each thing the box allows as its attribute "0" (formatCells="0" ...), the
// selection the other way round (selectLockedCells="1": locked cells can't be selected), and a password as Excel 2013
// and later keep it: SHA-512 of a 16-byte salt and the password (UTF-16), hashed again 100,000 times with the round's
// number after it, in base64.
import { sha512 } from '@noble/hashes/sha2';
import { kid, attr } from './xml.js';
import { xfOf } from './styles.js';

/** What the box allows, in Excel's order (each its attribute, written "0" when allowed). */
export const ALLOWS = ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns',
  'deleteRows', 'sort', 'autoFilter', 'pivotTables'];

/** Excel's words when a protected sheet's locked cell is changed. */
export const PROTECTED_MESSAGE = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.";
export const WRONG_PASSWORD = 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.';

// (CT_Worksheet: sheetProtection comes before these.)
const AFTER = ['protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr',
  'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
  'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture',
  'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

const on = (v) => v === '1' || v === 'true';

/** Sheet si's protection: null when it isn't protected, else { allow: { formatCells, ... }, selectLocked,
 *  selectUnlocked, objects, scenarios, password: whether one is set }. */
export function sheetProtection(wb, si) {
  const { doc } = wb._part(si);
  const root = doc.documentElement;
  const el = kid(root, 'sheetProtection', root.namespaceURI);
  if (!el || !on(attr(el, 'sheet'))) return null;
  const allow = Object.fromEntries(ALLOWS.map((k) => [k, attr(el, k) === '0' || attr(el, k) === 'false']));
  return {
    allow,
    selectLocked: !on(attr(el, 'selectLockedCells')),
    selectUnlocked: !on(attr(el, 'selectUnlockedCells')),
    objects: on(attr(el, 'objects')),
    scenarios: on(attr(el, 'scenarios')),
    password: !!(attr(el, 'hashValue') || attr(el, 'password')),
  };
}

const b64 = (bytes) => btoa(String.fromCharCode(...bytes));
const unb64 = (s) => Uint8Array.from(atob(s), (ch) => ch.charCodeAt(0));
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

/**
 * Protect sheet si: `allow` (ALLOWS' names, true where the box lets users do it), `selectLocked`, `selectUnlocked`
 * (both on by default), `objects` and `scenarios` (protected by default), and a `password` ('' for none). `salt` is for
 * the tests (16 random bytes otherwise).
 */
export function protectSheet(wb, si, { allow = {}, selectLocked = true, selectUnlocked = true, objects = true, scenarios = true, password = '', salt = null } = {}) {
  const part = wb._part(si);
  const { doc } = part;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const old = kid(root, 'sheetProtection', ns);
  const el = doc.createElementNS(ns, root.prefix ? `${root.prefix}:sheetProtection` : 'sheetProtection');
  if (password) {
    const s = salt || globalThis.crypto.getRandomValues(new Uint8Array(16));
    el.setAttribute('algorithmName', 'SHA-512');
    el.setAttribute('hashValue', hashPassword(password, s));
    el.setAttribute('saltValue', b64(s));
    el.setAttribute('spinCount', '100000');
  }
  el.setAttribute('sheet', '1');
  if (objects) el.setAttribute('objects', '1');
  if (scenarios) el.setAttribute('scenarios', '1');
  for (const k of ALLOWS.slice(0, 8)) if (allow[k]) el.setAttribute(k, '0');
  if (!selectLocked) el.setAttribute('selectLockedCells', '1');
  for (const k of ALLOWS.slice(8)) if (allow[k]) el.setAttribute(k, '0');
  if (!selectUnlocked) el.setAttribute('selectUnlockedCells', '1');
  if (old) root.replaceChild(el, old);
  else root.insertBefore(el, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName)) || null);
  wb.dirty.add(wb.sheets[si].path);
  wb.edited = true;
}

/** Unprotect sheet si with `password`: true, or false when it is wrong (the sheet stays protected). */
export function unprotectSheet(wb, si, password = '') {
  const part = wb._part(si);
  const root = part.doc.documentElement;
  const el = kid(root, 'sheetProtection', root.namespaceURI);
  if (!el) return true;
  if (!passwordFits(el, password)) return false;
  root.removeChild(el);
  wb.dirty.add(wb.sheets[si].path);
  wb.edited = true;
  return true;
}

/** Whether `password` opens a <sheetProtection> (or with `pre` 'workbook', a <workbookProtection>): none set, Excel
 *  2013's SHA hash, or the old 16-bit one. */
function passwordFits(el, password, pre = '') {
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

/** Whether cell (r, c) of sheet si is locked by its format (its own, else its row's or column's; Excel's default
 *  locked). */
export function isLocked(wb, si, r, c) {
  const sh = wb.sheet(si);
  const cell = sh.cells.get(r)?.get(c);
  let s = cell?.s;
  if (s == null) s = sh.rows.get(r)?.s ?? sh.cols.find((x) => c + 1 >= x.min && c + 1 <= x.max)?.style ?? 0;
  return xfOf(wb.styles, s || 0).protection.locked;
}

/** Why a change to sheet si can't be made while it is protected (Excel's words), or null: `kind` is 'contents' (the
 *  `cells` [{ r, c }] changed), or one of ALLOWS. */
export function protectionRefusal(wb, si, kind, cells = []) {
  const p = sheetProtection(wb, si);
  if (!p) return null;
  if (kind === 'contents') return cells.some(({ r, c }) => isLocked(wb, si, r, c)) ? PROTECTED_MESSAGE : null;
  return p.allow[kind] ? null : PROTECTED_MESSAGE;
}

// ---- Protect Workbook (its structure): <workbookProtection lockStructure="1"/> after workbookPr, a password as a
// sheet's with "workbook" before each name (workbookAlgorithmName ...), as Excel writes it (measured).
const AFTER_BOOK = ['bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews',
  'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];

/** The workbook's protection: null, or { structure, windows, password: whether one is set }. */
export function workbookProtection(wb) {
  const root = wb.doc.documentElement;
  const el = kid(root, 'workbookProtection', root.namespaceURI);
  if (!el || !(on(attr(el, 'lockStructure')) || on(attr(el, 'lockWindows')))) return null;
  return { structure: on(attr(el, 'lockStructure')), windows: on(attr(el, 'lockWindows')), password: !!(attr(el, 'workbookHashValue') || attr(el, 'workbookPassword')) };
}

/** Protect the workbook's structure (sheets added, deleted, renamed, moved, hidden, coloured), with a `password` ('' for
 *  none); `salt` for the tests. */
export function protectWorkbook(wb, { password = '', salt = null } = {}) {
  const doc = wb.doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const old = kid(root, 'workbookProtection', ns);
  const el = doc.createElementNS(ns, root.prefix ? `${root.prefix}:workbookProtection` : 'workbookProtection');
  if (password) {
    const s = salt || globalThis.crypto.getRandomValues(new Uint8Array(16));
    el.setAttribute('workbookAlgorithmName', 'SHA-512');
    el.setAttribute('workbookHashValue', hashPassword(password, s));
    el.setAttribute('workbookSaltValue', b64(s));
    el.setAttribute('workbookSpinCount', '100000');
  }
  el.setAttribute('lockStructure', '1');
  if (old) root.replaceChild(el, old);
  else root.insertBefore(el, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_BOOK.includes(n.localName)) || null);
  wb.bookDirty = true;
  wb.edited = true;
}

/** Unprotect the workbook with `password`: true, or false when it is wrong. */
export function unprotectWorkbook(wb, password = '') {
  const root = wb.doc.documentElement;
  const el = kid(root, 'workbookProtection', root.namespaceURI);
  if (!el) return true;
  if (!passwordFits(el, password, 'workbook')) return false;
  root.removeChild(el);
  wb.bookDirty = true;
  wb.edited = true;
  return true;
}

/** Said when a change to a protected workbook's sheets is refused. (Excel greys those commands out; its words for
 *  the refusal couldn't be read over COM, which gives only a general error.) */
export const WORKBOOK_PROTECTED = 'Workbook is protected and cannot be changed.';
