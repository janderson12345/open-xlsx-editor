// Review > Protect Sheet as Excel writes it (measured over COM, Worksheet.Protect with each option and a password,
// and the files Excel saved: docs/PROTECT.md): <sheetProtection> after sheetData, sheet="1", objects="1" and
// scenarios="1" while those are protected, each thing the box allows as its attribute "0" (formatCells="0" ...), the
// selection the other way round (selectLockedCells="1": locked cells can't be selected), and a password as Excel 2013
// and later keep it: SHA-512 of a 16-byte salt and the password (UTF-16), hashed again 100,000 times with the round's
// number after it, in base64.
import { kid, attr } from './xml.js';
import { b64, hashPassword, passwordFits } from './passwords.js';
import { inOpenRange, AFTER_PROTECTION as AFTER } from './editranges.js';
import { xfOf } from './styles.js';
import { lineStyle } from './sheet.js';

export { hashPassword, legacyHash } from './passwords.js';

/** What the box allows, in Excel's order (each its attribute, written "0" when allowed). */
export const ALLOWS = ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns',
  'deleteRows', 'sort', 'autoFilter', 'pivotTables'];

/** Excel's words when a protected sheet's locked cell is changed. */
export const PROTECTED_MESSAGE = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password.";
export const WRONG_PASSWORD = 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.';


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

/** Cell (r, c) of sheet si's protection by its format (its own, else its row's or column's): { locked (Excel's
 *  default), hidden }. */
function cellProtection(wb, si, r, c) {
  const sh = wb.sheet(si);
  const cell = sh.cells.get(r)?.get(c);
  let s = cell?.s;
  if (s == null) s = lineStyle(sh, r, c);
  return xfOf(wb.styles, s || 0).protection;
}

/** Whether cell (r, c) of sheet si is locked by its format. */
export function isLocked(wb, si, r, c) {
  return cellProtection(wb, si, r, c).locked;
}

/** Whether cell (r, c)'s contents are hidden: its format's Hidden (Format Cells > Protection) on a protected sheet.
 *  Then Excel's formula bar shows nothing, Show Formulas shows a formula as nothing (a constant as itself), and
 *  FORMULATEXT gives #N/A (measured). */
export function contentsHidden(wb, si, r, c) {
  return !!sheetProtection(wb, si) && cellProtection(wb, si, r, c).hidden;
}

/** Why a change to sheet si can't be made while it is protected (Excel's words), or null: `kind` is 'contents' (the
 *  `cells` [{ r, c }] changed), 'objects' (a new note; a chart moved, sized, changed or deleted: Edit objects), or one
 *  of ALLOWS; anything else is refused while the sheet is protected. */
export function protectionRefusal(wb, si, kind, cells = []) {
  const p = sheetProtection(wb, si);
  if (!p) return null;
  // (A locked cell changes in an Allow Edit Range without a password, or one unlocked: editranges.js.)
  if (kind === 'contents') return cells.some(({ r, c }) => isLocked(wb, si, r, c) && !inOpenRange(wb, si, r, c)) ? PROTECTED_MESSAGE : null;
  if (kind === 'objects') return p.objects ? PROTECTED_MESSAGE : null;
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

/**
 * Where a key takes the active cell on a protected sheet, as Excel's do (measured in its window, docs/PROTECT.md), or
 * undefined where the key moves as usual. `how`: 'arrow' (with `dr`, `dc`), 'tab' (dc ±1) or 'enter' (dr ±1).
 * With Select unlocked cells off (selectUnlockedCells), arrows and clicks do nothing; with Select locked cells off,
 * Right and Left go to the next or last unlocked cell in reading order (rows, wrapping round), Down and Up down the
 * columns; Tab always goes through the unlocked cells in rows, and Enter down the columns once locked cells can't be
 * selected. Returns { r, c }, null (stay), or undefined.
 */
export function protectedStep(wb, si, from, dr, dc, how) {
  const p = sheetProtection(wb, si);
  if (!p) return undefined;
  const none = !p.selectUnlocked;
  const unlockedOnly = !p.selectLocked;
  if (how === 'arrow' && none) return null;
  const rows = how === 'tab' || (how === 'arrow' && unlockedOnly && dc !== 0);
  const cols = (how === 'enter' && (none || unlockedOnly)) || (how === 'arrow' && unlockedOnly && dr !== 0);
  if (!rows && !cols) return undefined;
  const cells = [];
  for (const [r, line] of wb.sheet(si).cells) for (const c of line.keys()) if (!isLocked(wb, si, r, c)) cells.push({ r, c });
  if (!cells.length) return null;
  const key = rows ? (a) => a.r * 1e5 + a.c : (a) => a.c * 1e7 + a.r;
  cells.sort((a, b) => key(a) - key(b));
  const back = dr < 0 || dc < 0;
  const at = key(from);
  const next = back ? [...cells].reverse().find((x) => key(x) < at) : cells.find((x) => key(x) > at);
  return next || (back ? cells[cells.length - 1] : cells[0]);
}

/** Whether cell (r, c) can be selected on sheet si: not on a protected sheet that allows no selection, or only
 *  unlocked cells and this one locked (Excel ignores a click there). */
export function canSelect(wb, si, r, c) {
  const p = sheetProtection(wb, si);
  if (!p) return true;
  if (!p.selectUnlocked) return false;
  return p.selectLocked || !isLocked(wb, si, r, c);
}

