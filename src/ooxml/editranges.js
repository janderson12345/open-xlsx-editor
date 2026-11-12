// Review > Allow Edit Ranges as Excel writes them (measured over COM, Protection.AllowEditRanges.Add, and the file
// Excel saved: docs/PROTECT.md): <protectedRanges> after <sheetProtection>, a <protectedRange sqref name> for each
// (several areas apart by spaces), a password hashed as a sheet's. On a protected sheet a range's cells change though
// they are locked: one without a password at once, one with a password once it is unlocked, until the workbook is
// opened again (Excel doesn't keep the unlocking). While the sheet is protected, no range is added or changed.
import { kid, kids, attr } from './xml.js';
import { parseRange } from './refs.js';
import { b64, hashPassword, passwordFits } from './passwords.js';

// (CT_Worksheet: sheetProtection comes before these, protectedRanges first.)
export const AFTER_PROTECTION = ['protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr',
  'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
  'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture',
  'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

const rootOf = (wb, si) => wb._part(si).doc.documentElement;
const listOf = (wb, si) => { const root = rootOf(wb, si); return kid(root, 'protectedRanges', root.namespaceURI); };
const touched = (wb, si) => { wb.dirty.add(wb.sheets[si].path); wb.edited = true; };
const areasOf = (sqref) => String(sqref || '').split(/\s+/).map(parseRange).filter(Boolean);
/** The ranges unlocked with their passwords this session, by sheet: not saved, as in Excel. */
const unlockedOf = (wb, si) => {
  wb._unlockedRanges ||= new Map();
  if (!wb._unlockedRanges.has(si)) wb._unlockedRanges.set(si, new Set());
  return wb._unlockedRanges.get(si);
};

/** Sheet si's Allow Edit Ranges, in order: [{ name, sqref, areas: [{ r0, c0, r1, c1 }], password: whether it has
 *  one, unlocked: whether its password was given this session }]. */
export function editRanges(wb, si) {
  const list = listOf(wb, si);
  if (!list) return [];
  const unlocked = unlockedOf(wb, si);
  return kids(list, 'protectedRange', list.namespaceURI).map((el) => {
    const name = attr(el, 'name') || '';
    return { name, sqref: attr(el, 'sqref') || '', areas: areasOf(attr(el, 'sqref')), password: !!(attr(el, 'hashValue') || attr(el, 'password')), unlocked: unlocked.has(name) };
  });
}

/** Add a range `name` over `sqref` ("B2:C3", or areas apart by spaces) to sheet si, with a `password` ('' for none);
 *  `salt` for the tests. A range of that name is replaced. */
export function addEditRange(wb, si, { name, sqref, password = '', salt = null }) {
  const root = rootOf(wb, si);
  const ns = root.namespaceURI;
  const make = (tag) => root.ownerDocument.createElementNS(ns, root.prefix ? `${root.prefix}:${tag}` : tag);
  let list = listOf(wb, si);
  if (!list) {
    list = make('protectedRanges');
    root.insertBefore(list, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_PROTECTION.slice(1).includes(n.localName)) || null);
  }
  const el = make('protectedRange');
  if (password) {
    const s = salt || globalThis.crypto.getRandomValues(new Uint8Array(16));
    el.setAttribute('algorithmName', 'SHA-512');
    el.setAttribute('hashValue', hashPassword(password, s));
    el.setAttribute('saltValue', b64(s));
    el.setAttribute('spinCount', '100000');
  }
  el.setAttribute('sqref', sqref);
  el.setAttribute('name', name);
  const old = kids(list, 'protectedRange', ns).find((x) => attr(x, 'name') === name);
  if (old) list.replaceChild(el, old);
  else list.appendChild(el);
  unlockedOf(wb, si).delete(name);
  touched(wb, si);
}

/** Delete sheet si's range `name` (the list too, once empty). */
export function deleteEditRange(wb, si, name) {
  const list = listOf(wb, si);
  const el = list && kids(list, 'protectedRange', list.namespaceURI).find((x) => attr(x, 'name') === name);
  if (!el) return false;
  list.removeChild(el);
  if (!kids(list, 'protectedRange', list.namespaceURI).length) list.parentNode.removeChild(list);
  unlockedOf(wb, si).delete(name);
  touched(wb, si);
  return true;
}

/** Unlock sheet si's range `name` with its `password` for this session: true, or false when it is wrong. */
export function unlockEditRange(wb, si, name, password) {
  const list = listOf(wb, si);
  const el = list && kids(list, 'protectedRange', list.namespaceURI).find((x) => attr(x, 'name') === name);
  if (!el || !passwordFits(el, password)) return false;
  unlockedOf(wb, si).add(name);
  return true;
}

/** Whether cell (r, c) of sheet si can change in a range: one without a password, or one unlocked. */
export function inOpenRange(wb, si, r, c) {
  return editRanges(wb, si).some((x) => (!x.password || x.unlocked) && x.areas.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1));
}

/** The range with a password, still locked, that cell (r, c) of sheet si is in (its name), or null: the one the
 *  editor asks its password for. */
export function lockedRangeAt(wb, si, r, c) {
  return editRanges(wb, si).find((x) => x.password && !x.unlocked && x.areas.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1))?.name ?? null;
}
