// A workbook's sheets changed as Excel changes them: added, renamed, moved, hidden and shown, given a tab colour,
// deleted. References stay right: a rename rewrites every formula, defined name, chart series, hyperlink and pivot
// source that names the sheet; a delete turns references to it into #REF! (in the file, the whole reference, as Excel
// writes it) and takes away the names that belong to it. Names that belong to a sheet follow it (localSheetId), as do
// the workbook's active and first tab. What Excel does here was measured over COM (docs/SHEETS.md).
import { kid, kids, attr, serializeXml, REL_TYPES } from './xml.js';
import { renameSheetRefs } from '../calc/tokens.js';

/** Excel's own words when it refuses. */
export const MESSAGES = {
  invalid: 'You typed an invalid name for a sheet or chart. Make sure that:\n' +
    '• The name that you type does not exceed 31 characters.\n' +
    '• The name does not contain any of the following characters: : \\ / ? * [ or ]\n' +
    '• You did not leave the name blank.',
  taken: 'That name is already taken. Try a different one.',
  reserved: 'History is a reserved name.',
  lastVisible: 'A workbook must contain at least one visible worksheet.',
};

/** A sheet change Excel refuses; `message` is Excel's. */
export class SheetError extends Error {}

/** An entry with the XML element it came from, kept out of sight (not enumerable). */
export const withEl = (o, el) => Object.defineProperty(o, 'el', { value: el, writable: true });

const WORKSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const lower = (s) => s.toLowerCase();
const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** Excel's checks of a sheet name (sheet `except` is the one being renamed); throws a SheetError. A blank name, one
 *  longer than 31 characters, with : \ / ? * [ or ], or starting or ending with an apostrophe is invalid; "History"
 *  is reserved; names are compared without regard to case. */
export function checkSheetName(wb, name, except = -1) {
  if (typeof name !== 'string' || !name.length || name.length > 31 || /[:\\/?*[\]]/.test(name) || name.startsWith("'") || name.endsWith("'")) {
    throw new SheetError(MESSAGES.invalid);
  }
  if (lower(name) === 'history') throw new SheetError(MESSAGES.reserved);
  if (wb.sheets.some((s, i) => i !== except && lower(s.name) === lower(name))) throw new SheetError(MESSAGES.taken);
}

const visibleCount = (wb) => wb.sheets.filter((s) => s.state === 'visible').length;
const bookView = (wb) => kid(kid(wb.doc.documentElement, 'bookViews', wb.ns), 'workbookView', wb.ns);

/** After sheets were added, moved or deleted: `map(old index)` is the new index, or -1 for a deleted sheet. Names that
 *  belonged to a deleted sheet go; the others follow their sheet; the active tab follows its sheet (or, gone or hidden,
 *  becomes the nearest sheet that can be shown). */
function reindex(wb, map) {
  const root = wb.doc.documentElement;
  for (const d of [...wb.definedNames]) {
    if (d.sheet == null) continue;
    const to = map(d.sheet);
    if (to < 0) {
      d.el.parentNode.removeChild(d.el);
      wb.definedNames.splice(wb.definedNames.indexOf(d), 1);
    } else if (to !== d.sheet) {
      d.sheet = to;
      d.el.setAttribute('localSheetId', String(to));
    }
  }
  const names = kid(root, 'definedNames', wb.ns);
  if (names && !kids(names, 'definedName', wb.ns).length) root.removeChild(names);
  const view = bookView(wb);
  if (view) {
    const old = Number(attr(view, 'activeTab') || 0);
    let active = map(old);
    if (active < 0 || !wb.canShow(active)) active = nearest(wb, active < 0 ? Math.min(old, wb.sheets.length - 1) : active);
    if (active !== old || attr(view, 'activeTab') != null) view.setAttribute('activeTab', String(active));
    if (attr(view, 'firstSheet') != null) {
      const first = map(Number(attr(view, 'firstSheet')));
      view.setAttribute('firstSheet', String(Math.max(0, Math.min(first < 0 ? active : first, active))));
    }
  }
  wb.resetCalc();
  wb.bookDirty = true;
}

/** The sheet that can be shown nearest to index i: i itself, else the next one to the right, else to the left (as
 *  Excel shows the next tab when the one shown goes). */
export function nearest(wb, i) {
  for (let k = Math.max(0, i); k < wb.sheets.length; k++) if (wb.canShow(k)) return k;
  for (let k = Math.min(i, wb.sheets.length) - 1; k >= 0; k--) if (wb.canShow(k)) return k;
  return 0;
}

const FORMULA_ELEMENTS = ['f', 'formula', 'formula1', 'formula2'];

/** Every reference to sheet `from` made to point at `to` (null: #REF!), in the worksheets (cells, conditional formats,
 *  data validations, hyperlinks), the defined names, the charts and, for a rename, the pivot caches' sources. Returns
 *  the formula cells whose text changed, [{ si, r, c }]. */
function rewriteRefs(wb, from, to) {
  const map = (text) => renameSheetRefs(text, from, to);
  const changed = [];
  // (A part that doesn't mention the name, escaped as the XML has it, can't refer to the sheet: it isn't parsed.)
  const xmlName = lower(from).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, "''");
  const mentions = (path) => {
    const text = lower(wb.pkg.text(path) || '');
    return text.includes(xmlName) || text.includes(xmlName.replace(/'/g, '&apos;'));
  };
  wb.sheets.forEach((info, si) => {
    if (info.kind !== 'worksheet' || !info.path || !wb.pkg.has(info.path)) return;
    if (!wb._parsed.has(info) && !mentions(info.path)) return;
    const { doc, model } = wb._part(si);
    let dirty = false;
    for (const name of FORMULA_ELEMENTS) {
      for (const el of [...doc.getElementsByTagNameNS('*', name)]) {
        const text = el.textContent;
        const next = text ? map(text) : text;
        if (next !== text) { el.textContent = next; dirty = true; }
      }
    }
    for (const el of [...doc.getElementsByTagNameNS('*', 'hyperlink')]) {
      const loc = el.getAttribute('location');
      const next = loc ? map(loc) : loc;
      if (next !== loc) { el.setAttribute('location', next); dirty = true; }
    }
    for (const [r, line] of model.cells) {
      for (const [c, cell] of line) {
        if (cell.f == null || cell.f === '') continue;
        const next = map(cell.f);
        if (next !== cell.f) { cell.f = next; changed.push({ si, r, c }); }
      }
    }
    if (dirty) wb.dirty.add(info.path);
  });
  for (const d of wb.definedNames) {
    const next = map(d.text);
    if (next !== d.text) { d.text = next; d.el.textContent = next; wb.bookDirty = true; }
  }
  for (const path of Object.keys(wb.pkg.files)) {
    const chart = /^xl\/charts\/chart[^/]*\.xml$/i.test(path);
    const pivot = to != null && /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i.test(path);
    if (!(chart || pivot) || !mentions(path)) continue;
    const doc = wb.pkg.xml(path);
    let dirty = false;
    if (chart) {
      for (const el of [...doc.getElementsByTagNameNS('*', 'f')]) {
        const next = map(el.textContent);
        if (next !== el.textContent) { el.textContent = next; dirty = true; }
      }
    } else {
      for (const el of [...doc.getElementsByTagNameNS('*', 'worksheetSource')]) {
        if (lower(el.getAttribute('sheet') || '') === lower(from)) { el.setAttribute('sheet', to); dirty = true; }
      }
    }
    if (dirty) wb.pkg.setText(path, keepDeclaration(serializeXml(doc), wb.pkg.text(path)));
  }
  return changed;
}

/** Serialised XML with the original's declaration in front. */
function keepDeclaration(body, original) {
  const decl = /^\uFEFF?(<\?xml[^?]*\?>)/.exec(original || '');
  const rest = body.replace(/^<\?xml[^?]*\?>\s*/, '');
  return decl ? `${decl[1]}\r\n${rest}` : rest;
}

/** A new, empty worksheet at index `at` (Excel's "+" puts it after the sheet shown, Insert before it), named `name` or
 *  the next free SheetN. Written as Excel writes a new sheet: an empty sheetData, the workbook's default row height,
 *  Excel's default margins. Returns its index. */
export function addSheet(wb, at = wb.sheets.length, name = null) {
  if (name == null) {
    do name = `Sheet${wb.nextSheetNumber++}`; while (wb.sheets.some((s) => lower(s.name) === lower(name)));
  } else checkSheetName(wb, name);
  const pkg = wb.pkg;
  const sibling = wb.sheets.find((s) => s.kind === 'worksheet' && s.path && pkg.has(s.path));
  const dir = sibling ? sibling.path.slice(0, sibling.path.lastIndexOf('/') + 1) : `${pkg.mainDir}worksheets/`;
  let k = 1;
  while (pkg.has(`${dir}sheet${k}.xml`)) k++;
  const path = `${dir}sheet${k}.xml`;
  const row = sibling ? /defaultRowHeight="([0-9.]+)"/.exec(new TextDecoder().decode(pkg.files[sibling.path].subarray(0, 8192)))?.[1] : null;
  // (The relationships namespace and type are the workbook's own: Strict has its own.)
  const relAttr = wb.sheets.length ? [...wb.sheets[0].el.attributes].find((a) => a.localName === 'id' && a.namespaceURI) : null;
  const relNs = relAttr ? relAttr.namespaceURI : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const relType = pkg.rels.find((r) => r.type.endsWith('/worksheet'))?.type || REL_TYPES.worksheet;
  pkg.setText(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<worksheet xmlns="${wb.ns}" xmlns:r="${relNs}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="${row || 15}"/><sheetData/>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
  const target = path.startsWith(pkg.mainDir) ? path.slice(pkg.mainDir.length) : `/${path}`;
  const rId = pkg.addRel(relType, target);
  pkg.addOverride(path, WORKSHEET_CT);
  const doc = wb.doc;
  const el = make(doc, wb.ns, 'sheet');
  el.setAttribute('name', name);
  const sheetId = Math.max(0, ...wb.sheets.map((s) => s.sheetId || 0)) + 1;
  el.setAttribute('sheetId', String(sheetId));
  el.setAttributeNS(relNs, `${relAttr?.prefix || doc.documentElement.lookupPrefix(relNs) || 'r'}:id`, rId);
  const container = kid(doc.documentElement, 'sheets', wb.ns);
  container.insertBefore(el, wb.sheets[at]?.el || null);
  wb.sheets.splice(at, 0, withEl({ name, sheetId, state: 'visible', rId, path, kind: 'worksheet' }, el));
  reindex(wb, (i) => (i >= at ? i + 1 : i));
  return at;
}

/** Rename sheet i; every reference to it follows. */
export function renameSheet(wb, i, name) {
  checkSheetName(wb, name, i);
  const info = wb.sheets[i];
  const old = info.name;
  if (old === name) return;
  info.name = name;
  info.el.setAttribute('name', name);
  wb.bookDirty = true;
  rewriteRefs(wb, old, name);
  wb.resetCalc();
  wb.edited = true;
}

/** Move sheet `from` to index `to` (its index once moved). Formulas don't change: they name sheets, not positions. */
export function moveSheet(wb, from, to) {
  if (from === to) return;
  const before = [...wb.sheets];
  const [info] = wb.sheets.splice(from, 1);
  wb.sheets.splice(to, 0, info);
  const container = info.el.parentNode;
  container.removeChild(info.el);
  container.insertBefore(info.el, wb.sheets[to + 1]?.el || null);
  reindex(wb, (k) => wb.sheets.indexOf(before[k]));
}

/** Hide or show sheet i's tab. The last sheet shown can't be hidden. A hidden sheet loses its tab selection (a selected
 *  tab that is hidden confuses Excel's grouping), and the active tab moves off it. */
export function setSheetHidden(wb, i, hidden) {
  const info = wb.sheets[i];
  if (hidden === (info.state !== 'visible')) return;
  if (hidden && visibleCount(wb) <= 1) throw new SheetError(MESSAGES.lastVisible);
  info.state = hidden ? 'hidden' : 'visible';
  if (hidden) info.el.setAttribute('state', 'hidden');
  else info.el.removeAttribute('state');
  if (hidden && info.kind === 'worksheet' && info.path && wb.pkg.has(info.path)) {
    const root = wb._part(i).doc.documentElement;
    const view = kid(kid(root, 'sheetViews', root.namespaceURI), 'sheetView', root.namespaceURI);
    if (view && view.getAttribute('tabSelected')) { view.removeAttribute('tabSelected'); wb.dirty.add(info.path); }
  }
  reindex(wb, (k) => k);
}

/** Sheet i's tab colour: { rgb: 'FFRRGGBB' }, { theme, tint } (a theme colour, as Excel writes its palette's), or null
 *  for none. Kept in the sheet's sheetPr, which comes first in the worksheet. */
export function setTabColor(wb, i, color) {
  const info = wb.sheets[i];
  const { doc } = wb._part(i);
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let pr = kid(root, 'sheetPr', ns);
  let tab = kid(pr, 'tabColor', ns);
  if (!color) {
    if (!tab) return;
    pr.removeChild(tab);
    if (!pr.attributes.length && ![...pr.childNodes].some((n) => n.nodeType === 1)) root.removeChild(pr);
  } else {
    if (!pr) { pr = make(doc, ns, 'sheetPr'); root.insertBefore(pr, [...root.childNodes].find((n) => n.nodeType === 1) || null); }
    if (!tab) { tab = make(doc, ns, 'tabColor'); pr.insertBefore(tab, [...pr.childNodes].find((n) => n.nodeType === 1) || null); }
    for (const a of [...tab.attributes]) tab.removeAttribute(a.name);
    if (color.rgb) tab.setAttribute('rgb', color.rgb);
    else {
      tab.setAttribute('theme', String(color.theme));
      if (color.tint) tab.setAttribute('tint', String(color.tint));
    }
  }
  wb.dirty.add(info.path);
}

/** Delete sheet i, its parts and the parts only it led to (drawings, comments, tables). The last sheet shown can't be
 *  deleted. References to it become #REF!, its own names go, and the workbook is worked out again. */
export function deleteSheet(wb, i) {
  const info = wb.sheets[i];
  if (info.state === 'visible' && visibleCount(wb) <= 1) throw new SheetError(MESSAGES.lastVisible);
  wb.sheets.splice(i, 1);
  info.el.parentNode.removeChild(info.el);
  reindex(wb, (k) => (k === i ? -1 : k > i ? k - 1 : k));
  rewriteRefs(wb, info.name, null);
  if (info.path) {
    wb.pkg.removeTree(info.path);
    wb.dirty.delete(info.path);
  }
  wb._parsed.delete(info);
  // (calcChain.xml names cells by sheetId: some of them are gone.)
  wb._dropCalcChain();
  wb.edited = true;
  wb.calculate();
}
