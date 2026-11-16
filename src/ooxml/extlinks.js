// Links to other workbooks, as Excel writes them (measured over COM: tools/excel-compare/excel-extlinks.ps1 and
// excel-breaklink.ps1; docs/LINKS.md). Each is a part xl/externalLinks/externalLinkN.xml listed in the workbook's
// <externalReferences>: the other workbook's file (an externalLinkPath relationship), its sheets' names, its defined
// names, and the values of the cells this workbook reads (sheetDataSet), which formulas show until the link is
// updated. A formula writes the link's place in that list: [1]Sheet1!A1, '[1]Data Two'!$A$1, [1]!Rate; it is shown
// and typed with the file's name instead, [Book.xlsx]Sheet1!A1 (bookNames, bookIndexes). A part is read again only
// when its bytes change, and written only when something changes in it, so a workbook's links not touched are kept
// byte for byte.
import { parseXml, serializeXml, kids, kid, attr, NS } from './xml.js';
import { refName, parseRef } from './refs.js';
import { mapRefs, sheetPrefix, tokenize } from '../calc/tokens.js';
import { Arr, errorOf, E, isError } from '../calc/values.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const LINK_REL = `${REL}/externalLink`;
const PATH_REL = `${REL}/externalLinkPath`;
const CT_LINK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml';
// (CT_Workbook's order: externalReferences after sheets and functionGroups, before definedNames and the rest.)
const AFTER_REFS = ['definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const memo = new WeakMap();

/** A file name without its folder: C:\Data\Book.xlsx, file:///C:/x/Book.xlsx and Book.xlsx are Book.xlsx. */
export const fileOf = (target) => decodeURI(String(target || '')).replace(/^.*[\\/]/, '');

const rid = (el) => [...el.attributes].find((a) => a.localName === 'id' && a.namespaceURI !== null)?.value || null;

/** One link's part read: { sheetNames, names: [{ name, refersTo, sheetId }], cells: [Map 'r,c' -> { v, t }] by sheet,
 *  errors: [sheet ids with refreshError], rId of its file }. */
function readPart(text) {
  const doc = parseXml(text);
  const book = kid(doc.documentElement, 'externalBook');
  if (!book) return { doc, book: null, sheetNames: [], names: [], cells: [], errors: new Set(), fileRid: null };
  const sheetNames = kids(kid(book, 'sheetNames'), 'sheetName').map((el) => attr(el, 'val'));
  const names = kids(kid(book, 'definedNames'), 'definedName').map((el) => ({ name: attr(el, 'name'), refersTo: attr(el, 'refersTo') || '', sheetId: attr(el, 'sheetId') == null ? null : Number(attr(el, 'sheetId')) }));
  const cells = sheetNames.map(() => new Map());
  const errors = new Set();
  for (const sd of kids(kid(book, 'sheetDataSet'), 'sheetData')) {
    const id = Number(attr(sd, 'sheetId'));
    if (attr(sd, 'refreshError') === '1') errors.add(id);
    if (!cells[id]) cells[id] = new Map();
    for (const row of kids(sd, 'row')) {
      for (const c of kids(row, 'cell')) {
        const { r, c: col } = parseRef(attr(c, 'r'));
        const v = kid(c, 'v');
        if (v) cells[id].set(`${r},${col}`, { v: v.textContent, t: attr(c, 't') || 'n' });
      }
    }
  }
  return { doc, book, sheetNames, names, cells, errors, fileRid: rid(book) };
}

/** The workbook's links, in the order formulas number them: [{ index (1 up), path, target (the file as its
 *  relationship gives it), name (the file's name alone), sheetNames, names, cells, errors }]. */
export function readLinks(wb) {
  const refs = kids(kid(wb.doc.documentElement, 'externalReferences', wb.ns), 'externalReference', wb.ns);
  return refs.map((el, k) => {
    const rel = wb.pkg.rels.find((x) => x.id === rid(el));
    const path = rel && !rel.external ? rel.target : null;
    const bytes = path && wb.pkg.files[path];
    if (!bytes) return { index: k + 1, path, target: null, name: '', sheetNames: [], names: [], cells: [], errors: new Set() };
    let m = memo.get(bytes);
    if (!m) { m = readPart(wb.pkg.text(path)); memo.set(bytes, m); }
    // (Its file from its relationships, as they are now: Change Source changes them alone.)
    const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    let file = memo.get(wb.pkg.files[relsPath]);
    if (!file && wb.pkg.files[relsPath]) { file = wb.pkg.readRels(path); memo.set(wb.pkg.files[relsPath], file); }
    const target = (file || []).find((x) => x.id === m.fileRid)?.rawTarget ?? null;
    return { index: k + 1, path, target, name: fileOf(target), sheetNames: m.sheetNames, names: m.names, cells: m.cells, errors: m.errors };
  });
}

/** The link a formula's book stands for: its number ("1") or its file's name ("Book.xlsx", capitals aside). */
export function linkOf(wb, book) {
  if (book == null) return null;
  const links = readLinks(wb);
  if (/^\d+$/.test(book)) return links[Number(book) - 1] || null;
  const want = fileOf(book).toLowerCase();
  return links.find((l) => l.name.toLowerCase() === want) || null;
}

/** A cached cell's value as a formula reads it. */
function cachedValue(x) {
  if (!x) return null;
  if (x.t === 'b') return x.v === '1' || x.v === 'true';
  if (x.t === 'e') return errorOf(x.v);
  if (x.t === 'str' || x.t === 's' || x.t === 'inlineStr') return x.v;
  const n = Number(x.v);
  return Number.isFinite(n) ? n : x.v;
}

/** What a reference to another workbook reads, from the values kept for it: one cell's value (null when blank), an
 *  Arr for a range, #REF! for a sheet the link doesn't have; undefined when there is no such link. */
export function externalValue(wb, book, sheet, area) {
  const link = linkOf(wb, book);
  if (!link) return undefined;
  const id = link.sheetNames.findIndex((n) => n.toLowerCase() === String(sheet || '').toLowerCase());
  if (id < 0 || !area || link.errors.has(id)) return E.ref;
  const cells = link.cells[id] || new Map();
  const at = (r, c) => cachedValue(cells.get(`${r},${c}`));
  if (area.r0 === area.r1 && area.c0 === area.c1) return at(area.r0, area.c0);
  // (A whole column or row: only as far as the values kept go.)
  const keys = [...cells.keys()].map((k) => k.split(',').map(Number));
  const r1 = Math.min(area.r1, Math.max(area.r0, ...keys.map(([r]) => r)));
  const c1 = Math.min(area.c1, Math.max(area.c0, ...keys.map(([, c]) => c)));
  const rows = [];
  for (let r = area.r0; r <= r1; r++) {
    const line = [];
    for (let c = area.c0; c <= c1; c++) line.push(at(r, c));
    rows.push(line);
  }
  return new Arr(rows);
}

/** An external workbook's defined name ([1]!Rate) as a formula's text of that workbook, or null: its refersTo, the
 *  sheets in it written with the link ('Sheet1'!$B$1 is [1]Sheet1!$B$1), so it reads the values kept. */
export function externalNameFormula(wb, book, name) {
  const link = linkOf(wb, book);
  const n = link?.names.find((x) => x.name.toLowerCase() === String(name).toLowerCase() && x.sheetId == null);
  if (!n) return null;
  const text = n.refersTo.replace(/^=/, '');
  try {
    return mapRefs(text, (t) => (t.type === 'ref' && t.book == null && t.sheet != null && !t.sheets ? bookPrefix(link.index, t.sheet) + t.text.slice(t.prefix.length) : null));
  } catch {
    return null;
  }
}

/** A reference's prefix for sheet `sheet` of link `index`: [1]Sheet1!, or '[1]Data Two'! when the name needs quotes. */
export function bookPrefix(index, sheet) {
  const plain = sheetPrefix(sheet);
  return plain.startsWith("'") ? `'[${index}]${plain.slice(1)}` : `[${index}]${plain}`;
}

/** A formula as the file has it shown with its links' file names: [1]Sheet1!A1 as [Book.xlsx]Sheet1!A1, '[1]My
 *  Data'!A1 as '[Book.xlsx]My Data'!A1, [1]!Rate as [Book.xlsx]!Rate. A link not found is left as written. */
export function bookNames(wb, text) {
  if (!text || !text.includes('[')) return text;
  const links = readLinks(wb);
  if (!links.length) return text;
  try {
    return mapRefs(text, (t) => {
      if (t.book == null || !/^\d+$/.test(t.book)) return null;
      const link = links[Number(t.book) - 1];
      if (!link?.name) return null;
      if (t.type === 'name') return `[${link.name}]!${t.name}`;
      if (t.type !== 'ref' || t.sheet == null) return null;
      const body = t.text.slice(t.prefix.length);
      const plain = sheetPrefix(`[${link.name}]${t.sheet}`);
      // ([Book.xlsx]Sheet1! unquoted when the sheet's name and the file's need no quotes.)
      const bare = /^[\w.-]+$/.test(link.name) && !sheetPrefix(t.sheet).startsWith("'");
      return (bare ? `[${link.name}]${t.sheet}!` : plain) + body;
    });
  } catch {
    return text;
  }
}

/** A typed formula's references to other workbooks by file name ([Book.xlsx]Sheet1!A1) written with their links'
 *  numbers, as the file keeps them; a workbook not linked yet gets a link (addLink: its sheets from `find(name)`, the
 *  workbook open with that name, if any). Returns the formula as the file writes it. */
export function bookIndexes(wb, text, find = wb.openBook) {
  if (!text || !text.includes('[')) return text;
  return mapRefs(text, (t) => {
    if (t.book == null || /^\d+$/.test(t.book)) return null;
    let link = linkOf(wb, t.book);
    if (!link) {
      const src = find?.(t.book) || null;
      // (One not open here: its values can't be read, refreshError, so it reads #REF! until it is updated.)
      const index = addLink(wb, { target: t.book, sheetNames: src ? src.sheets.map((s) => s.name) : [t.sheet].filter(Boolean), names: src ? bookNamesOf(src) : [], unread: !src });
      link = readLinks(wb)[index - 1];
    }
    if (t.type === 'name') return `[${link.index}]!${t.name}`;
    if (t.type !== 'ref' || t.sheet == null) return null;
    return bookPrefix(link.index, t.sheet) + t.text.slice(t.prefix.length);
  });
}

/** The cells a link's references read, by sheet name (capitals aside): Map name -> Set 'r,c', from every formula of
 *  the workbook and its defined names. Whole columns and rows count as far as `limit` (the source's used range). */
export function referencedCells(wb, index, limit = () => ({ maxRow: 0, maxCol: 0 })) {
  const out = new Map();
  const add = (sheet, a) => {
    if (!a) return;
    const key = sheet.toLowerCase();
    if (!out.has(key)) out.set(key, { sheet, cells: new Set() });
    const lim = limit(sheet);
    const r1 = a.kind === 'cols' ? Math.min(a.r1, lim.maxRow) : a.r1;
    const c1 = a.kind === 'rows' ? Math.min(a.c1, lim.maxCol) : a.c1;
    for (let r = a.r0; r <= r1 && r - a.r0 < 100000; r++) for (let c = a.c0; c <= c1; c++) out.get(key).cells.add(`${r},${c}`);
  };
  const scan = (text) => {
    if (!text || !text.includes('[')) return;
    try {
      for (const t of tokenize(text)) {
        if (t.type === 'ref' && t.book === String(index) && t.sheet != null) add(t.sheet, t.area);
      }
    } catch { /* a formula that doesn't read: nothing */ }
  };
  wb.sheets.forEach((info, k) => {
    if (info.kind !== 'worksheet' || !info.path || !wb.pkg.has(info.path)) return;
    for (const [, line] of wb.sheet(k).cells) for (const [, cell] of line) if (cell.f) scan(cell.f);
  });
  for (const d of wb.definedNames) scan(d.text);
  return out;
}

/** The workbook's <externalReferences>, made in its place when there is none. */
function refsElement(wb) {
  const root = wb.doc.documentElement;
  let el = kid(root, 'externalReferences', wb.ns);
  if (el) return el;
  el = wb.doc.createElementNS(wb.ns, 'externalReferences');
  const before = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_REFS.includes(n.localName));
  root.insertBefore(el, before || null);
  return el;
}

/**
 * A new link to the workbook `target` (its file name, as the relationship gives it: Book.xlsx, a file in the same
 * folder), with its sheets' names: a part as Excel writes one (minus the absolute path Excel adds beside the relative
 * one), its relationship, its content type and its place in <externalReferences>. Returns its number ([n]); a link to
 * that file already there is reused (its sheets added to).
 */
export function addLink(wb, { target, sheetNames = [], names = [], unread = false }) {
  const have = linkOf(wb, target);
  if (have) {
    const missing = sheetNames.filter((s) => !have.sheetNames.some((n) => n.toLowerCase() === s.toLowerCase()));
    if (missing.length) updateLink(wb, have.index, { sheetNames: [...have.sheetNames, ...missing] });
    return have.index;
  }
  let n = 1;
  while (wb.pkg.has(`xl/externalLinks/externalLink${n}.xml`)) n++;
  const path = `xl/externalLinks/externalLink${n}.xml`;
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const sheets = sheetNames.map((s) => `<sheetName val="${esc(s)}"/>`).join('');
  const defs = names.length ? `<definedNames>${names.map((d) => `<definedName name="${esc(d.name)}" refersTo="${esc(d.refersTo)}"${d.sheetId != null ? ` sheetId="${d.sheetId}"` : ''}/>`).join('')}</definedNames>` : '';
  wb.pkg.setText(path, `${DECL}<externalLink xmlns="${NS.x}"><externalBook xmlns:r="${REL}" r:id="rId1">${sheets ? `<sheetNames>${sheets}</sheetNames>` : ''}${defs}<sheetDataSet>${sheetNames.map((_, k) => `<sheetData sheetId="${k}"${unread ? ' refreshError="1"' : ''}/>`).join('')}</sheetDataSet></externalBook></externalLink>`);
  wb.pkg.setText(`xl/externalLinks/_rels/externalLink${n}.xml.rels`, `${DECL}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${PATH_REL}" Target="${esc(target)}" TargetMode="External"/></Relationships>`);
  wb.pkg.addOverride(path, CT_LINK);
  const id = wb.pkg.addRel(LINK_REL, `externalLinks/externalLink${n}.xml`);
  const ref = wb.doc.createElementNS(wb.ns, 'externalReference');
  ref.setAttributeNS(REL, 'r:id', id);
  refsElement(wb).appendChild(ref);
  wb.bookDirty = true;
  return readLinks(wb).length;
}

/**
 * Change link `index`'s part: `sheetNames` (the sheets it has, in order), `names` (its defined names), `cells`
 * (Map sheet name -> Map 'r,c' -> value: number, text, boolean, an error ({ code } or an XlError), or null for blank,
 * merged into what is kept; `replace` to keep only these), `target` (its file). Writes the part only if it changed.
 */
export function updateLink(wb, index, { sheetNames = null, names = null, cells = null, target = null, replace = false } = {}) {
  const link = readLinks(wb)[index - 1];
  if (!link?.path) return false;
  const before = wb.pkg.files[link.path];
  const doc = parseXml(wb.pkg.text(link.path));
  const ns = doc.documentElement.namespaceURI;
  const book = kid(doc.documentElement, 'externalBook', ns);
  const make = (name) => doc.createElementNS(ns, name);
  const place = (el, order) => {
    const at = [...book.childNodes].find((n) => n.nodeType === 1 && order.includes(n.localName));
    book.insertBefore(el, at || null);
  };
  const oldNames = link.sheetNames;
  const nextNames = sheetNames || oldNames;
  if (sheetNames) {
    let list = kid(book, 'sheetNames', ns);
    if (!list) { list = make('sheetNames'); place(list, ['definedNames', 'sheetDataSet', 'extLst']); }
    while (list.firstChild) list.removeChild(list.firstChild);
    for (const s of sheetNames) { const el = make('sheetName'); el.setAttribute('val', s); list.appendChild(el); }
  }
  if (names) {
    let list = kid(book, 'definedNames', ns);
    if (!list && names.length) { list = make('definedNames'); place(list, ['sheetDataSet', 'extLst']); }
    if (list) {
      while (list.firstChild) list.removeChild(list.firstChild);
      for (const d of names) {
        const el = make('definedName');
        el.setAttribute('name', d.name);
        el.setAttribute('refersTo', d.refersTo);
        if (d.sheetId != null) el.setAttribute('sheetId', String(d.sheetId));
        list.appendChild(el);
      }
      if (!names.length) book.removeChild(list);
    }
  }
  if (sheetNames || cells) {
    // (The values by sheet: what is kept, moved to the sheets' new places, with `cells` over it.)
    const byName = new Map();
    oldNames.forEach((s, k) => byName.set(s.toLowerCase(), replace ? new Map() : new Map([...(link.cells[k] || new Map())].map(([key, x]) => [key, cachedValue(x)]))));
    for (const [s, m] of cells || []) {
      const key = s.toLowerCase();
      if (!byName.has(key)) byName.set(key, new Map());
      for (const [k, v] of m) byName.get(key).set(k, v);
    }
    let set = kid(book, 'sheetDataSet', ns);
    if (!set) { set = make('sheetDataSet'); place(set, ['extLst']); }
    while (set.firstChild) set.removeChild(set.firstChild);
    // (A sheet that couldn't be read stays so until values for it come: refreshError.)
    const given = new Set([...(cells || new Map()).keys()].map((x) => x.toLowerCase()));
    const unread = new Set(replace ? [] : [...link.errors].map((k) => oldNames[k]?.toLowerCase()).filter((x) => x && !given.has(x)));
    nextNames.forEach((s, k) => {
      const sd = make('sheetData');
      sd.setAttribute('sheetId', String(k));
      if (unread.has(s.toLowerCase())) sd.setAttribute('refreshError', '1');
      const m = byName.get(s.toLowerCase()) || new Map();
      const keys = [...m.keys()].filter((key) => m.get(key) != null).map((key) => key.split(',').map(Number)).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      let row = null;
      for (const [r, c] of keys) {
        if (!row || Number(row.getAttribute('r')) !== r + 1) { row = make('row'); row.setAttribute('r', String(r + 1)); sd.appendChild(row); }
        const v = m.get(`${r},${c}`);
        const cell = make('cell');
        cell.setAttribute('r', refName(r, c));
        const err = isError(v) ? v.code ?? String(v) : v && typeof v === 'object' && v.code ? v.code : null;
        if (err) cell.setAttribute('t', 'e');
        else if (typeof v === 'boolean') cell.setAttribute('t', 'b');
        else if (typeof v === 'string') cell.setAttribute('t', 'str');
        const ve = make('v');
        ve.textContent = err || (typeof v === 'boolean' ? (v ? '1' : '0') : String(v));
        cell.appendChild(ve);
        row.appendChild(cell);
      }
      set.appendChild(sd);
    });
  }
  if (target != null) {
    const relsPath = link.path.replace(/([^/]+)$/, '_rels/$1.rels');
    const rels = parseXml(wb.pkg.text(relsPath));
    const book2 = rid(book);
    for (const r of kids(rels.documentElement, 'Relationship', NS.rel)) {
      // (Excel's second, absolute path, xxl21:alternateUrls, goes with the old file.)
      if (r.getAttribute('Id') === book2) r.setAttribute('Target', target);
    }
    const alt = [...book.childNodes].find((n) => n.nodeType === 1 && n.localName === 'alternateUrls');
    if (alt) book.removeChild(alt);
    wb.pkg.setText(relsPath, serializeXml(rels));
  }
  const text = `${DECL}${serializeXml(doc.documentElement)}`;
  if (text !== wb.pkg.text(link.path)) wb.pkg.setText(link.path, text);
  return wb.pkg.files[link.path] !== before || target != null;
}

/** A workbook's own defined names as a link to it keeps them: the workbook's (not a sheet's, nor Excel's own), each
 *  sheet in their formulas quoted, as Excel writes them (measured: SrcRate, =Sheet1!$B$1, as ='Sheet1'!$B$1). */
function bookNamesOf(src) {
  const quoted = (text) => {
    try {
      return mapRefs(text, (t) => (t.type === 'ref' && t.sheet != null && !t.sheets && t.book == null ? `'${t.sheet.replace(/'/g, "''")}'!${t.text.slice(t.prefix.length)}` : null));
    } catch {
      return text;
    }
  };
  return src.definedNames.filter((d) => d.sheet == null && !/^_xlnm\./i.test(d.name)).map((d) => ({ name: d.name, refersTo: `=${quoted(d.text)}` }));
}

/** A value of a workbook's cell as a link keeps it: its value, or its formula's result (null when blank). */
function sourceValue(src, si, r, c) {
  const cell = src.sheet(si).cells.get(r)?.get(c);
  if (!cell || cell.v == null || cell.v === '') return null;
  if (cell.t === 'e') return { code: String(cell.v) };
  return cell.v;
}

/** Link `index` updated from `src`, the workbook it names (open in another editor, or read from its file), as
 *  Excel's Update Values does: its sheets' names, its workbook-wide names, and the value of every cell this workbook
 *  reads there. Returns whether anything changed (the formulas are then worked out again). */
export function refreshLink(wb, index, src) {
  const link = readLinks(wb)[index - 1];
  if (!link) return false;
  const sheetNames = src.sheets.map((s) => s.name);
  const limit = (name) => {
    const k = src.sheetIndexOf(name);
    if (k < 0 || src.sheets[k].kind !== 'worksheet') return { maxRow: 0, maxCol: 0 };
    const m = src.sheet(k);
    return { maxRow: m.maxRow, maxCol: m.maxCol };
  };
  const wanted = referencedCells(wb, index, limit);
  // (The cells already kept are read again too.)
  link.cells.forEach((m, k) => {
    const s = link.sheetNames[k];
    if (s == null) return;
    const key = s.toLowerCase();
    if (!wanted.has(key)) wanted.set(key, { sheet: s, cells: new Set() });
    for (const cell of m.keys()) wanted.get(key).cells.add(cell);
  });
  const cells = new Map();
  for (const { sheet, cells: want } of wanted.values()) {
    const k = src.sheetIndexOf(sheet);
    if (k < 0 || src.sheets[k].kind !== 'worksheet') continue;
    const m = new Map();
    for (const key of want) { const [r, c] = key.split(',').map(Number); m.set(key, sourceValue(src, k, r, c)); }
    cells.set(sheet, m);
  }
  const names = bookNamesOf(src);
  const changed = updateLink(wb, index, { sheetNames, names, cells, replace: true });
  if (changed) { wb.resetCalc(); wb.calculate(); }
  return changed;
}

/** A copy of cells of another workbook (clipops.js snapshot) put into the link to it, made if there is none: its
 *  sheet and the copied cells' values. Returns the link's number. */
export function linkCopy(wb, clip, target) {
  const src = clip.wb;
  const index = addLink(wb, { target, sheetNames: src.sheets.map((s) => s.name), names: bookNamesOf(src) });
  const m = new Map();
  for (let r = clip.rect.r0; r <= clip.rect.r1; r++) for (let c = clip.rect.c0; c <= clip.rect.c1; c++) m.set(`${r},${c}`, sourceValue(src, clip.si, r, c));
  updateLink(wb, index, { cells: new Map([[clip.sheetName, m]]) });
  return index;
}

/**
 * Edit Links' Break Link, as Excel's BreakLink does it (measured): every cell whose formula refers to link `index`
 * takes the value it shows instead; formulas reading it through a name, and the names themselves, stay as they were,
 * and so do the link and the other links' numbers. Returns the cells changed: [{ si, contents }] for the editor.
 */
export function breakLinkSteps(wb, index) {
  const steps = [];
  wb.sheets.forEach((info, k) => {
    if (info.kind !== 'worksheet' || !info.path || !wb.pkg.has(info.path)) return;
    const contents = [];
    for (const [r, line] of wb.sheet(k).cells) {
      for (const [c, cell] of line) {
        if (!cell.f || !cell.f.includes('[')) continue;
        let uses = false;
        // (A reference to its cells, or to one of its names, [1]!Rate.)
        try { uses = tokenize(cell.f).some((t) => (t.type === 'ref' || t.type === 'name') && t.book === String(index)); } catch { /* not read: kept */ }
        if (uses) contents.push({ r, c, value: cell.t === 'e' ? { code: String(cell.v) } : cell.v ?? null });
      }
    }
    if (contents.length) steps.push({ si: k, contents, styles: [], merges: wb.sheet(k).merges });
  });
  return steps;
}

/** Whether anything of the workbook still uses link `index` (a formula or a defined name). */
export function linkUsed(wb, index) {
  return referencedCells(wb, index).size > 0 || wb.definedNames.some((d) => d.text.includes(`[${index}]`));
}

/** The values a formula's references to other workbooks read, kept in their links when that workbook is open here
 *  (`find`, by file name: an editor's on the page), as Excel reads them when the formula is entered. */
export function cacheRefs(wb, formula, find = wb.openBook) {
  if (!formula || !formula.includes('[') || !find) return;
  let toks;
  try { toks = tokenize(formula); } catch { return; }
  for (const t of toks) {
    if (t.type !== 'ref' || t.book == null || t.sheet == null || !t.area) continue;
    const link = linkOf(wb, t.book);
    const src = link && find(link.name);
    const k = src ? src.sheetIndexOf(t.sheet) : -1;
    if (k < 0 || src.sheets[k].kind !== 'worksheet') continue;
    const m = new Map();
    const a = t.area;
    const r1 = Math.min(a.r1, src.sheet(k).maxRow);
    const c1 = Math.min(a.c1, src.sheet(k).maxCol);
    for (let r = a.r0; r <= r1 && r - a.r0 < 100000; r++) for (let c = a.c0; c <= c1; c++) m.set(`${r},${c}`, sourceValue(src, k, r, c));
    updateLink(wb, link.index, { cells: new Map([[t.sheet, m]]) });
  }
}
