// Data > Queries & Connections: the workbook's connections (xl/connections.xml) and its Power Query queries (the M
// of the DataMashup part, customXml/itemN.xml: base64 of a binary holding a package whose Formulas/Section1.m has
// them, as [MS-QDEFF] lays it out), each with where it loads (a table, through a queryTable that names its
// connection; the data model; or a connection only). A query's M can be changed: Section1.m is written again in the
// package, everything else of the DataMashup byte for byte. Refresh writes a host's rows into a query's table as
// Excel's Refresh does (excel-queries-refresh.ps1, test/queries-refresh-excel.json): the table's rows inserted or
// deleted as cells of its columns (what is under them moves), its columns kept by name, new ones made and dropped
// ones gone from the table and queryTable parts, its ExternalData name following. docs/QUERIES.md.
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { parseXml, serializeXml, kids, kid, attr, NS } from './xml.js';
import { rangeName, colName } from './refs.js';
import { readTables } from './tables.js';
import { findName, editName } from './names.js';

/** Excel for the web's words when a connection can't be refreshed (here: no host to get the data from). */
export const QUERY_MESSAGES = {
  cant: (names) => `Unable to refresh data for a data connection in the workbook. The following connections failed to refresh:\n\n${names.join('\n')}`,
  noTable: (name) => `The query ${name} isn't loaded to a table on a worksheet, so there's nothing here to refresh.`,
  rows: 'The data returned for the query is not a table: give rows (arrays, the first the column names) or objects.',
  syntax: (why) => `The query's formula can't be written: ${why}.`,
};

const CONNECTIONS_REL = `${NS.r}/connections`;
const QUERY_TABLE_REL = `${NS.r}/queryTable`;
const MASHUP_NS = 'http://schemas.microsoft.com/DataMashup';

// ---- The DataMashup part ----

/** Bytes as text: UTF-16 (little or big endian) when the BOM says so, else UTF-8 (a BOM dropped). */
function decodeText(u8) {
  if (u8[0] === 0xff && u8[1] === 0xfe) return { enc: 'utf-16le', bom: true, text: new TextDecoder('utf-16le').decode(u8.subarray(2)) };
  if (u8[0] === 0xfe && u8[1] === 0xff) return { enc: 'utf-16be', bom: true, text: new TextDecoder('utf-16be').decode(u8.subarray(2)) };
  const bom = u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf;
  return { enc: 'utf-8', bom, text: new TextDecoder('utf-8').decode(bom ? u8.subarray(3) : u8) };
}
function encodeText(s, enc, bom) {
  if (enc === 'utf-8') { const b = strToU8(s); if (!bom) return b; const out = new Uint8Array(b.length + 3); out.set([0xef, 0xbb, 0xbf]); out.set(b, 3); return out; }
  const out = new Uint8Array(2 + s.length * 2);
  const le = enc === 'utf-16le';
  out[0] = le ? 0xff : 0xfe; out[1] = le ? 0xfe : 0xff;
  for (let i = 0; i < s.length; i++) { const u = s.charCodeAt(i); out[2 + 2 * i] = le ? u & 255 : u >> 8; out[3 + 2 * i] = le ? u >> 8 : u & 255; }
  return out;
}
function fromBase64(s) {
  const bin = atob(s.replace(/\s+/g, ''));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function toBase64(u8) {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
  return btoa(s);
}
const u32 = (u8, o) => (u8[o] | (u8[o + 1] << 8) | (u8[o + 2] << 16) | (u8[o + 3] << 24)) >>> 0;

/** The DataMashup binary's pieces ([MS-QDEFF] 2.2): version, then the package parts, permissions, metadata and
 *  permission bindings, each after its 4-byte length. Null when it isn't laid out so. */
function splitMashup(bin) {
  if (bin.length < 4) return null;
  let o = 4;
  const pieces = [];
  for (let k = 0; k < 4; k++) {
    if (o + 4 > bin.length) return null;
    const n = u32(bin, o);
    o += 4;
    if (o + n > bin.length) return null;
    pieces.push(bin.subarray(o, o + n));
    o += n;
  }
  return { version: u32(bin, 0), pkg: pieces[0], permissions: pieces[1], metadata: pieces[2], bindings: pieces[3], rest: bin.subarray(o) };
}
function joinMashup(m) {
  const parts = [m.pkg, m.permissions, m.metadata, m.bindings];
  const out = new Uint8Array(4 + parts.reduce((s, p) => s + 4 + p.length, 0) + m.rest.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, m.version, true);
  let o = 4;
  for (const p of parts) { dv.setUint32(o, p.length, true); out.set(p, o + 4); o += 4 + p.length; }
  out.set(m.rest, o);
  return out;
}

/** Each query's metadata ([MS-QDEFF] 2.5.2: Section1/Name's StableEntries): { Name: { FillEnabled, ... } }. */
function readMetadata(md) {
  if (md.length < 8) return {};
  const n = u32(md, 4);
  const doc = (() => { try { return parseXml(decodeText(md.subarray(8, 8 + n)).text); } catch { return null; } })();
  const out = {};
  if (!doc) return out;
  for (const item of [...doc.getElementsByTagName('Item')]) {
    const type = item.getElementsByTagName('ItemType')[0]?.textContent;
    const path = item.getElementsByTagName('ItemPath')[0]?.textContent || '';
    const m = /^Section1\/([^/]+)$/.exec(path);
    if (type !== 'Formula' || !m) continue;
    const entries = {};
    for (const e of [...item.getElementsByTagName('Entry')]) {
      const v = e.getAttribute('Value') || '';
      // (Typed by their first letter: l a number, s text, d a date, b a boolean.)
      entries[e.getAttribute('Type')] = v[0] === 'l' ? Number(v.slice(1)) : v.slice(1);
    }
    out[decodeURIComponent(m[1])] = entries;
  }
  return out;
}

/** The workbook's DataMashup: { path, enc, bom, before, after, bin pieces, files (its package), section (the
 *  Section1.m text), queries: parseSection's, metadata }, or null when it has none (or none read). */
export function readMashup(wb) {
  for (const path of Object.keys(wb.pkg.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))) {
    const { enc, bom, text } = decodeText(wb.pkg.files[path]);
    const m = /(<DataMashup\b[^>]*>)([\s\S]*?)(<\/DataMashup>)/.exec(text);
    if (!m || !m[1].includes(MASHUP_NS)) continue;
    try {
      const pieces = splitMashup(fromBase64(m[2]));
      if (!pieces) continue;
      const files = unzipSync(pieces.pkg);
      const name = Object.keys(files).find((k) => /^Formulas\/Section1\.m$/i.test(k));
      const sec = name ? decodeText(files[name]) : { enc: 'utf-8', bom: false, text: '' };
      return { path, enc, bom, before: text.slice(0, m.index + m[1].length), after: text.slice(m.index + m[1].length + m[2].length), ...pieces, files, sectionPath: name, sectionBom: sec.bom,
        section: sec.text, queries: parseSection(sec.text), meta: readMetadata(pieces.metadata) };
    } catch { continue; }
  }
  return null;
}

/**
 * The members of an M section document ("section Section1; shared Sales = let ... in ...; ..."): [{ name, m, start,
 * end }], m the expression (Workbook.Queries(name).Formula's text) at [start, end) of the text. Strings ("" inside),
 * quoted names (#"..."), comments (// and /* *\/) are stepped over; a member ends at the ; outside brackets.
 */
export function parseSection(text) {
  const out = [];
  const n = text.length;
  let i = 0;
  let depth = 0;
  const skip = () => {
    // (Steps over a string, a quoted name or a comment at i; true when it did.)
    const ch = text[i];
    if (ch === '"' || (ch === '#' && text[i + 1] === '"')) {
      i += ch === '#' ? 2 : 1;
      while (i < n) { if (text[i] === '"') { if (text[i + 1] === '"') { i += 2; continue; } i++; break; } i++; }
      return true;
    }
    if (ch === '/' && text[i + 1] === '/') { while (i < n && text[i] !== '\n') i++; return true; }
    if (ch === '/' && text[i + 1] === '*') { const e = text.indexOf('*/', i + 2); i = e < 0 ? n : e + 2; return true; }
    return false;
  };
  const word = /[A-Za-z_][A-Za-z0-9_.]*/y;
  while (i < n) {
    if (skip()) continue;
    const ch = text[i];
    if ('([{'.includes(ch)) { depth++; i++; continue; }
    if (')]}'.includes(ch)) { depth--; i++; continue; }
    if (depth === 0 && ch === 's' && text.startsWith('shared', i) && !/[A-Za-z0-9_.]/.test(text[i - 1] || '') && /\s/.test(text[i + 6] || '')) {
      i += 6;
      while (/\s/.test(text[i] || '')) i++;
      let name;
      if (text[i] === '#' && text[i + 1] === '"') {
        const s0 = i;
        skip();
        name = text.slice(s0 + 2, i - 1).replace(/""/g, '"');
      } else {
        word.lastIndex = i;
        const w = word.exec(text);
        if (!w) continue;
        name = w[0];
        i += w[0].length;
      }
      while (/\s/.test(text[i] || '')) i++;
      if (text[i] !== '=') continue;
      i++;
      while (/\s/.test(text[i] || '')) i++;
      const start = i;
      let d = 0;
      while (i < n) {
        if (skip()) continue;
        const c = text[i];
        if ('([{'.includes(c)) d++;
        else if (')]}'.includes(c)) d--;
        else if (c === ';' && d <= 0) break;
        i++;
      }
      let end = i;
      while (end > start && /\s/.test(text[end - 1])) end--;
      out.push({ name, m: text.slice(start, end), start, end });
      i++;
      continue;
    }
    i++;
  }
  return out;
}

/** Why `m` can't stand as a member's expression (an open string, comment or bracket, or a ; outside brackets that
 *  would end it early), or null. */
export function formulaProblem(m) {
  if (!String(m).trim()) return 'it is empty';
  const probe = parseSection(`section Section1;\n\nshared Q = ${m};`);
  if (probe.length !== 1 || probe[0].m !== String(m).trim()) return 'a ; ends it before its end';
  let depth = 0;
  let i = 0;
  const s = String(m);
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"' || (ch === '#' && s[i + 1] === '"')) {
      i += ch === '#' ? 2 : 1;
      let closed = false;
      while (i < s.length) { if (s[i] === '"') { if (s[i + 1] === '"') { i += 2; continue; } i++; closed = true; break; } i++; }
      if (!closed) return 'a string isn\'t closed';
      continue;
    }
    if (ch === '/' && s[i + 1] === '/') { while (i < s.length && s[i] !== '\n') i++; continue; }
    if (ch === '/' && s[i + 1] === '*') { const e = s.indexOf('*/', i + 2); if (e < 0) return 'a comment isn\'t closed'; i = e + 2; continue; }
    if ('([{'.includes(ch)) depth++;
    if (')]}'.includes(ch) && --depth < 0) return 'a bracket closes that wasn\'t opened';
    i++;
  }
  return depth ? 'a bracket isn\'t closed' : null;
}

/** The query `name`'s M changed to `m` (Advanced Editor's Done): Section1.m written again in the DataMashup's package,
 *  its other parts, permissions, metadata and bindings as they were. Returns { error } when it can't. */
export function setQueryFormula(wb, name, m) {
  const mash = readMashup(wb);
  const q = mash?.queries.find((x) => x.name === name);
  if (!q) return { error: `There's no query named ${name}` };
  const why = formulaProblem(m);
  if (why) return { error: QUERY_MESSAGES.syntax(why) };
  const text = mash.section.slice(0, q.start) + String(m).trim() + mash.section.slice(q.end);
  if (text === mash.section) return null;
  const files = { ...mash.files, [mash.sectionPath]: encodeText(text, 'utf-8', mash.sectionBom) };
  // (Deflated as System.IO.Packaging writes it; the entries in their order.)
  const pkg = zipSync(files, { level: 6, mtime: new Date(1980, 0, 1) });
  const bin = joinMashup({ ...mash, pkg });
  wb.pkg.files[mash.path] = encodeText(mash.before + toBase64(bin) + mash.after, mash.enc, mash.bom);
  wb.edited = true;
  return null;
}

// ---- Connections ----

/** The connections part's path, or null. */
function connectionsPath(wb) {
  return wb.pkg.readRels(wb.pkg.mainPath).find((r) => r.type === CONNECTIONS_REL)?.target || null;
}

/** The connection string's Location (the query a Mashup connection reads), or null. */
function locationOf(conn) {
  if (!/Microsoft\.Mashup\.OleDb/i.test(conn || '')) return null;
  const m = /(?:^|;)\s*Location\s*=\s*("(?:[^"]|"")*"|[^;]*)/i.exec(conn);
  if (!m) return null;
  const v = m[1].trim();
  return v.startsWith('"') ? v.slice(1, -1).replace(/""/g, '"') : v;
}

const CONNECTION_TYPES = { 1: 'ODBC', 2: 'DAO', 3: 'File', 4: 'Web', 5: 'OLE DB', 6: 'Text', 7: 'ADO', 8: 'DSP', 100: 'Data model' };

/** Every connection of xl/connections.xml: [{ id, name, description, type, typeName, connection, command, query
 *  (the Power Query it reads, or null), model (true for one loaded to the data model, or the model's own),
 *  refreshedVersion }]. */
export function readConnections(wb) {
  const path = connectionsPath(wb);
  const doc = path && wb.pkg.xml(path);
  if (!doc) return [];
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  return kids(root, 'connection', ns).map((el) => {
    const db = kid(el, 'dbPr', ns);
    const type = Number(attr(el, 'type') || 0);
    const connection = attr(db, 'connection') || '';
    const name = attr(el, 'name') || '';
    // (Loaded to the data model: type 100 with an x15:connection, or the model's own ThisWorkbookDataModel.)
    const x15 = [...el.getElementsByTagNameNS('*', 'connection')].find((x) => x !== el);
    const model = type === 100 || x15?.getAttribute('model') === '1';
    let query = locationOf(connection);
    if (!query && type === 100) query = /^Query - (.+)$/.exec(name)?.[1] || null;
    return { id: Number(attr(el, 'id') || 0), name, description: attr(el, 'description') || '', type, typeName: CONNECTION_TYPES[type] || String(type), connection, command: attr(db, 'command') || '',
      query, model, own: name === 'ThisWorkbookDataModel', refreshedVersion: attr(el, 'refreshedVersion') };
  });
}

/** The query tables of the workbook's tables (tableType queryTable): [{ table (readTables'), path (the queryTable
 *  part), name (ExternalData_N), connectionId }]. */
function queryTables(wb) {
  const out = [];
  for (const t of readTables(wb)) {
    const rel = wb.pkg.readRels(t.path).find((r) => r.type === QUERY_TABLE_REL);
    const doc = rel && wb.pkg.xml(rel.target);
    if (!doc) continue;
    const root = doc.documentElement;
    out.push({ table: t, path: rel.target, name: attr(root, 'name'), connectionId: Number(attr(root, 'connectionId') || 0) });
  }
  return out;
}

/**
 * The Queries & Connections pane's lists: { queries: [{ name, m, load: 'table' | 'model' | 'connection', connection
 * (readConnections', or null), table ({ name, si, r0, c0, r1, c1, ref, sheet } or null), rows (a table's data rows; else FillCount, the rows last
 * loaded, or null),
 * error }], connections: readConnections' with each one's table }.
 */
export function readQueries(wb) {
  const mash = readMashup(wb);
  const conns = readConnections(wb);
  const qts = queryTables(wb);
  const tableOf = (c) => {
    const qt = c && qts.find((x) => x.connectionId === c.id);
    return qt ? { name: qt.table.name, si: qt.table.si, r0: qt.table.r0, c0: qt.table.c0, r1: qt.table.r1, c1: qt.table.c1, ref: rangeName(qt.table), sheet: wb.sheets[qt.table.si]?.name, queryTable: qt.path, part: qt.table.path } : null;
  };
  const queries = (mash?.queries || []).map((q) => {
    const meta = mash.meta[q.name] || {};
    const mine = conns.filter((c) => c.query === q.name);
    const withTable = mine.find((c) => tableOf(c));
    const connection = withTable || mine[0] || null;
    const table = tableOf(withTable);
    const load = table ? 'table' : mine.some((c) => c.model) || meta.FillToDataModelEnabled === 1 ? 'model' : 'connection';
    // (Rows loaded: a table's own data rows, as they are now; else the metadata's FillCount, as Excel last loaded it.)
    const qt = withTable && qts.find((x) => x.connectionId === withTable.id)?.table;
    const rows = qt ? qt.r1 - qt.r0 - qt.totals : typeof meta.FillCount === 'number' ? meta.FillCount : null;
    return { name: q.name, m: q.m, load, connection, table, rows,
      error: meta.FillErrorCode && meta.FillErrorCode !== 'Unknown' && meta.FillErrorCode !== 'None' ? meta.FillErrorCode : null };
  });
  return { queries, connections: conns.map((c) => ({ ...c, table: tableOf(c) })), mashup: !!mash };
}

/** What the pane says under a query, as Excel's does: "4 rows loaded.", "Connection only." (and "Loaded to the data
 *  model." for one there). */
export function queryStatus(q) {
  if (q.error) return `${q.error}.`;
  if (q.load === 'connection') return 'Connection only.';
  if (q.load === 'model' && q.rows == null) return 'Loaded to the data model.';
  if (q.rows == null) return q.load === 'model' ? 'Loaded to the data model.' : 'Loaded to a table.';
  return `${q.rows.toLocaleString('en-US')} row${q.rows === 1 ? '' : 's'} loaded.`;
}

// ---- Refresh ----

/** A host's answer as { columns, rows }: rows of arrays (the first the column names), or of objects (their keys the
 *  columns, in the first one's order then any new ones'); { columns, rows } as it is. Null when it is neither. */
export function tableOfRows(data) {
  if (data && !Array.isArray(data) && Array.isArray(data.columns) && Array.isArray(data.rows)) return { columns: data.columns.map(String), rows: data.rows.map((r) => (Array.isArray(r) ? r : data.columns.map((c) => r?.[c]))) };
  if (!Array.isArray(data) || !data.length) return null;
  if (data.every(Array.isArray)) return { columns: data[0].map((v) => String(v ?? '')), rows: data.slice(1) };
  if (data.every((r) => r && typeof r === 'object')) {
    const columns = [];
    for (const r of data) for (const k of Object.keys(r)) if (!columns.includes(k)) columns.push(k);
    return { columns, rows: data.map((r) => columns.map((k) => r[k])) };
  }
  return null;
}

/** A value as a cell's: numbers, text, true/false; a Date as its serial number (1900 date system); null and
 *  undefined an empty cell; anything else its text. */
function cellValue(v, date1904) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'boolean' || typeof v === 'string') return v;
  if (v instanceof Date) return (v.getTime() - Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30)) / 86400000;
  return String(v);
}

const guid = () => {
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
  const h = [...b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
};

/**
 * Refresh the query `q` (readQueries') loaded to a table with `data` (tableOfRows'), as Excel's Refresh writes a
 * query's result (excel-queries-refresh.ps1): the table's data rows made as many as the result's (one empty row for
 * none), inserted or deleted as cells of its columns so what is below moves; its query columns the result's, inserted
 * or deleted as cells of its rows before the user's own columns (unbound: dataBound="0", kept at their side, a
 * calculated one's formula in its new rows), so what is right of it moves; a column kept by name keeps its
 * tableColumn and queryTableField (formats and ids), a new one takes the lowest free id and the queryTable's nextId,
 * one gone is dropped from both parts; the header row and the values written; the queryTable's ExternalData name over
 * the query's columns. Returns { rows, area } or { error }.
 */
export function refreshQueryTable(wb, q, data) {
  const res = tableOfRows(data);
  if (!res) return { error: QUERY_MESSAGES.rows };
  if (!q.table) return { error: QUERY_MESSAGES.noTable(q.name) };
  const t0 = readTables(wb).find((t) => t.path === q.table.part);
  if (!t0) return { error: QUERY_MESSAGES.noTable(q.name) };
  const si = t0.si;
  // (The parts' columns as they were, before the lines move.)
  const tdoc = wb.pkg.xml(t0.path);
  const qdoc = wb.pkg.xml(q.table.queryTable);
  const tns = tdoc.documentElement.namespaceURI;
  const qns = qdoc.documentElement.namespaceURI;
  const oldCols = kids(kid(tdoc.documentElement, 'tableColumns', tns), 'tableColumn', tns);
  const refreshEl = kid(qdoc.documentElement, 'queryTableRefresh', qns);
  const fieldsEl = kid(refreshEl, 'queryTableFields', qns);
  const oldFields = kids(fieldsEl, 'queryTableField', qns);
  const left = Math.min(oldCols.length, Number(attr(refreshEl, 'unboundColumnsLeft') || 0));
  const right = Math.min(oldCols.length - left, Number(attr(refreshEl, 'unboundColumnsRight') || 0));
  const oldBound = oldCols.length - left - right;
  const bound = Math.max(1, res.columns.length);
  const height = Math.max(1, res.rows.length);
  const oldH = t0.r1 - t0.r0 - t0.totals;
  // Columns first (cells of the table's rows shifting right or left, at the end of its query columns), then rows
  // (cells of all its columns down or up).
  const at = t0.c0 + left + oldBound;
  if (bound > oldBound) wb.changeLines(si, 'col', at, bound - oldBound, [t0.r0, t0.r1]);
  else if (bound < oldBound) wb.changeLines(si, 'col', t0.c0 + left + bound, bound - oldBound, [t0.r0, t0.r1]);
  const width = left + bound + right;
  const c1 = t0.c0 + width - 1;
  if (height > oldH) wb.changeLines(si, 'row', t0.r0 + oldH, height - oldH, [t0.c0, c1]);
  else if (height < oldH) wb.changeLines(si, 'row', t0.r0 + 1 + height, height - oldH, [t0.c0, c1]);
  const area = { r0: t0.r0, c0: t0.c0, r1: t0.r0 + height, c1 };
  const q0 = t0.c0 + left;
  // The table part: its range, and its columns: the user's on their sides, the query's the result's.
  const t1 = readTables(wb).find((t) => t.path === t0.path);
  const doc = wb.pkg.xml(t1.path);
  const root = doc.documentElement;
  root.setAttribute('ref', rangeName({ ...area, r1: area.r1 + t1.totals }));
  const filter = kid(root, 'autoFilter', tns);
  if (filter) filter.setAttribute('ref', rangeName(area));
  const wrap = kid(root, 'tableColumns', tns);
  for (const el of kids(wrap, 'tableColumn', tns)) wrap.removeChild(el);
  const fieldOf = (el) => oldFields.find((f) => attr(f, 'tableColumnId') === attr(el, 'id'));
  const boundCols = oldCols.slice(left, left + oldBound);
  const used = new Set();
  const picked = res.columns.map((name) => {
    const el = boundCols.find((c) => !used.has(c) && String(attr(c, 'name')).toLowerCase() === name.toLowerCase());
    if (el) used.add(el);
    return el || null;
  });
  const ids = new Set([...picked.filter(Boolean), ...oldCols.slice(0, left), ...oldCols.slice(left + oldBound)].map((el) => Number(attr(el, 'id'))));
  let nextField = Number(attr(refreshEl, 'nextId') || 1);
  const xr3 = root.lookupNamespaceURI('xr3');
  const columnEls = [];
  const fieldEls = [];
  const keep = (el) => { columnEls.push(doc.importNode(el, true)); const f = fieldOf(el); if (f) fieldEls.push(f); };
  oldCols.slice(0, left).forEach(keep);
  res.columns.forEach((name, k) => {
    let el = picked[k];
    let field = el && fieldOf(el);
    if (el) {
      el = doc.importNode(el, true);
      el.setAttribute('name', name);
    } else {
      let id = 1;
      while (ids.has(id)) id++;
      ids.add(id);
      el = doc.createElementNS(tns, 'tableColumn');
      el.setAttribute('id', String(id));
      if (xr3) el.setAttributeNS(xr3, 'xr3:uid', guid());
      el.setAttribute('uniqueName', String(id));
      el.setAttribute('name', name);
      el.setAttribute('queryTableFieldId', String(nextField));
      field = qdoc.createElementNS(qns, 'queryTableField');
      field.setAttribute('id', String(nextField++));
      field.setAttribute('name', name);
      field.setAttribute('tableColumnId', String(id));
    }
    columnEls.push(el);
    if (field) { field.setAttribute('name', name); fieldEls.push(field); }
  });
  const rightCols = oldCols.slice(left + oldBound);
  rightCols.forEach(keep);
  for (const el of columnEls) wrap.appendChild(el);
  wrap.setAttribute('count', String(columnEls.length));
  wb.pkg.setText(t1.path, serializeXml(doc));
  wb.dirty.add(t1.path);
  // The queryTable part: its fields in the columns' order, their count and next id.
  for (const f of oldFields) fieldsEl.removeChild(f);
  for (const f of fieldEls) fieldsEl.appendChild(f);
  fieldsEl.setAttribute('count', String(fieldEls.length));
  refreshEl.setAttribute('nextId', String(nextField));
  wb.pkg.setText(q.table.queryTable, serializeXml(qdoc));
  wb.dirty.add(q.table.queryTable);
  // (ExternalData_N, hidden and the sheet's own, over the query's columns.)
  const qname = attr(qdoc.documentElement, 'name');
  const dn = qname && findName(wb, qname, si);
  if (dn) {
    const abs = `$${colName(q0)}$${area.r0 + 1}:$${colName(q0 + bound - 1)}$${area.r1 + 1}`;
    const text = dn.text.includes('!') ? dn.text.replace(/![^!]*$/, `!${abs}`) : abs;
    if (text !== dn.text) editName(wb, { name: qname, sheet: si }, { refersTo: `=${text}` });
  }
  // The header row and the values; a calculated column of the user's, its formula in every data row.
  const date1904 = !!wb.date1904;
  const cells = [];
  res.columns.forEach((name, k) => cells.push({ r: area.r0, c: q0 + k, value: name }));
  for (let r = 0; r < height; r++) for (let k = 0; k < bound; k++) cells.push({ r: area.r0 + 1 + r, c: q0 + k, value: cellValue(res.rows[r]?.[k], date1904) });
  const sideCols = [...oldCols.slice(0, left).map((el, k) => [el, t0.c0 + k]), ...rightCols.map((el, k) => [el, q0 + bound + k])];
  for (const [el, c] of sideCols) {
    const f = kid(el, 'calculatedColumnFormula', tns)?.textContent;
    if (!f) continue;
    for (let r = area.r0 + 1; r <= area.r1; r++) if (wb.sheet(si).cells.get(r)?.get(c)?.f == null) cells.push({ r, c, formula: f });
  }
  wb.apply(si, cells);
  wb.edited = true;
  wb.resetCalc();
  return { rows: res.rows.length, area };
}
