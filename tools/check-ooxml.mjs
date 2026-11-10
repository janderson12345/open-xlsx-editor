// Structural check of .xlsx files, the things that make Excel offer to repair a file: every XML part well-formed,
// every part with a content type, every internal relationship's target present, every r:id the workbook, sheets and
// drawings use defined in their relationships; in the workbook, sheet names valid and unique; in each worksheet, rows
// and cells in order, shared-string and style indexes in range, merged ranges not overlapping.
// (Not a full XSD validation: that needs the ECMA-376 schemas. Anything new in the format is opened in Excel too.)
// usage: node tools/check-ooxml.mjs <file.xlsx | dir>...
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

const files = process.argv.slice(2).filter((a) => fs.existsSync(a)).flatMap((a) => (fs.statSync(a).isDirectory() ? fs.readdirSync(a).filter((f) => /\.xls[xm]$/i.test(f)).map((f) => path.join(a, f)) : [a]));
if (!files.length) { console.error('usage: node tools/check-ooxml.mjs <file.xlsx | dir>...'); process.exit(2); }

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = /\/officeDocument$/;
const parse = (text, where, problems) => {
  try {
    return new DOMParser({ onError: (level, msg) => { if (level !== 'warning') throw new Error(msg); } }).parseFromString(text, 'text/xml');
  } catch (e) {
    problems.push(`${where}: not well-formed: ${e.message.split('\n')[0]}`);
    return null;
  }
};
const resolve = (base, target) => path.posix.normalize(path.posix.join(base ? path.posix.dirname(base) : '', target)).replace(/^\//, '');
/** Elements by local name, whatever the prefix. */
const byName = (node, name) => [...node.getElementsByTagName('*')].filter((e) => e.localName === name);
const kidsNamed = (el, name) => (el ? [...el.childNodes].filter((n) => n.nodeType === 1 && n.localName === name) : []);
const colIndex = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const cellRef = (s) => { const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(s || ''); return m ? { r: Number(m[2]) - 1, c: colIndex(m[1]) } : null; };
const rangeRef = (s) => { const [a, b = a] = String(s || '').split(':'); const p = cellRef(a); const q = cellRef(b); return p && q ? { r0: Math.min(p.r, q.r), c0: Math.min(p.c, q.c), r1: Math.max(p.r, q.r), c1: Math.max(p.c, q.c) } : null; };

/** Rows and cells in increasing order, values and indexes in range, merges apart. */
function checkWorksheet(n, d, sstCount, xfCount, problems) {
  const sheetData = byName(d, 'sheetData')[0];
  if (!sheetData) { problems.push(`${n}: no sheetData`); return; }
  let r = -1;
  for (const row of kidsNamed(sheetData, 'row')) {
    const given = row.getAttribute('r');
    const next = given ? Number(given) - 1 : r + 1;
    if (!(next > r)) problems.push(`${n}: row ${next + 1} after row ${r + 1}`);
    if (next > 1048575) problems.push(`${n}: row ${next + 1} past the last row`);
    r = next;
    let c = -1;
    for (const cell of kidsNamed(row, 'c')) {
      const at = cellRef(cell.getAttribute('r'));
      if (cell.getAttribute('r') && !at) { problems.push(`${n}: cell reference "${cell.getAttribute('r')}"`); continue; }
      if (at && at.r !== r) problems.push(`${n}: ${cell.getAttribute('r')} inside row ${r + 1}`);
      const col = at ? at.c : c + 1;
      if (!(col > c)) problems.push(`${n}: row ${r + 1}: column ${col + 1} after column ${c + 1}`);
      if (col > 16383) problems.push(`${n}: row ${r + 1}: column ${col + 1} past XFD`);
      c = col;
      const where = cell.getAttribute('r') || `row ${r + 1} column ${col + 1}`;
      const s = cell.getAttribute('s');
      if (s != null && xfCount != null && !(Number(s) >= 0 && Number(s) < xfCount)) problems.push(`${n}: ${where}: style s="${s}", cellXfs has ${xfCount}`);
      const v = kidsNamed(cell, 'v')[0];
      const t = cell.getAttribute('t') || 'n';
      if (t === 's' && v) {
        const i = Number(v.textContent);
        if (!(Number.isInteger(i) && i >= 0 && i < (sstCount ?? 0))) problems.push(`${n}: ${where}: shared string ${v.textContent}, the table has ${sstCount ?? 'none'}`);
      }
      if (t === 'n' && v && v.textContent !== '' && !Number.isFinite(Number(v.textContent))) problems.push(`${n}: ${where}: "${v.textContent}" is not a number`);
      if (t === 'b' && v && !/^(0|1|true|false)$/.test(v.textContent)) problems.push(`${n}: ${where}: "${v.textContent}" is not a boolean`);
      if (t === 'inlineStr' && !kidsNamed(cell, 'is').length) problems.push(`${n}: ${where}: inline string without <is>`);
      if (!/^(n|s|b|e|str|inlineStr|d)$/.test(t)) problems.push(`${n}: ${where}: cell type "${t}"`);
    }
  }
  const merges = byName(d, 'mergeCell').map((m) => [m.getAttribute('ref'), rangeRef(m.getAttribute('ref'))]);
  for (let i = 0; i < merges.length; i++) {
    if (!merges[i][1]) { problems.push(`${n}: merged range "${merges[i][0]}"`); continue; }
    for (let j = 0; j < i; j++) {
      const a = merges[i][1];
      const b = merges[j][1];
      if (b && a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1) problems.push(`${n}: merged ranges ${merges[j][0]} and ${merges[i][0]} overlap`);
    }
  }
}

/** Sheet names Excel accepts: 1 to 31 characters, none of : \ / ? * [ ], not starting or ending with ', unique. */
function checkSheetNames(wb, problems) {
  const seen = new Set();
  const ids = new Set();
  for (const s of byName(wb, 'sheet')) {
    const name = s.getAttribute('name') || '';
    if (!name.length || name.length > 31) problems.push(`xl/workbook.xml: sheet name "${name}" is ${name.length} characters`);
    if (/[:\\/?*[\]]/.test(name) || /^'|'$/.test(name)) problems.push(`xl/workbook.xml: sheet name "${name}" has a character Excel refuses`);
    if (seen.has(name.toLowerCase())) problems.push(`xl/workbook.xml: two sheets named "${name}"`);
    seen.add(name.toLowerCase());
    const id = s.getAttribute('sheetId');
    if (ids.has(id)) problems.push(`xl/workbook.xml: sheetId ${id} used twice`);
    ids.add(id);
  }
}

let bad = 0;
let skipped = 0;
for (const file of files) {
  const problems = [];
  const bytes = new Uint8Array(fs.readFileSync(file));
  // (A workbook saved with a password is a compound file holding the encrypted package: nothing to check without the
  // password. test/encryption.test.js checks those the editor writes.)
  if (bytes.length >= 8 && [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((b, i) => bytes[i] === b)) { skipped++; console.log(`skip ${file} (saved with a password)`); continue; }
  let zip;
  try { zip = unzipSync(bytes); } catch (e) { zip = null; problems.push(`not a zip: ${e.message}`); }
  const names = zip ? Object.keys(zip) : [];
  const xml = new Map();
  for (const n of names) if (/\.(xml|rels|vml)$/i.test(n)) { const d = parse(strFromU8(zip[n]), n, problems); if (d) xml.set(n, d); }
  const types = xml.get('[Content_Types].xml');
  if (zip && !types) problems.push('[Content_Types].xml missing');
  else if (types) {
    const defaults = new Set([...types.getElementsByTagName('Default')].map((e) => e.getAttribute('Extension').toLowerCase()));
    const overrides = new Set([...types.getElementsByTagName('Override')].map((e) => e.getAttribute('PartName').replace(/^\//, '')));
    for (const n of names) {
      if (n.endsWith('/') || n === '[Content_Types].xml') continue;
      if (!overrides.has(n) && !defaults.has(n.split('.').pop().toLowerCase())) problems.push(`${n}: no content type`);
    }
    for (const o of overrides) if (!names.includes(o)) problems.push(`[Content_Types].xml: override for missing part ${o}`);
  }
  // Relationships: ids unique, internal targets present.
  const relsOf = new Map(); // part -> Map(id -> { type, target })
  for (const [n, d] of xml) {
    if (!n.endsWith('.rels')) continue;
    const at = n.lastIndexOf('_rels/');
    const part = n.slice(0, at) + n.slice(at + 6).replace(/\.rels$/, ''); // ('' for the package's own)
    const ids = new Map();
    for (const r of d.getElementsByTagName('Relationship')) {
      const id = r.getAttribute('Id');
      if (ids.has(id)) problems.push(`${n}: relationship ${id} defined twice`);
      const t = r.getAttribute('Target');
      const external = r.getAttribute('TargetMode') === 'External';
      const full = external ? t : t.startsWith('/') ? t.slice(1) : resolve(part, t);
      ids.set(id, { type: r.getAttribute('Type'), target: full });
      if (!external && !names.includes(full)) problems.push(`${n}: ${id} points to missing ${full}`);
    }
    relsOf.set(part, ids);
  }
  const main = [...(relsOf.get('') || new Map()).values()].find((r) => OFFICE_DOC.test(r.type))?.target;
  if (zip && !main) problems.push('no officeDocument relationship (the workbook part)');
  else if (zip && !xml.has(main)) problems.push(`${main} missing`);
  // The workbook, sheets, drawings and the like: every r:id they use is defined in their relationships.
  for (const [n, d] of xml) {
    if (n.endsWith('.rels') || n === '[Content_Types].xml' || !n.startsWith('xl/')) continue;
    const ids = relsOf.get(n) || new Map();
    const all = d.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
      for (const a of [...all[i].attributes]) {
        if (a.namespaceURI === R && !ids.has(a.value)) problems.push(`${n}: ${all[i].tagName} uses ${a.name}="${a.value}", not in its relationships`);
      }
    }
  }
  if (main && xml.has(main)) {
    const wb = xml.get(main);
    checkSheetNames(wb, problems);
    const rels = relsOf.get(main) || new Map();
    const find = (kind) => [...rels.values()].find((r) => r.type.endsWith('/' + kind))?.target;
    const sst = find('sharedStrings') && xml.get(find('sharedStrings'));
    const sstCount = sst ? byName(sst, 'si').length : null;
    const styles = find('styles') && xml.get(find('styles'));
    const xfCount = styles ? kidsNamed(byName(styles, 'cellXfs')[0], 'xf').length : null;
    for (const s of byName(wb, 'sheet')) {
      const id = [...s.attributes].find((a) => a.localName === 'id' && a.namespaceURI)?.value;
      const rel = rels.get(id);
      if (!rel) continue; // (reported above as an undefined r:id)
      if (!/\/(worksheet|chartsheet|dialogsheet|xlMacrosheet|xlIntlMacrosheet)$/.test(rel.type)) problems.push(`${main}: sheet "${s.getAttribute('name')}" points to a ${rel.type.split('/').pop()} part`);
      if (rel.type.endsWith('/worksheet') && xml.has(rel.target)) checkWorksheet(rel.target, xml.get(rel.target), sstCount, xfCount, problems);
    }
  }
  if (problems.length) { bad++; console.log(`FAIL ${file}`); for (const p of problems.slice(0, 20)) console.log(`  ${p}`); if (problems.length > 20) console.log(`  ...and ${problems.length - 20} more`); }
  else console.log(`ok   ${file}`);
}
const passed = skipped ? ` (${skipped} saved with a password skipped)` : '';
console.log(bad ? `${bad} of ${files.length - skipped} files with problems${passed}` : `all ${files.length - skipped} files ok${passed}`);
process.exit(bad ? 1 : 0);
