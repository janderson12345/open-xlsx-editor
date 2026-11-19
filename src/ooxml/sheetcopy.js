// Move or Copy > Create a copy, as Excel copies a sheet (measured over COM: excel-datatools.ps1's Worksheet.Copy): the
// copy is named "Data (2)" (the first number free), its part the source's text for text but for a new xr:uid and its
// tab not selected; its notes, VML, drawings and charts copied with it (the charts' references to the source sheet
// now to the copy); the source's local names copied as the copy's, and a workbook-wide name that refers to the source
// given a copy local to the copy, referring to it, as Excel's did. A sheet with tables, PivotTables, threaded comments
// or controls is refused for now: their parts need names and ids of their own.
import { addSheet, MESSAGES, withEl } from './sheetops.js';
import { sheetPrefix } from '../calc/tokens.js';
import { serializeXml } from './xml.js';

const COPIED = new Set(['comments', 'vmlDrawing', 'drawing', 'hyperlink', 'printerSettings', 'image']);
const REFUSED = /\/(table|pivotTable|threadedComment|ctrlProp|control|slicer|timeline|oleObject|package)$/;
const CT = { comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml', drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', printerSettings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings' };
const hex = (n) => Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16).toUpperCase()).join('');
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const freePath = (pkg, path) => { const m = /^(.*?)(\d*)(\.\w+)$/.exec(path); for (let n = 1; ; n++) { const p = `${m[1]}${n}${m[3]}`; if (!pkg.has(p)) return p; } };
const relsPath = (part) => { const d = part.lastIndexOf('/') + 1; return `${part.slice(0, d)}_rels/${part.slice(d)}.rels`; };
const relTarget = (from, to) => { const a = from.slice(0, from.lastIndexOf('/')).split('/'); const b = to.split('/'); let k = 0; while (k < a.length && a[k] === b[k]) k++; return [...a.slice(k).map(() => '..'), ...b.slice(k)].join('/'); };

/** The name a copy of sheet `name` gets: "name (2)", or the next number free, cut to Excel's 31 characters. */
export function copyName(wb, name) {
  const base = name.replace(/ \(\d+\)$/, '');
  for (let n = 2; ; n++) {
    const tail = ` (${n})`;
    const t = base.slice(0, 31 - tail.length) + tail;
    if (!wb.sheets.some((s) => s.name.toLowerCase() === t.toLowerCase())) return t;
  }
}

/** A part copied to a free path of its own, with the parts it points to that only it uses (a drawing's charts and
 *  pictures), `fix(text)` changing the copy's text. Returns the copy's path. */
function copyPart(pkg, path, fix = (t) => t) {
  const to = freePath(pkg, path);
  const bytes = pkg.files[path];
  if (/\.(xml|vml|rels)$/i.test(path)) pkg.setText(to, fix(pkg.text(path))); else pkg.files[to] = bytes;
  const ct = /\/charts\/chart\d*\.xml$/.test(path) ? CT.chart : /\/drawings\/drawing\d*\.xml$/.test(path) ? CT.drawing : /comments\d*\.xml$/.test(path) ? CT.comments : /printerSettings/.test(path) ? CT.printerSettings : null;
  if (ct) pkg.addOverride(to, ct);
  const rels = pkg.text(relsPath(path));
  if (rels) {
    let out = rels;
    for (const r of pkg.readRels(path)) {
      if (r.external) continue;
      // (A drawing's charts are copied, theirs to the copy; pictures shared.)
      if (/\/chart$/.test(r.type)) { const c = copyPart(pkg, r.target, fix); out = out.replace(new RegExp(`(Id="${esc(r.id)}"[^>]*Target=")[^"]*"`), `$1${relTarget(to, c)}"`); }
    }
    pkg.setText(relsPath(to), out);
  }
  return to;
}

/** Copy sheet i to a new sheet placed at `at` (just after it by default). Returns the copy's index. */
export function copySheet(wb, i, at = i + 1) {
  const pkg = wb.pkg;
  const src = wb.sheets[i];
  if (src.kind !== 'worksheet') throw new Error(MESSAGES.notWorksheet || "Only worksheets can be copied here yet.");
  const rels = pkg.readRels(src.path);
  if (rels.some((r) => REFUSED.test(r.type))) throw new Error("This sheet can't be copied here yet: it has tables, PivotTables, threaded comments or controls.");
  const name = copyName(wb, src.name);
  const k = addSheet(wb, at, name);
  const from = k <= i ? i + 1 : i;
  const source = wb.sheets[from];
  const copy = wb.sheets[k];
  // (References to the source sheet in the copy's charts: to the copy.)
  const oldRef = new RegExp(esc(sheetPrefix(source.name)), 'g');
  const fix = (t) => t.replace(oldRef, sheetPrefix(name));
  // (The source as it is now: its part may have changed since it was read.)
  const now = wb._parsed.get(source) && wb._parsed.get(source).hasDoc !== false ? serializeXml(wb._part(from).doc) : pkg.text(source.path);
  let xml = now.replace(/xr:uid="\{[^}]*\}"/, `xr:uid="{${hex(8)}-${hex(4)}-${hex(4)}-${hex(4)}-${hex(12)}}"`).replace(/ tabSelected="1"/, '');
  const newRels = [];
  for (const r of pkg.readRels(source.path)) {
    const kind = r.type.slice(r.type.lastIndexOf('/') + 1);
    if (!COPIED.has(kind)) continue;
    if (r.external) { newRels.push(r); continue; }
    const to = kind === 'drawing' || kind === 'vmlDrawing' || kind === 'comments' || kind === 'printerSettings' ? copyPart(pkg, r.target, fix) : r.target;
    // (A VML drawing's shapes take ids from a block of 1,024 of its own: the copy, the first block free.)
    if (kind === 'vmlDrawing') {
      const used = new Set(Object.keys(pkg.files).filter((p) => /\.vml$/i.test(p) && p !== to).flatMap((p) => [...(pkg.text(p) || '').matchAll(/<o:idmap\b[^>]*\sdata="([^"]+)"/g)].flatMap((m) => m[1].split(',').map(Number))));
      let block = 1;
      while (used.has(block)) block++;
      const text = pkg.text(to);
      const old = Number(/<o:idmap\b[^>]*\sdata="(\d+)/.exec(text)?.[1] || 1);
      pkg.setText(to, text.replace(/(<o:idmap\b[^>]*\sdata=")[^"]+"/, `$1${block}"`).replace(/_x0000_s(\d+)/g, (m, n) => `_x0000_s${Number(n) - old * 1024 + block * 1024}`));
    }
    newRels.push({ ...r, target: to });
  }
  if (newRels.length) {
    const body = newRels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.external ? r.target.replace(/&/g, '&amp;') : relTarget(copy.path, r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('');
    pkg.setText(relsPath(copy.path), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`);
  } else xml = xml.replace(/<(legacyDrawing|drawing)\b[^>]*\/>/g, '');
  pkg.setText(copy.path, xml);
  wb._parsed.delete(copy);
  wb._part(k); // (read as a sheet, to be saved as the others are)
  // Names: the source's local ones, and workbook-wide ones referring to it, as the copy's own.
  const refs = new RegExp(`(^|[^A-Za-z0-9_.'])${esc(sheetPrefix(source.name))}`);
  for (const d of [...wb.definedNames]) {
    const local = d.sheet === from;
    const global = d.sheet == null && refs.test(d.text) && !wb.definedNames.some((x) => x.name === d.name && x.sheet === k);
    if (!local && !global) continue;
    const el = wb.doc.createElementNS(wb.ns, 'definedName');
    for (const a of [...d.el.attributes]) el.setAttribute(a.name, a.value);
    el.setAttribute('localSheetId', String(k));
    const text = d.text.split(sheetPrefix(source.name)).join(sheetPrefix(name));
    el.textContent = text;
    d.el.parentNode.insertBefore(el, local ? d.el.nextSibling : d.el);
    wb.definedNames.push(withEl({ ...d, sheet: k, text }, el));
  }
  wb.bookDirty = true;
  wb.dirty.add(copy.path);
  wb.resetCalc();
  wb.edited = true;
  return k;
}
