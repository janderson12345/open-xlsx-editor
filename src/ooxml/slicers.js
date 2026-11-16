// Slicers and timelines when rows and columns move, as Excel does it (measured over COM:
// tools/excel-compare/excel-slicers.ps1 on slicer-cases.mjs, test/slicers-excel.json). Their drawing's anchors move
// as any drawing's (anchors.js moveDrawings: a slicer is put in "move but don't size" by Excel's ribbon, editAs
// oneCell); a table slicer's cache names its table by id and its column by the column's id, so columns inserted into
// the table change nothing in it. A table slicer whose column is deleted, or whose table is, is deleted by Excel: the
// slicer, its place in the drawing, its cache part, the cache's defined name (Slicer_Region) and the workbook's and the
// sheet's lists of them. PivotTable slicers and timelines stay, their fields deleted from the source or not (Excel
// keeps them until the pivot is refreshed).
import { serializeXml, parseXml } from './xml.js';
import { deleteName, findName } from './names.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SLICER_CACHE = /\/slicerCache$/;
const SLICER = /\/slicer$/;
const DRAWING = /\/drawing$/;

const elements = (el) => [...el.childNodes].filter((n) => n.nodeType === 1);
const remove = (el) => el?.parentNode && el.parentNode.removeChild(el);
const relId = (el) => el.getAttributeNS(REL, 'id') || el.getAttribute('r:id');
const attrOf = (attrs, name) => new RegExp(`\\s${name}="([^"]*)"`).exec(attrs)?.[1] ?? null;

/** A part's XML text from its DOM, with the declaration the part had. */
function textOf(doc, original) {
  const decl = /^(?:\uFEFF)?(<\?xml[^?]*\?>)/.exec(original || '');
  const body = serializeXml(doc).replace(/^<\?xml[^?]*\?>\s*/, '');
  return decl ? `${decl[1]}\r\n${body}` : body;
}

/** Elements named `name` whose relationship id is `id` taken out of `root`; the lists and extensions they leave empty
 *  too (a workbook's or sheet's extLst entry for a slicer or cache). True when any was. */
function dropListed(root, name, id) {
  let any = false;
  for (const el of [...root.getElementsByTagNameNS('*', name)]) {
    if (relId(el) !== id) continue;
    let up = el.parentNode;
    remove(el);
    any = true;
    while (up && up !== root && !elements(up).length) { const next = up.parentNode; remove(up); up = next; }
  }
  return any;
}

/** Remove relationship `id` from a part's relationships. */
function dropRel(pkg, partPath, id) {
  const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
  const relsPath = `${dir}_rels/${partPath.slice(dir.length)}.rels`;
  const text = pkg.text(relsPath);
  if (text) pkg.setText(relsPath, text.replace(new RegExp(`<Relationship\\b[^>]*\\sId="${id}"[^>]*/>`), ''));
}

/**
 * The table slicers whose column or table was deleted (`dead`, anchors.js applyTables's [{ id, columns, columnIds }]:
 * the table's id, and the ids of the columns deleted, or columns null when all of it was) deleted, as Excel does.
 */
export function dropTableSlicers(wb, dead) {
  if (!dead.length) return;
  const pkg = wb.pkg;
  const caches = [];
  for (const rel of pkg.readRels(pkg.mainPath).filter((r) => SLICER_CACHE.test(r.type) && !r.external)) {
    const text = pkg.text(rel.target) || '';
    const attrs = /<(?:\w+:)?tableSlicerCache\b([^>]*)>/.exec(text)?.[1];
    if (attrs == null) continue;
    const table = attrOf(attrs, 'tableId');
    const column = attrOf(attrs, 'column');
    if (!dead.some((d) => d.id === table && (d.columns == null || (d.columnIds || []).includes(column)))) continue;
    caches.push({ rel, name: /<slicerCacheDefinition\b[^>]*\sname="([^"]*)"/.exec(text)?.[1] });
  }
  for (const cache of caches) {
    const names = dropSlicers(wb, cache.name);
    for (const name of names) dropFromDrawings(wb, name);
    if (cache.name && findName(wb, cache.name, null)) deleteName(wb, { name: cache.name });
    if (dropListed(wb.doc.documentElement, 'slicerCache', cache.rel.id)) wb.bookDirty = true;
    pkg.removePart(cache.rel.target);
  }
}

/** The slicers showing cache `cacheName`, on every sheet, taken out of their slicers parts (a part left empty goes,
 *  with the sheet's relationship and list entry). Returns their names. */
function dropSlicers(wb, cacheName) {
  const pkg = wb.pkg;
  const names = [];
  wb.sheets.forEach((info, si) => {
    if (info.kind !== 'worksheet' || !info.path || !pkg.has(info.path)) return;
    for (const rel of pkg.readRels(info.path).filter((r) => SLICER.test(r.type) && !r.external)) {
      const text = pkg.text(rel.target);
      if (!text) continue;
      const out = text.replace(/<slicer\b([^>]*?)(?:\/>|>[\s\S]*?<\/slicer>)/g, (el, attrs) => {
        if (attrOf(attrs, 'cache') !== cacheName) return el;
        names.push(attrOf(attrs, 'name'));
        return '';
      });
      if (out === text) continue;
      if (/<slicer\b/.test(out)) { pkg.setText(rel.target, out); continue; }
      const { doc } = wb._part(si);
      dropListed(doc.documentElement, 'slicer', rel.id);
      dropRel(pkg, info.path, rel.id);
      pkg.removePart(rel.target);
      wb.dirty.add(info.path);
    }
  });
  return names;
}

/** A slicer's place in its sheet's drawing (the anchor, or the alternate content round it, that shows it) removed. */
function dropFromDrawings(wb, name) {
  const pkg = wb.pkg;
  for (const info of wb.sheets) {
    if (info.kind !== 'worksheet' || !info.path || !pkg.has(info.path)) continue;
    for (const rel of pkg.readRels(info.path).filter((r) => DRAWING.test(r.type) && !r.external)) {
      const text = pkg.text(rel.target);
      if (!text || !text.includes(`name="${name}"`)) continue;
      const doc = parseXml(text);
      const shows = (el) => [...el.getElementsByTagNameNS('*', 'slicer')].some((s) => s.getAttribute('name') === name);
      const gone = elements(doc.documentElement).filter(shows);
      if (!gone.length) continue;
      gone.forEach(remove);
      pkg.setText(rel.target, textOf(doc, text));
    }
  }
}
