// What Home > Clear takes from a sheet besides its cells, as Excel does it (measured over COM, Range.ClearComments,
// ClearHyperlinks, Hyperlinks.Delete, Clear and ClearContents, and the files Excel then saved: docs/FORMATS.md):
// - a hyperlink that meets the cleared cells goes whole (D1:E2's, clearing D1), with its relationship;
// - Clear Hyperlinks leaves the cells' formats (the Hyperlink style stays); Remove Hyperlinks takes the formats of the
//   hyperlinks' cells away too, theirs only (a bold cell beside them stays bold);
// - Clear Comments and Notes takes the notes and the threaded comments (with their replies) of the cells;
// - Clear All does all of it, Clear Contents the hyperlinks, Clear Formats neither.
import { kids, serializeXml, NS } from './xml.js';
import { parseRange } from './refs.js';
import { readNotes } from './notes.js';
import { deleteNote } from './noteedit.js';
import { clearChanges } from './clearops.js';

const meets = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** Sheet si's hyperlinks that meet `area`: [{ el, range }]. */
export function hyperlinksIn(wb, si, area) {
  const { doc } = wb._part(si);
  const list = doc.getElementsByTagNameNS(doc.documentElement.namespaceURI, 'hyperlinks')[0];
  if (!list) return [];
  return kids(list, 'hyperlink', doc.documentElement.namespaceURI)
    .map((el) => ({ el, range: parseRange(el.getAttribute('ref') || '') }))
    .filter((h) => h.range && meets(h.range, area));
}

/** The hyperlinks of sheet si meeting `area` taken away, each whole, with its relationship (an address outside the
 *  workbook). Returns their ranges. */
export function clearHyperlinks(wb, si, area) {
  const found = hyperlinksIn(wb, si, area);
  if (!found.length) return [];
  const { doc } = wb._part(si);
  const path = wb.sheets[si].path;
  const ids = [];
  for (const { el } of found) {
    const id = el.getAttributeNS(NS.r, 'id');
    if (id) ids.push(id);
    const list = el.parentNode;
    list.removeChild(el);
    if (!kids(list, 'hyperlink', doc.documentElement.namespaceURI).length) list.parentNode.removeChild(list);
  }
  if (ids.length) {
    const dir = path.slice(0, path.lastIndexOf('/') + 1);
    const relsPath = `${dir}_rels/${path.slice(dir.length)}.rels`;
    const rels = wb.pkg.xml(relsPath);
    if (rels) {
      for (const r of kids(rels.documentElement, 'Relationship', NS.rel)) if (ids.includes(r.getAttribute('Id'))) rels.documentElement.removeChild(r);
      wb.pkg.setText(relsPath, serializeXml(rels));
    }
  }
  wb.dirty.add(path);
  wb.edited = true;
  return found.map((h) => h.range);
}

/** The notes and threaded comments of sheet si's cells in `area` taken away. Returns how many. */
export function clearNotes(wb, si, area) {
  let n = 0;
  for (const key of readNotes(wb, si).keys()) {
    const [r, c] = key.split(',').map(Number);
    if (r >= area.r0 && r <= area.r1 && c >= area.c0 && c <= area.c1 && deleteNote(wb, si, r, c)) n++;
  }
  return n;
}

/** Whether cells of `area` on sheet si have notes or threaded comments. */
export function notesIn(wb, si, area) {
  for (const key of readNotes(wb, si).keys()) {
    const [r, c] = key.split(',').map(Number);
    if (r >= area.r0 && r <= area.r1 && c >= area.c0 && c <= area.c1) return true;
  }
  return false;
}
/** Whether `area` of sheet si has hyperlinks, or with `notes` notes and comments too, for a Clear to take away. */
export const hasParts = (wb, si, area, { notes = false } = {}) => hyperlinksIn(wb, si, area).length > 0 || (notes && notesIn(wb, si, area));

/** The Clears that are of the sheet's parts alone: Clear Comments and Notes, Clear Hyperlinks, Remove Hyperlinks. */
export const PARTS_ONLY = ['comments', 'hyperlinks', 'removeHyperlinks'];

/** A Clear of `areas` on sheet si as work on the sheet's parts, `run(book, k)`, with the cells' step (clearops.js) for
 *  Clear All and Clear Contents done in it too; null when there is nothing of the parts to take (Clear All and Clear
 *  Contents then go as cells alone). */
export function partsClear(wb, si, areas, what, step = null) {
  if (what === 'formats') return null;
  const notes = what === 'all' || what === 'comments';
  const links = what !== 'comments';
  if (!areas.some((a) => (links && hyperlinksIn(wb, si, a).length) || (notes && notesIn(wb, si, a)))) return null;
  return (book, k) => {
    if (step?.contents.length) book.apply(k, step.contents);
    if (what === 'all') { book.setStyles(k, step.styles); book.setMerges(k, step.merges); }
    for (const a of areas) {
      if (notes) clearNotes(book, k, a);
      if (!links) continue;
      const gone = clearHyperlinks(book, k, a);
      // (Remove Hyperlinks: the formats of the hyperlinks' cells, the whole of each, taken away as Clear Formats does.)
      if (what === 'removeHyperlinks') for (const range of gone) { const f = clearChanges(book, k, range, 'formats'); if (!f.error) book.setStyles(k, f.styles); }
    }
  };
}
