// The Watch Window's cells as Excel keeps them in the file (measured: Application.Watches.Add, then the file saved and
// opened again): a sheet's <cellWatches> of <cellWatch r="B2"/>, in the order they were added, after pageMargins and
// the rest of the page set-up, before ignoredErrors. A range is watched a cell at a time.
import { kid, kids, attr } from './xml.js';
import { parseRef, refName } from './refs.js';

// (cellWatches comes after these parts of a worksheet, and before the rest.)
const AFTER = ['ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

const docOf = (wb, si) => wb._part(si)?.doc;

/** Every watched cell of the workbook, sheet by sheet in their order, each sheet's in the order added: [{ si, r, c }]. */
export function readWatches(wb) {
  const out = [];
  wb.sheets.forEach((info, si) => {
    if (info.kind !== 'worksheet' || !info.path) return;
    const doc = docOf(wb, si);
    const root = doc?.documentElement;
    if (!root) return;
    for (const el of kids(kid(root, 'cellWatches', root.namespaceURI), 'cellWatch', root.namespaceURI)) {
      const p = parseRef(attr(el, 'r') || '');
      if (p) out.push({ si, r: p.r, c: p.c });
    }
  });
  return out;
}

/** Add Watch: cells [{ r, c }] of sheet si watched, after those already (a cell watched already stays where it is).
 *  Returns how many were added. */
export function addWatches(wb, si, cells) {
  const doc = docOf(wb, si);
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const make = (name) => doc.createElementNS(ns, root.prefix ? `${root.prefix}:${name}` : name);
  let list = kid(root, 'cellWatches', ns);
  const have = new Set(kids(list, 'cellWatch', ns).map((el) => attr(el, 'r')));
  let added = 0;
  for (const { r, c } of cells) {
    const ref = refName(r, c);
    if (have.has(ref)) continue;
    if (!list) {
      list = make('cellWatches');
      const next = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName));
      root.insertBefore(list, next || null);
    }
    const el = make('cellWatch');
    el.setAttribute('r', ref);
    list.appendChild(el);
    have.add(ref);
    added++;
  }
  if (added) { wb.dirty.add(wb.sheets[si].path); wb.edited = true; }
  return added;
}

/** Delete Watch: watched cells [{ si, r, c }] no longer watched; a sheet left with none loses its cellWatches. */
export function deleteWatches(wb, cells) {
  for (const { si, r, c } of cells) {
    const root = docOf(wb, si)?.documentElement;
    const list = root && kid(root, 'cellWatches', root.namespaceURI);
    if (!list) continue;
    for (const el of kids(list, 'cellWatch', root.namespaceURI)) if (attr(el, 'r') === refName(r, c)) list.removeChild(el);
    if (!kids(list, 'cellWatch', root.namespaceURI).length) root.removeChild(list);
    wb.dirty.add(wb.sheets[si].path);
    wb.edited = true;
  }
}
