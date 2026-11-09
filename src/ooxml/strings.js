// Cell text: the shared string table (xl/sharedStrings.xml) and inline strings, which hold the same `si`-shaped content.
import { kids, kid, decodeXstring } from './xml.js';

/** The text of an `si` (or a cell's `is`): its own `t`, or the `t` of each rich-text run `r`. Phonetic runs (`rPh`,
 *  the reading guide over Japanese text) are not part of the cell's text and are skipped. */
export function stringItemText(si, ns) {
  if (!si) return '';
  const t = kid(si, 't', ns);
  let s = t ? t.textContent : '';
  for (const r of kids(si, 'r', ns)) {
    const rt = kid(r, 't', ns);
    if (rt) s += rt.textContent;
  }
  return decodeXstring(s);
}

/** Every string of a sharedStrings part, by index. */
export function readSharedStrings(doc) {
  if (!doc) return [];
  const root = doc.documentElement;
  return kids(root, 'si', root.namespaceURI).map((si) => stringItemText(si, root.namespaceURI));
}
