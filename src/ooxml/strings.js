// Cell text: the shared string table (xl/sharedStrings.xml) and inline strings, which hold the same `si`-shaped content.
import { kids, kid, attr, onOff, decodeXstring } from './xml.js';
import { readColor } from './styles.js';

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

/** A run's font (`rPr`), as Excel writes it (measured: each property given, the font's name as rFont): { name, size,
 *  scheme, b, i, u, strike, color, vertAlign }, null where it doesn't say. */
function runFont(rPr, ns) {
  const val = (name) => attr(kid(rPr, name, ns), 'val');
  const u = kid(rPr, 'u', ns);
  const color = kid(rPr, 'color', ns);
  return {
    name: val('rFont') ?? null,
    size: val('sz') != null ? Number(val('sz')) : null,
    scheme: val('scheme') ?? null,
    b: !!onOff(kid(rPr, 'b', ns)),
    i: !!onOff(kid(rPr, 'i', ns)),
    u: u ? attr(u, 'val') || 'single' : null,
    strike: !!onOff(kid(rPr, 'strike', ns)),
    color: color ? readColor(color) : null,
    vertAlign: val('vertAlign') ?? null,
  };
}

/** The rich-text runs of an `si` (or `is`): [{ text, font }] (font null: the cell's own, as Excel takes a run without
 *  rPr), or null when it is plain text. */
export function stringItemRuns(si, ns) {
  const runs = si ? kids(si, 'r', ns) : [];
  if (!runs.length) return null;
  return runs.map((r) => {
    const rPr = kid(r, 'rPr', ns);
    return { text: decodeXstring(kid(r, 't', ns)?.textContent || ''), font: rPr ? runFont(rPr, ns) : null };
  });
}

/** Every string of a sharedStrings part, by index; `.runs`, a Map of index to the runs of those with rich text. */
export function readSharedStrings(doc) {
  if (!doc) return Object.assign([], { runs: new Map() });
  const root = doc.documentElement;
  const runs = new Map();
  const list = kids(root, 'si', root.namespaceURI).map((si, k) => {
    const r = stringItemRuns(si, root.namespaceURI);
    if (r) runs.set(k, r);
    return stringItemText(si, root.namespaceURI);
  });
  return Object.assign(list, { runs });
}
