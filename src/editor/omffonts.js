// open-metric-fonts' Morrow Sans Narrow and Morrow Sans, the metric twins of Aptos Narrow and Aptos (fonts.js), loaded
// from a folder the page serves itself: `fonts: { omf }` given to mountXlsxEditor, the open-metric-fonts web bundle's
// folder (each package as <package>/files/), as the Word editor and a site serving one copy for every viewer use it.
// Each face is added to document.fonts per alphabet with its unicode-range, so the browser downloads a file only when
// a cell shows text it covers; when one arrives, index.js has the cells placed and measured again (fonts.js
// fontsChanged). FontFace objects, not a stylesheet: nothing inline, so a policy of font-src 'self' is all it needs.
import { OMF, OMF_RANGES } from './omf.js';

const added = new Set();
const TWINS = { 'Aptos Narrow': 'Morrow Sans Narrow', Aptos: 'Morrow Sans' };

/** Adds the faces under `base` (a URL ending in "/", absolute or relative to the page) to document.fonts, once per
 *  base and file. Returns how many were added. */
export function addOmfFaces(base) {
  if (typeof document === 'undefined' || !document.fonts || typeof FontFace === 'undefined' || !base) return 0;
  const root = new URL(String(base).replace(/\/?$/, '/'), document.baseURI).href;
  let n = 0;
  for (const [name, family] of Object.entries(TWINS)) {
    const [pkg, , faces] = OMF[name];
    for (const [weight, italic, stem, alphabets] of faces) {
      for (const a of alphabets.split(',')) {
        const url = `${root}${pkg}/files/${stem}-${a}.woff2`;
        if (added.has(url)) continue;
        added.add(url);
        const range = OMF_RANGES[`${pkg}:${a}`] || OMF_RANGES[a];
        document.fonts.add(new FontFace(family, `url("${url}") format("woff2")`, { weight: String(weight), style: italic ? 'italic' : 'normal', unicodeRange: range, display: 'swap' }));
        n++;
      }
    }
  }
  return n;
}
