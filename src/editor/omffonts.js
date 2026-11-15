// open-metric-fonts' Morrow Sans Narrow and Morrow Sans, the metric twins of Aptos Narrow and Aptos (fonts.js), loaded
// from a folder the page serves itself: `fonts: { omf }` given to mountXlsxEditor, the open-metric-fonts web bundle's
// folder (each package as <package>/files/), as the Word editor and a site serving one copy for every viewer use it.
// Each face is added to document.fonts per alphabet with its unicode-range, so the browser downloads a file only when
// a cell shows text it covers; when one arrives, index.js has the cells placed and measured again (fonts.js
// fontsChanged). FontFace objects, not a stylesheet: nothing inline, so a policy of font-src 'self' is all it needs.
import { OMF_FACES } from './omf.js';

const added = new Set();

/** Adds the faces under `base` (a URL ending in "/", absolute or relative to the page) to document.fonts, once per
 *  base and file. Returns how many were added. */
export function addOmfFaces(base) {
  if (typeof document === 'undefined' || !document.fonts || typeof FontFace === 'undefined' || !base) return 0;
  const root = new URL(String(base).replace(/\/?$/, '/'), document.baseURI).href;
  let n = 0;
  for (const [family, weight, style, file, range] of OMF_FACES) {
    const url = `${root}morrow-sans/files/${file}`;
    if (added.has(url)) continue;
    added.add(url);
    document.fonts.add(new FontFace(family, `url("${url}") format("woff2")`, { weight: String(weight), style, unicodeRange: range, display: 'swap' }));
    n++;
  }
  return n;
}
