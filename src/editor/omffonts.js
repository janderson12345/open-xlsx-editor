// open-metric-fonts' metric twins of 151 Microsoft families (omf.js: Morrow Sans Narrow for Aptos Narrow, Tamarack for
// Tahoma, Glade for Verdana, Compass for Century Gothic, Colines for Garamond...), loaded from a folder the page serves
// itself: `fonts: { omf }` given to mountXlsxEditor, the open-metric-fonts web bundle's folder (each package as
// <package>/files/), as the Word editor and a site serving one copy for every viewer use it.
// A family's faces are added to document.fonts the first time a workbook names it (fonts.js fontStack), as
// "oxe-omf <family>", per alphabet with its unicode-range, after the font itself in the stack: the browser downloads a
// file only when a cell shows text in that font, the computer lacks it, and the file covers the text's characters.
// When one arrives, index.js has the cells placed and measured again (fonts.js fontsChanged). FontFace objects, not a
// stylesheet: nothing inline, so a policy of font-src 'self' is all it needs.
import { OMF, OMF_RANGES } from './omf.js';

const BY_NAME = new Map(Object.keys(OMF).map((k) => [k.toLowerCase(), k]));
let root = null;
const added = new Set();

/** Where the twins load from: `base` (a URL, absolute or relative to the page, of a folder holding each package as
 *  <package>/files/; none: no twins). Returns it as an absolute URL ending in "/" ('' when there's none or no
 *  browser fonts). */
export function useOmf(base) {
  root = null;
  if (typeof document === 'undefined' || !document.fonts || typeof FontFace === 'undefined' || !base) return '';
  root = new URL(String(base).replace(/\/?$/, '/'), document.baseURI).href;
  return root;
}

/** The twin's entry in omf.js for a typeface name (any case): [package, metrics, faces], or null. */
export const omfEntry = (name) => OMF[BY_NAME.get(String(name).toLowerCase())] || null;

/** A file's characters (its unicode-range). */
export const omfRange = (pkg, alphabet) => OMF_RANGES[`${pkg}:${alphabet}`] || OMF_RANGES[alphabet];

/** The family the twin of `name` is drawn under, its faces added to document.fonts the first time; null when it has
 *  none or no folder was given. */
export function omfFamily(name) {
  const ms = BY_NAME.get(String(name).toLowerCase());
  if (!root || !ms) return null;
  const family = `oxe-omf ${ms}`;
  if (added.has(family)) return family;
  added.add(family);
  const [pkg, , faces] = OMF[ms];
  for (const [weight, italic, stem, alphabets] of faces) {
    for (const a of alphabets.split(',')) {
      const url = `${root}${pkg}/files/${stem}-${a}.woff2`;
      document.fonts.add(new FontFace(family, `url("${url}") format("woff2")`, { weight: String(weight), style: italic ? 'italic' : 'normal', unicodeRange: omfRange(pkg, a), display: 'swap' }));
    }
  }
  return family;
}

/** The twin's face the browser draws for bold and italic, as CSS font matching picks it (the other style where the
 *  family has none: Tahoma has no italic, the browser slants its regular): [package, file stem, alphabets], or null. */
export function omfFace(name, bold, italic) {
  const entry = omfEntry(name);
  if (!entry) return null;
  const [pkg, , faces] = entry;
  const score = ([w, i]) => (+!!i === +!!italic ? 0 : 2) + ((w >= 600) === !!bold ? 0 : 1);
  const [, , stem, alphabets] = [...faces].sort((a, b) => score(a) - score(b))[0];
  return [pkg, stem, alphabets.split(',')];
}
