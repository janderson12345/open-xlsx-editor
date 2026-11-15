// For exportPdf (pdf.js): an open-metric-fonts twin the page serves (editor/omffonts.js, "oxe-omf <family>" in a
// cell's font stack), embedded from the same files the browser draws it from: the face's WOFF2 file of each alphabet
// a piece of text uses, fetched from the editor's `fonts.omf` folder (the browser has most in its cache already) and
// embedded as a subset. fontkit reads WOFF2 (measured on Tamarack, Tahoma's twin: its latin file subset to a 5.2 KB
// PDF, as its TrueType file to 5.5 KB; the face's whole-alphabet WOFF2 came out at 1.4 MB, so the per-alphabet files
// are used).
import { omfFace, omfRange } from './editor/omffonts.js';

const PREFIX = 'oxe-omf ';

/** A unicode-range as [first, last] code point pairs. */
const parseRange = (r) => r.split(',').map((p) => { const [a, b] = p.replace(/^U\+/i, '').split('-'); return [parseInt(a, 16), parseInt(b ?? a, 16)]; });

/** The twin for a font-family name from a stack ("oxe-omf Tahoma"), bold and italic as the browser picks its face:
 *  { files: [{ ranges, load }] }, each file's font loaded once when first needed, or null when it isn't a twin or the
 *  folder isn't known. Its latin file is loaded first: null too when that can't be had. */
export async function twinFor(pdf, family, bold, italic, root) {
  if (!root || !family.startsWith(PREFIX)) return null;
  const face = omfFace(family.slice(PREFIX.length), bold, italic);
  if (!face) return null;
  const [pkg, stem, alphabets] = face;
  const files = alphabets.map((a) => {
    let loading = null;
    const load = () => (loading ||= fetch(`${root}${pkg}/files/${stem}-${a}.woff2`)
      .then((r) => (r.ok ? r.arrayBuffer() : null))
      .then(async (b) => { if (!b) return null; const font = await pdf.embedFont(new Uint8Array(b), { subset: true }); return { font, chars: new Set(font.getCharacterSet()) }; })
      .catch(() => null));
    return { ranges: parseRange(omfRange(pkg, a)), load };
  });
  const latin = files[alphabets.indexOf('latin')];
  if (!latin || !(await latin.load())) return null;
  return { files, latin };
}

/** `text` in the twin's fonts: [{ font, text }], each character in the file whose range holds it (as the browser
 *  picks it by the faces' unicode-ranges), one it lacks drawn as "?" in the latin file's font. */
export async function twinSegments(twin, text) {
  const out = [];
  const base = await twin.latin.load();
  for (const ch of text) {
    const c = ch.codePointAt(0);
    let f = null;
    for (const file of twin.files) {
      if (!file.ranges.some(([a, b]) => c >= a && c <= b)) continue;
      const got = await file.load();
      if (got && got.chars.has(c)) { f = got; break; }
    }
    const [font, t] = f ? [f.font, ch] : [base.font, ch.trim() ? '?' : ' '];
    if (out.length && out[out.length - 1].font === font) out[out.length - 1].text += t;
    else out.push({ font, text: t });
  }
  return out;
}
