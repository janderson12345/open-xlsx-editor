// Pictures in headers and footers (Page Setup > Header/Footer > Insert Picture, the &G code): a VML drawing the sheet
// names in legacyDrawingHF, one shape a picture, its id the section it is in, as Excel writes them (its files, made over
// COM): LH, CH, RH for the header's left, centre and right, LF, CF, RF for the footer's, FIRST or EVEN after the id
// for the first page's and the even pages' (LHFIRST, RFEVEN); the shape's style gives the picture's size
// (`width:60pt;height:30pt`), its v:imagedata o:relid the image, through the drawing's own relationships.

const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
const UNIT_PT = { pt: 1, px: 0.75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12 };

/** A CSS length of a VML style in points (a bare number is pixels). */
function lengthPt(text) {
  const m = /^\s*(-?[\d.]+)\s*(pt|px|in|cm|mm|pc)?\s*$/.exec(text || '');
  return m ? Number(m[1]) * UNIT_PT[m[2] || 'px'] : null;
}

/** Bytes as base64, in a browser or in Node. */
function base64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** Sheet si's header and footer pictures: { [shape id]: { src (a data URL), w, h (points) } }, {} when it has none. */
export function headerPictures(wb, si) {
  const pkg = wb.pkg;
  const path = wb.sheets[si]?.path;
  const xml = path && pkg.text(path);
  const id = xml && /<(?:\w+:)?legacyDrawingHF\b[^>]*\br:id="([^"]+)"/.exec(xml)?.[1];
  if (!id) return {};
  const rel = pkg.readRels(path).find((r) => r.id === id);
  const vml = rel && !rel.external && pkg.text(rel.target);
  if (!vml) return {};
  const images = pkg.readRels(rel.target);
  const out = {};
  for (const m of vml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)) {
    const attr = (name, text) => new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`).exec(text)?.slice(1).find((v) => v != null);
    const sid = attr('id', m[1]);
    const relid = attr('o:relid', m[2]) ?? attr('r:id', m[2]);
    const style = Object.fromEntries((attr('style', m[1]) || '').split(';').map((p) => p.split(':').map((s) => s.trim())));
    const img = images.find((r) => r.id === relid);
    const bytes = img && !img.external && pkg.files[img.target];
    const w = lengthPt(style.width);
    const h = lengthPt(style.height);
    if (!sid || !bytes || !w || !h) continue;
    const mime = MIME[img.target.split('.').pop().toLowerCase()] || 'application/octet-stream';
    out[sid] = { src: `data:${mime};base64,${base64(bytes)}`, w, h };
  }
  return out;
}
