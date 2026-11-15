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

const escText = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// (An attribute's value: its quotes escaped too, whatever a run's font was given through the API.)
const escAttr = (t) => escText(t).replace(/"/g, '&quot;');
/** A font colour (styles.js readColor) as its element. */
const colorXml = (c) => {
  if (!c) return '';
  if (c.auto) return '<color auto="1"/>';
  if (c.rgb) return `<color rgb="${escAttr(c.rgb)}"/>`;
  if (c.theme != null) return `<color theme="${escAttr(c.theme)}"${c.tint ? ` tint="${escAttr(c.tint)}"` : ''}/>`;
  if (c.indexed != null) return `<color indexed="${c.indexed}"/>`;
  return '';
};

/**
 * A cell's runs as the `r` elements of its string, as Excel writes rich text (test/richtext.test.js, measured over COM):
 * a run in the cell's own font without rPr; another with an rPr of what it changes (bold, italic, strikethrough,
 * underline, super- or subscript) and then the font it is in: size, colour, name (rFont), family and scheme as the
 * cell's (`base`: { name, size, color, family, scheme }), or a size, colour or font of its own. `runs`: [{ text, font:
 * { b, i, strike, u, vertAlign, size, color, name, family } or null }], the colour 'FFRRGGBB' or, as a run read from
 * a file has it (runFont), { rgb } or { theme, tint } or { indexed }; a font of its own keeps its family.
 */
export function richRunsXml(runs, base) {
  return runs.filter((r) => r.text).map((r) => {
    const f = r.font;
    const t = `<t${/^\s|\s$|[\t\n]/.test(r.text) ? ' xml:space="preserve"' : ''}>${escText(r.text)}</t>`;
    if (!f || !Object.values(f).some((v) => v != null && v !== false)) return `<r>${t}</r>`;
    const own = f.name && f.name !== base.name;
    const props = [f.b && '<b/>', f.i && '<i/>', f.strike && '<strike/>', f.u && (f.u === 'single' || f.u === true ? '<u/>' : `<u val="${escAttr(f.u)}"/>`), f.vertAlign && `<vertAlign val="${escAttr(f.vertAlign)}"/>`,
      `<sz val="${escAttr(f.size ?? base.size)}"/>`, (f.color && (typeof f.color === 'string' ? `<color rgb="${escAttr(f.color)}"/>` : colorXml(f.color))) || colorXml(base.color),
      `<rFont val="${escAttr(own ? f.name : base.name)}"/>`, own && f.family != null ? `<family val="${escAttr(f.family)}"/>` : '',
      !own && base.family != null ? `<family val="${escAttr(base.family)}"/>` : '', !own && base.scheme ? `<scheme val="${escAttr(base.scheme)}"/>` : ''].filter(Boolean).join('');
    return `<r><rPr>${props}</rPr>${t}</r>`;
  }).join('');
}

