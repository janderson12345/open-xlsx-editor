// A sheet's drawing objects as its drawing parts have them (docs/LAYOUT.md, Pictures, shapes and icons): every anchor
// (twoCellAnchor, oneCellAnchor, absoluteAnchor), in the drawing's order, which is the order they are drawn in, each
// with what it holds: a picture (xdr:pic: its image through the drawing's relationships, an SVG icon's own picture
// beside its PNG fallback, its crop (srcRect), rotation and flips), a shape or connector (xdr:sp, xdr:cxnSp: its
// preset or custom geometry, fill, line, arrowheads, text, the theme colours its style names), a group (xdr:grpSp, its
// children in its own coordinates) or a chart (drawn by chartview.js; listed here for its place among the others).
// Anything else (SmartArt, slicers, ink) is listed as 'frame' and not drawn. Anchors inside markup compatibility's
// Fallback are counted (for the anchors' order) and not drawn: their Choice is.
//
// The edits: an object moved or sized (its anchor's corners, the extent it keeps, its own offset and size) and objects
// deleted (their anchors, relationships only they used, and the parts only those led to). The rest of the part stays
// byte for byte.
import { kids, kid, attr, NS } from './xml.js';
import { emuToPx, EMU_PER_PX } from './units.js';

const A = NS.a;
const XDR = NS.xdr;
const R = NS.r;
const MC = NS.mc;
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const SVG_BLIP = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const REL_DRAWING = `${R}/drawing`;
const ANCHORS = new Set(['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']);
const CONTENT = new Set(['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame', 'contentPart']);
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf' };

/** Bytes as base64, in a browser or in Node. */
function base64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
// (A picture's data URL, made once for its bytes: a sheet is read again after every edit.)
const urls = new WeakMap();
function dataUrl(bytes, mime) {
  let u = urls.get(bytes);
  if (!u) { u = `data:${mime};base64,${base64(bytes)}`; urls.set(bytes, u); }
  return u;
}

/** All the anchors of a drawing part in document order, with the content each is drawn by. */
function anchorsOf(doc) {
  const out = [];
  const walk = (el, fallback) => {
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if (c.namespaceURI === XDR && ANCHORS.has(c.localName)) out.push({ el: c, fallback, top: c.parentNode === doc.documentElement });
      else walk(c, fallback || (c.namespaceURI === MC && c.localName === 'Fallback'));
    }
  };
  walk(doc.documentElement, false);
  return out;
}

/** An anchor's content element: its first sp, pic, grpSp, cxnSp or graphicFrame, or the one in its AlternateContent's
 *  first Choice. */
function contentOf(anchor) {
  for (let c = anchor.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1) continue;
    if (c.namespaceURI === XDR && CONTENT.has(c.localName)) return c;
    if (c.namespaceURI === MC && c.localName === 'AlternateContent') {
      const choice = kid(c, 'Choice', MC);
      for (let d = choice?.firstChild; d; d = d.nextSibling) if (d.nodeType === 1 && d.namespaceURI === XDR && CONTENT.has(d.localName)) return d;
    }
  }
  return null;
}

const num = (v, d = 0) => (v == null || v === '' || Number.isNaN(Number(v)) ? d : Number(v));
/** An anchor corner: { c, r, dx, dy } (dx and dy in pixels into the cell). */
const corner = (el) => el && ({ c: num(kid(el, 'col', XDR)?.textContent), r: num(kid(el, 'row', XDR)?.textContent), dx: emuToPx(num(kid(el, 'colOff', XDR)?.textContent)), dy: emuToPx(num(kid(el, 'rowOff', XDR)?.textContent)) });

/** A DrawingML colour element's colour: { rgb | scheme | ph, mods: [[name, value]] }, or null. */
export function colorOf(el) {
  if (!el) return null;
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1 || c.namespaceURI !== A) continue;
    let base = null;
    if (c.localName === 'srgbClr') base = { rgb: attr(c, 'val', null) };
    else if (c.localName === 'schemeClr') base = attr(c, 'val', null) === 'phClr' ? { ph: true } : { scheme: attr(c, 'val', null) };
    else if (c.localName === 'sysClr') base = { rgb: attr(c, 'lastClr', null) || (attr(c, 'val', null) === 'window' ? 'FFFFFF' : '000000') };
    else if (c.localName === 'prstClr') base = { rgb: PRESET_COLORS[attr(c, 'val', null)] || '000000' };
    else if (c.localName === 'scrgbClr') base = { rgb: ['r', 'g', 'b'].map((k) => Math.round(255 * Math.min(1, num(attr(c, k, null)) / 100000)).toString(16).padStart(2, '0')).join('') };
    else continue;
    base.mods = kids(c, null).filter((m) => m.namespaceURI === A).map((m) => [m.localName, num(attr(m, 'val', null)) / 100000]);
    return base;
  }
  return null;
}
const PRESET_COLORS = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', cyan: '00FFFF', magenta: 'FF00FF', lime: '00FF00', navy: '000080', silver: 'C0C0C0', maroon: '800000', olive: '808000', teal: '008080' };

/** A fill (spPr's, or a run's): { none } or { color } (a gradient as its middle stop's colour), or undefined when it
 *  sets none (the style's then). */
function fillOf(el) {
  if (!el) return undefined;
  if (kid(el, 'noFill', A)) return { none: true };
  const solid = kid(el, 'solidFill', A);
  if (solid) return { color: colorOf(solid) };
  const grad = kid(el, 'gradFill', A);
  if (grad) {
    const stops = kids(kid(grad, 'gsLst', A), 'gs', A);
    const mid = stops[Math.floor(stops.length / 2)] || stops[0];
    return mid ? { color: colorOf(mid), gradient: true } : { none: true };
  }
  if (kid(el, 'pattFill', A)) return { color: colorOf(kid(kid(el, 'pattFill', A), 'fgClr', A)) };
  if (kid(el, 'grpFill', A)) return { group: true };
  return undefined;
}

/** A line (a:ln): { none } or { color, w (px), dash, cap, head, tail }, undefined fields where it says nothing. */
function lineOf(ln) {
  if (!ln) return undefined;
  const fill = fillOf(ln);
  const end = (name) => { const e = kid(ln, name, A); return e && attr(e, 'type', null) && attr(e, 'type', null) !== 'none' ? { type: attr(e, 'type', null), w: attr(e, 'w', null) || 'med', len: attr(e, 'len', null) || 'med' } : null; };
  return { none: !!fill?.none, color: fill?.color, w: attr(ln, 'w', null) != null ? emuToPx(attr(ln, 'w', null)) : undefined,
    dash: attr(kid(ln, 'prstDash', A), 'val', null) || undefined, cap: attr(ln, 'cap', null) || undefined, head: end('headEnd'), tail: end('tailEnd') };
}

/** A transform (a:xfrm): { x, y, w, h } (px), rot (degrees), flipH, flipV, and a group's child offset and extent. */
function xfrmOf(el) {
  if (!el) return null;
  const off = kid(el, 'off', A);
  const ext = kid(el, 'ext', A);
  const chOff = kid(el, 'chOff', A);
  const chExt = kid(el, 'chExt', A);
  return {
    x: emuToPx(num(attr(off, 'x', null))), y: emuToPx(num(attr(off, 'y', null))), w: emuToPx(num(attr(ext, 'cx', null))), h: emuToPx(num(attr(ext, 'cy', null))),
    rot: num(attr(el, 'rot', null)) / 60000, flipH: attr(el, 'flipH', null) === '1' || attr(el, 'flipH', null) === 'true', flipV: attr(el, 'flipV', null) === '1' || attr(el, 'flipV', null) === 'true',
    ch: chExt ? { x: emuToPx(num(attr(chOff, 'x', null))), y: emuToPx(num(attr(chOff, 'y', null))), w: emuToPx(num(attr(chExt, 'cx', null))), h: emuToPx(num(attr(chExt, 'cy', null))) } : null,
  };
}

/** A shape's geometry: { prst, av } or { cust: paths } (presets.js customPaths takes them). */
function geometryOf(spPr) {
  const prst = kid(spPr, 'prstGeom', A);
  if (prst) {
    const av = {};
    for (const gd of kids(kid(prst, 'avLst', A), 'gd', A)) {
      const m = /^val\s+(-?\d+)/.exec(attr(gd, 'fmla', null) || '');
      if (m) av[attr(gd, 'name', null)] = Number(m[1]);
    }
    return { prst: attr(prst, 'prst', null) || 'rect', av };
  }
  const cust = kid(spPr, 'custGeom', A);
  if (cust) {
    const pt = (p) => [num(attr(p, 'x', null)) / EMU_PER_PX, num(attr(p, 'y', null)) / EMU_PER_PX];
    const paths = kids(kid(cust, 'pathLst', A), 'path', A).map((p) => ({
      w: num(attr(p, 'w', null)) / EMU_PER_PX, h: num(attr(p, 'h', null)) / EMU_PER_PX, fill: attr(p, 'fill', null) !== 'none', stroke: attr(p, 'stroke', null) !== '0' && attr(p, 'stroke', null) !== 'false',
      cmds: kids(p, null).filter((c) => c.namespaceURI === A).map((c) => ({
        op: c.localName, pts: kids(c, 'pt', A).map(pt),
        wR: num(attr(c, 'wR', null)) / EMU_PER_PX, hR: num(attr(c, 'hR', null)) / EMU_PER_PX, stAng: num(attr(c, 'stAng', null)), swAng: num(attr(c, 'swAng', null)),
      })),
    }));
    return { cust: paths };
  }
  return { prst: 'rect', av: {} };
}

/** A run's or paragraph's properties (a:rPr, a:defRPr, a:endParaRPr): { size (pt), bold, italic, underline, strike,
 *  color, font }, fields it doesn't set left out. */
function runPr(el) {
  if (!el) return {};
  const out = {};
  if (attr(el, 'sz', null) != null) out.size = num(attr(el, 'sz', null)) / 100;
  if (attr(el, 'b', null) != null) out.bold = attr(el, 'b', null) === '1' || attr(el, 'b', null) === 'true';
  if (attr(el, 'i', null) != null) out.italic = attr(el, 'i', null) === '1' || attr(el, 'i', null) === 'true';
  if (attr(el, 'u', null) != null) out.underline = attr(el, 'u', null) !== 'none';
  if (attr(el, 'strike', null) != null) out.strike = attr(el, 'strike', null) !== 'noStrike';
  const fill = fillOf(el);
  if (fill?.color) out.color = fill.color;
  const latin = attr(kid(el, 'latin', A), 'typeface', null);
  if (latin) out.font = latin;
  return out;
}

/** A shape's text (xdr:txBody): { paras: [{ align, runs: [{ text, ...runPr }], end (runPr of an empty paragraph) }],
 *  anchor ('t', 'ctr', 'b'), insets (px), wrap, vert, lvl1 (the list style's first level) }, or null when it has none. */
function textOf(txBody) {
  if (!txBody) return null;
  const body = kid(txBody, 'bodyPr', A);
  const lvl1 = kid(kid(txBody, 'lstStyle', A), 'lvl1pPr', A);
  const ins = (name, d) => emuToPx(num(attr(body, name, null), d));
  const paras = kids(txBody, 'p', A).map((p) => {
    const pPr = kid(p, 'pPr', A);
    const runs = [];
    for (let c = p.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1 || c.namespaceURI !== A) continue;
      if (c.localName === 'r' || c.localName === 'fld') runs.push({ text: kid(c, 't', A)?.textContent || '', ...runPr(kid(c, 'rPr', A)) });
      else if (c.localName === 'br') runs.push({ text: '\n', ...runPr(kid(c, 'rPr', A)) });
    }
    return { align: attr(pPr, 'algn', null) || attr(lvl1, 'algn', null) || 'l', def: { ...runPr(kid(lvl1, 'defRPr', A)), ...runPr(kid(pPr, 'defRPr', A)) }, runs, end: runPr(kid(p, 'endParaRPr', A)) };
  });
  if (!paras.some((p) => p.runs.some((r) => r.text.trim()))) return null;
  return {
    paras, anchor: attr(body, 'anchor', null) || 't', anchorCtr: attr(body, 'anchorCtr', null) === '1',
    insets: { l: ins('lIns', 91440), t: ins('tIns', 45720), r: ins('rIns', 91440), b: ins('bIns', 45720) },
    wrap: attr(body, 'wrap', null) !== 'none', vert: attr(body, 'vert', null) || 'horz', rot: num(attr(body, 'rot', null)) / 60000,
  };
}

/** A shape's style references (xdr:style): { line: { idx, color }, fill, font: { idx (major, minor), color } }. */
function styleOf(el) {
  if (!el) return null;
  const ref = (name) => { const r = kid(el, name, A); return r ? { idx: attr(r, 'idx', null), color: colorOf(r) } : null; };
  return { line: ref('lnRef'), fill: ref('fillRef'), font: ref('fontRef') };
}

/** The name, description and hidden flag of an element's non-visual properties. */
function nvOf(el) {
  const nv = kids(el, null).find((c) => c.namespaceURI === XDR && /^nv/.test(c.localName));
  const c = kid(nv, 'cNvPr', XDR);
  return { id: attr(c, 'id', null), name: attr(c, 'name', null) || '', descr: attr(c, 'descr', null) || '', hidden: attr(c, 'hidden', null) === '1' || attr(c, 'hidden', null) === 'true',
    txBox: attr(kid(nv, 'cNvSpPr', XDR), 'txBox', null) === '1' };
}

/** One element of a drawing (a picture, shape, connector or group): what draws it, or null for what isn't drawn. */
function itemOf(el, ctx) {
  const nv = nvOf(el);
  if (el.localName === 'grpSp') {
    const xfrm = xfrmOf(kid(kid(el, 'grpSpPr', XDR), 'xfrm', A));
    const fill = fillOf(kid(el, 'grpSpPr', XDR));
    const children = [];
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      let d = c;
      if (c.namespaceURI === MC && c.localName === 'AlternateContent') d = [...(kid(c, 'Choice', MC)?.childNodes || [])].find((x) => x.nodeType === 1 && x.namespaceURI === XDR);
      if (!d || d.namespaceURI !== XDR || !CONTENT.has(d.localName) || d.localName === 'graphicFrame' || d.localName === 'contentPart') continue;
      const it = itemOf(d, { ...ctx, groupFill: fill?.none ? null : fill });
      if (it) children.push(it);
    }
    return { type: 'group', ...nv, xfrm, children };
  }
  if (el.localName === 'pic') {
    const spPr = kid(el, 'spPr', XDR);
    const blipFill = kid(el, 'blipFill', XDR);
    const blip = kid(blipFill, 'blip', A);
    const src = ctx.image(attr(blip, 'embed', R) ?? attr(blip, 'r:embed', null), attr(blip, 'link', R));
    // (An SVG icon: its own picture in the blip's extension, the PNG in the blip for those that can't draw SVG.)
    const svgBlip = blip && [...blip.getElementsByTagNameNS(SVG_BLIP, 'svgBlip')][0];
    const svg = svgBlip ? ctx.image(attr(svgBlip, 'embed', R) ?? attr(svgBlip, 'r:embed', null)) : null;
    const rect = kid(blipFill, 'srcRect', A);
    const crop = rect ? ['l', 't', 'r', 'b'].reduce((o, k) => ({ ...o, [k]: num(attr(rect, k, null)) / 100000 }), {}) : null;
    return { type: 'picture', ...nv, xfrm: xfrmOf(kid(spPr, 'xfrm', A)), geom: geometryOf(spPr), src, svg, crop, line: lineOf(kid(spPr, 'ln', A)), fill: fillOf(spPr) };
  }
  if (el.localName === 'sp' || el.localName === 'cxnSp') {
    const spPr = kid(el, 'spPr', XDR);
    return { type: 'shape', connector: el.localName === 'cxnSp', ...nv, xfrm: xfrmOf(kid(spPr, 'xfrm', A)), geom: geometryOf(spPr), fill: fillOf(spPr), groupFill: ctx.groupFill || null,
      line: lineOf(kid(spPr, 'ln', A)), style: styleOf(kid(el, 'style', XDR)), text: textOf(kid(el, 'txBody', XDR)) };
  }
  return null;
}

/**
 * Sheet si's drawing objects: [{ k (its place among the sheet's anchors, which the edits go by), kind ('picture',
 * 'shape', 'group', 'chart', 'frame'), chart (readCharts's index, for a chart), name, descr, hidden, print (false:
 * not printed), anchor: { type, editAs, from, to, pos, ext }, item (what draws it, for the others) }], in the order
 * they are drawn. Anchors in a Fallback are left out (their k still counted).
 */
export function readObjects(wb, si) {
  const path = wb.sheets[si]?.path;
  if (!path) return [];
  const pkg = wb.pkg;
  const out = [];
  let k = 0;
  let charts = 0;
  for (const rel of pkg.readRels(path).filter((r) => r.type === REL_DRAWING)) {
    const doc = pkg.xml(rel.target);
    if (!doc) continue;
    const rels = pkg.readRels(rel.target);
    const image = (id, link) => {
      const r = rels.find((x) => x.id === (id || link));
      if (!r || r.external) return null;
      const bytes = pkg.files[r.target];
      if (!bytes) return null;
      const mime = MIME[r.target.split('.').pop().toLowerCase()] || 'application/octet-stream';
      return { path: r.target, mime, get url() { return dataUrl(bytes, mime); }, bytes };
    };
    for (const { el, fallback, top } of anchorsOf(doc)) {
      const index = k++;
      const content = contentOf(el);
      const type = el.localName;
      const ext = kid(el, 'ext', XDR);
      const pos = kid(el, 'pos', XDR);
      const anchor = {
        type, editAs: attr(el, 'editAs', null) || (type === 'oneCellAnchor' ? 'oneCell' : type === 'absoluteAnchor' ? 'absolute' : 'twoCell'),
        from: corner(kid(el, 'from', XDR)), to: corner(kid(el, 'to', XDR)),
        ext: ext ? { w: emuToPx(num(attr(ext, 'cx', null))), h: emuToPx(num(attr(ext, 'cy', null))) } : null,
        pos: pos ? { x: emuToPx(num(attr(pos, 'x', null))), y: emuToPx(num(attr(pos, 'y', null))) } : null,
      };
      const clientData = kid(el, 'clientData', XDR);
      const print = attr(clientData, 'fPrintsWithSheet', null) !== '0' && attr(clientData, 'fPrintsWithSheet', null) !== 'false';
      // (A chart: as readCharts counts them, the sheet's top-level two-cell anchors whose chart part is there.)
      const chartRef = content?.localName === 'graphicFrame' && [...content.getElementsByTagNameNS(C, 'chart')][0];
      if (chartRef) {
        const target = rels.find((r) => r.id === attr(chartRef, 'id', R))?.target;
        const isChart = top && type === 'twoCellAnchor' && target && pkg.xml(target);
        if (fallback) continue;
        out.push({ k: index, kind: isChart ? 'chart' : 'frame', chart: isChart ? charts++ : null, ...nvOf(content), anchor, print });
        continue;
      }
      if (fallback || !content) continue;
      const item = itemOf(content, { image });
      if (!item) { out.push({ k: index, kind: 'frame', ...nvOf(content), anchor, print }); continue; }
      out.push({ k: index, kind: item.type, name: item.name, descr: item.descr, hidden: item.hidden, anchor, print, item });
    }
  }
  return out;
}

/**
 * Where an object sits on the sheet, in pixels: { x, y, w, h }, by `rows` and `cols` (geometry.js axes: start(i),
 * size(i)). A two-cell anchor goes from corner to corner, whatever its editAs (Excel's PDF of drawings.xlsx: its
 * editAs oneCell picture over a hidden row drawn that much shorter, though its extent is the full height); a one-cell
 * anchor is its corner and extent, an absolute anchor its position and extent.
 */
export function anchorRect(anchor, rows, cols, item = null) {
  const at = (p) => ({ x: cols.start(p.c) + Math.min(p.dx, cols.size(p.c)), y: rows.start(p.r) + Math.min(p.dy, rows.size(p.r)) });
  const own = anchor.ext || (item?.xfrm && item.xfrm.w + item.xfrm.h > 0 ? { w: item.xfrm.w, h: item.xfrm.h } : null);
  if (anchor.type === 'absoluteAnchor') return { x: anchor.pos?.x || 0, y: anchor.pos?.y || 0, w: own?.w || 0, h: own?.h || 0 };
  const from = anchor.from ? at(anchor.from) : { x: 0, y: 0 };
  if (anchor.type === 'oneCellAnchor' || !anchor.to) return { ...from, w: own?.w || 0, h: own?.h || 0 };
  const to = anchor.to ? at(anchor.to) : from;
  return { ...from, w: to.x - from.x, h: to.y - from.y };
}

// --- Edits -----------------------------------------------------------------------------------------------------------

const ANCHOR_RE = /<(\w+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/\1?\2>/g;

/** Sheet si's anchors as text, in readObjects's order: [{ drawing (its part), text }]. */
function anchorTexts(wb, si) {
  const pkg = wb.pkg;
  const out = [];
  for (const rel of pkg.readRels(wb.sheets[si].path).filter((r) => r.type === REL_DRAWING)) {
    const text = pkg.text(rel.target);
    if (!text || !pkg.xml(rel.target)) continue;
    for (const m of text.matchAll(ANCHOR_RE)) out.push({ drawing: rel.target, text: m[0] });
  }
  return out;
}

/**
 * Object k of sheet si moved or sized: `from` and `to` its new corners ({ col, colOff, row, rowOff }, offsets in EMU),
 * `rect` its new place in EMU from the sheet's top left ({ x, y, w, h }). A two-cell anchor takes both corners, a
 * one-cell anchor its corner and the new extent, an absolute anchor the position and extent; the object's own offset
 * and size (its xfrm, a group's children scaled with it) go with them, as Excel writes them.
 */
export function setObjectAnchor(wb, si, k, from, to, rect) {
  const part = anchorTexts(wb, si)[k];
  if (!part) throw new Error('No such object');
  const pt = (tag, p, a) => a.replace(new RegExp(`<(\\w+:)?${tag}>[\\s\\S]*?</\\1?${tag}>`), (m, pre = '') => `<${pre}${tag}><${pre}col>${p.col}</${pre}col><${pre}colOff>${p.colOff}</${pre}colOff><${pre}row>${p.row}</${pre}row><${pre}rowOff>${p.rowOff}</${pre}rowOff></${pre}${tag}>`);
  let next = part.text;
  if (/^<(\w+:)?absoluteAnchor/.test(next)) next = next.replace(/(<(?:\w+:)?pos\b[^>]*?\sx=")-?\d+("[^>]*?\sy=")-?\d+/, `$1${Math.round(rect.x)}$2${Math.round(rect.y)}`);
  else next = pt('from', from, next);
  if (/^<(\w+:)?twoCellAnchor/.test(next)) next = pt('to', to, next);
  // (The anchor's own extent, then the first transform's offset and extent: the object's.)
  if (!/^<(\w+:)?twoCellAnchor/.test(next)) next = next.replace(/(<(?!a:)(?:\w+:)?ext\s+cx=")\d+("\s+cy=")\d+/, `$1${Math.round(rect.w)}$2${Math.round(rect.h)}`);
  next = next.replace(/(<a:xfrm\b[^>]*>\s*<a:off\s+x=")-?\d+("\s+y=")-?\d+("\s*\/>\s*<a:ext\s+cx=")\d+("\s+cy=")\d+/, `$1${Math.round(rect.x)}$2${Math.round(rect.y)}$3${Math.round(rect.w)}$4${Math.round(rect.h)}`);
  const text = wb.pkg.text(part.drawing);
  wb.pkg.setText(part.drawing, text.replace(part.text, () => next));
}

/**
 * Objects `ks` of sheet si deleted (charts among them): their anchors, then each relationship of the drawing that no
 * anchor left names, and the parts only those led to (pictures, chart parts).
 */
export function deleteObjects(wb, si, ks) {
  const pkg = wb.pkg;
  const list = anchorTexts(wb, si);
  const byDrawing = new Map();
  for (const k of [...new Set(ks)].sort((a, b) => b - a)) {
    const part = list[k];
    if (!part) continue;
    if (!byDrawing.has(part.drawing)) byDrawing.set(part.drawing, []);
    byDrawing.get(part.drawing).push(part.text);
  }
  for (const [drawing, gone] of byDrawing) {
    let text = pkg.text(drawing);
    const ids = new Set();
    for (const a of gone) {
      for (const m of a.matchAll(/\br:(?:id|embed|link|pict|dm|lo|qs|cs)="([^"]+)"/g)) ids.add(m[1]);
      // (An anchor in markup compatibility's Choice goes with its whole AlternateContent, Fallback and all.)
      const at = text.indexOf(a);
      if (at < 0) continue;
      const before = text.lastIndexOf('<mc:AlternateContent', at);
      const close = before >= 0 ? text.indexOf('</mc:AlternateContent>', before) : -1;
      if (before >= 0 && close > at && !text.slice(before, at).includes('</mc:AlternateContent>')) {
        const whole = text.slice(before, close + '</mc:AlternateContent>'.length);
        for (const m of whole.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)) ids.add(m[1]);
        text = text.replace(whole, '');
      } else text = text.slice(0, at) + text.slice(at + a.length);
    }
    pkg.setText(drawing, text);
    for (const id of ids) {
      if (text.includes(`"${id}"`)) continue; // (still used by another object)
      const r = pkg.readRels(drawing).find((x) => x.id === id);
      pkg.removeRel(drawing, id);
      if (r && !r.external && pkg.has(r.target)) pkg.removeTree(r.target);
    }
  }
}
