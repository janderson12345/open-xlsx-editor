// What a drawing object (ooxml/drawobjects.js) paints, as a list the screen (drawview.js), print and PDF export
// (pdf.js) each draw the same way: paths (filled, stroked, with their dashes and arrowheads), pictures (an image put
// on the sheet by a matrix, clipped to the picture's frame and geometry, its crop shown as Excel shows it) and text
// boxes, all in the sheet's pixels with the object's rotation, flips and group placing already in them. Colours are
// the file's own, or the theme's through the shape's style (its fillRef, lnRef and fontRef), with DrawingML's colour
// changes (lumMod, lumOff, shade, tint, alpha). No DOM here.
import { presetPaths, customPaths, arrowHead, transformCmds, LINES } from '../ooxml/presets.js';
import { resolveDrawingColor } from '../ooxml/theme.js';

// --- Matrices: [a, b, c, d, e, f], x' = a x + c y + e, y' = b x + d y + f (SVG's order). ------------------------------
export const mul = (p, q) => [p[0] * q[0] + p[2] * q[1], p[1] * q[0] + p[3] * q[1], p[0] * q[2] + p[2] * q[3], p[1] * q[2] + p[3] * q[3], p[0] * q[4] + p[2] * q[5] + p[4], p[1] * q[4] + p[3] * q[5] + p[5]];
const T = (x, y) => [1, 0, 0, 1, x, y];
const S = (x, y) => [x, 0, 0, y, 0, 0];
const Rot = (deg) => { const r = (deg * Math.PI) / 180; const c = Math.cos(r); const s = Math.sin(r); return [c, s, -s, c, 0, 0]; };
export const apply = (m, [x, y]) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

/** The matrix from an element's own frame (0..w by 0..h) to its parent's, placed at (x, y), rotated round its middle,
 *  and flipped (`flips` false: not flipped, for text, which Excel keeps readable; turned over for flipV). */
function frame(x, y, w, h, rot, flipH, flipV, flips = true) {
  const turn = rot + (!flips && flipV ? 180 : 0);
  const mirror = flips ? S(flipH ? -1 : 1, flipV ? -1 : 1) : S(1, 1);
  return mul(mul(mul(T(x + w / 2, y + h / 2), Rot(turn)), mirror), T(-w / 2, -h / 2));
}

// --- Colours ---------------------------------------------------------------------------------------------------------
const lin = (v) => { v /= 255; return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
const srgb = (v) => Math.round(255 * Math.max(0, Math.min(1, v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055)));
const hex = (rgb) => `#${rgb.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
const parts = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));

/** A colour's luminance (HSL's, in fractions: not the integer HLS of cells' tints) times `mod`, plus `off`, as
 *  Excel's shapes have it (drawings.xlsx: accent 2, ED7D31, at 75% is C55A11). */
function luminance(css, mod, off) {
  const [r, g, b] = parts(css).map((v) => v / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const d = max - min;
  let h = 0;
  if (d) h = max === r ? ((g - b) / d + 6) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  const l0 = (max + min) / 2;
  const s = d ? d / (1 - Math.abs(2 * l0 - 1)) : 0;
  const l = Math.max(0, Math.min(1, l0 * mod + off));
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs((h % 2) - 1));
  const m = l - c / 2;
  const [rr, gg, bb] = h < 1 ? [c, x, 0] : h < 2 ? [x, c, 0] : h < 3 ? [0, c, x] : h < 4 ? [0, x, c] : h < 5 ? [x, 0, c] : [c, 0, x];
  return hex([rr, gg, bb].map((v) => Math.round((v + m) * 255)));
}

/** A colour (drawobjects.js colorOf) as { css: '#RRGGBB', alpha }, `ph` standing for a style's placeholder colour;
 *  null for none. */
export function resolve(c, theme, ph = null) {
  if (!c) return null;
  if (c.ph) { if (!ph) return null; c = { ...ph, mods: [...(ph.mods || []), ...(c.mods || [])] }; }
  let css = resolveDrawingColor(c.rgb ? { rgb: c.rgb } : { scheme: c.scheme }, theme);
  if (!css) return null;
  let alpha = 1;
  let lumMod = null;
  let lumOff = null;
  const flush = () => { if (lumMod != null || lumOff != null) { css = luminance(css, lumMod ?? 1, lumOff ?? 0); lumMod = null; lumOff = null; } };
  for (const [name, v] of c.mods || []) {
    if (name === 'lumMod') lumMod = v;
    else if (name === 'lumOff') lumOff = v;
    else {
      flush();
      if (name === 'alpha') alpha = v;
      else if (name === 'shade') css = hex(parts(css).map((x) => srgb(lin(x) * v)));
      else if (name === 'tint') css = hex(parts(css).map((x) => srgb(1 - (1 - lin(x)) * v)));
    }
  }
  flush();
  return { css, alpha };
}

const DASHES = { solid: null, dot: [1, 3], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
// (The theme's line styles by lnRef idx, its own or as Office's themes since 2013 have them: 0.5, 1 and 1.5pt.)
const LINE_STYLES = [6350, 12700, 19050];
/** Text's default: 11pt, black, in the theme's body font, as Excel's shapes and text boxes write theirs. */
const TEXT_PT = 11;

/** A shape's fill and line as painted: { fill: { css, alpha } | null, line: { css, alpha, width, dash, cap, head, tail } | null }. */
function looks(it, theme) {
  const style = it.style;
  let fill = null;
  if (it.fill?.none) fill = null;
  else if (it.fill?.color) fill = resolve(it.fill.color, theme);
  else if (it.fill?.group) fill = it.groupFill?.color ? resolve(it.groupFill.color, theme) : null;
  else if (style?.fill && Number(style.fill.idx) > 0) fill = resolve(style.fill.color, theme);
  let line = null;
  const ln = it.line;
  if (!ln?.none) {
    const ref = style?.line && Number(style.line.idx) > 0 ? style.line : null;
    const c = ln?.color ? resolve(ln.color, theme) : ref ? resolve(ref.color, theme) : null;
    if (c) {
      const width = ln?.w ?? (ref ? (theme?.lines || LINE_STYLES)[Math.min(3, Number(ref.idx)) - 1] / 9525 || 0.75 : 0.75);
      line = { ...c, width: Math.max(width, 0.5), dash: DASHES[ln?.dash] || null, cap: ln?.cap === 'rnd' ? 'round' : ln?.cap === 'sq' ? 'square' : 'butt', head: ln?.head || null, tail: ln?.tail || null };
    }
  }
  return { fill, line };
}

/** A run's font as drawn: { family, size (pt), bold, italic, underline, strike, color: { css, alpha } }. */
function fontOf(run, para, it, theme) {
  const pick = (k) => run[k] ?? para.def[k];
  const ref = it.style?.font;
  const face = pick('font') || (ref?.idx === 'major' ? '+mj-lt' : '+mn-lt');
  const family = face === '+mj-lt' || face === '+mj-ea' ? theme?.fonts?.major || 'Aptos Display' : face.startsWith('+mn') ? theme?.fonts?.minor || 'Aptos Narrow' : face;
  const color = (pick('color') && resolve(pick('color'), theme)) || (ref?.color && resolve(ref.color, theme)) || { css: '#000000', alpha: 1 };
  return { family, size: pick('size') ?? TEXT_PT, bold: !!pick('bold'), italic: !!pick('italic'), underline: !!pick('underline'), strike: !!pick('strike'), color };
}

/** A shape's text as drawn: paragraphs of runs with their fonts. */
function textBody(it, theme) {
  const t = it.text;
  return {
    anchor: t.anchor, wrap: t.wrap, insets: t.insets,
    paras: t.paras.map((p) => ({ align: p.align, runs: p.runs.map((r) => ({ text: r.text, ...fontOf(r, p, it, theme) })), empty: fontOf(p.end || {}, p, it, theme) })),
  };
}

/** The ends of a line's paths for its arrowheads, in the sheet's pixels: [{ tip, from }] for its first and last point. */
function ends(cmds) {
  const pts = [];
  for (const c of cmds) {
    if (c[0] === 'M' || c[0] === 'L') pts.push([c[1], c[2]]);
    else if (c[0] === 'C') pts.push([c[1], c[2]], [c[3], c[4]], [c[5], c[6]]);
  }
  if (pts.length < 2) return null;
  const head = { tip: pts[0], from: pts.find((p) => Math.hypot(p[0] - pts[0][0], p[1] - pts[0][1]) > 0.01) || pts[1] };
  const last = pts[pts.length - 1];
  const tail = { tip: last, from: [...pts].reverse().find((p) => Math.hypot(p[0] - last[0], p[1] - last[1]) > 0.01) || pts[pts.length - 2] };
  return { head, tail };
}

/** One element's paint, its frame mapped to the sheet by `m` (w by h its own size). */
function paintItem(it, m, w, h, mText, theme, out) {
  if (it.hidden) return;
  if (it.type === 'group') {
    const ch = it.xfrm?.ch;
    const sx = ch?.w ? w / ch.w : 1;
    const sy = ch?.h ? h / ch.h : 1;
    for (const c of it.children) {
      const x = c.xfrm || { x: 0, y: 0, w: 0, h: 0, rot: 0 };
      const cx = (x.x - (ch?.x || 0)) * sx;
      const cy = (x.y - (ch?.y || 0)) * sy;
      const cw = x.w * sx;
      const chh = x.h * sy;
      paintItem(c, mul(m, frame(cx, cy, cw, chh, x.rot || 0, x.flipH, x.flipV)), cw, chh, mul(m, frame(cx, cy, cw, chh, x.rot || 0, x.flipH, x.flipV, false)), theme, out);
    }
    return;
  }
  const geom = it.geom || { prst: 'rect', av: {} };
  const shape = geom.cust ? { paths: customPaths(geom.cust, w, h).map((p) => p.cmds), line: false, known: true, text: { x: 0, y: 0, w, h }, custom: customPaths(geom.cust, w, h) } : presetPaths(geom.prst, w, h, geom.av);
  if (it.type === 'picture') {
    const clip = shape.paths.flatMap((p) => transformCmds(p, m));
    if (it.src && /^image\/(png|jpeg|gif|bmp|webp|svg\+xml)$/.test(it.src.mime)) {
      const c = it.crop || { l: 0, t: 0, r: 0, b: 0 };
      const W = w / (1 - c.l - c.r || 1);
      const H = h / (1 - c.t - c.b || 1);
      out.push({ t: 'image', m: mul(m, [W, 0, 0, H, -c.l * W, -c.t * H]), clip, src: it.src, svg: it.svg && /svg/.test(it.svg.mime) ? it.svg : null, name: it.name });
    } else {
      // (A picture that can't be drawn here (EMF, WMF, TIFF, a linked file): a grey box where it is.)
      out.push({ t: 'path', cmds: clip, fill: { css: '#F2F2F2', alpha: 1 }, line: { css: '#A6A6A6', alpha: 1, width: 1, dash: null, cap: 'butt' }, missing: true });
    }
    const { line } = looks({ ...it, fill: { none: true } }, theme);
    if (line) out.push({ t: 'path', cmds: clip, fill: null, line });
    return;
  }
  const { fill, line } = looks(it, theme);
  const isLine = shape.line || (it.connector && !geom.cust);
  const subs = shape.custom || shape.paths.map((cmds) => ({ cmds, fill: true, stroke: true }));
  for (const p of subs) {
    const cmds = transformCmds(p.cmds, m);
    const f = !isLine && p.fill ? fill : null;
    const l = p.stroke ? line : null;
    if (f || l) out.push({ t: 'path', cmds, fill: f, line: l ? { ...l, head: undefined, tail: undefined } : null });
    // (Arrowheads at a line's ends: headEnd at its first point, tailEnd at its last.)
    if (l && (l.head || l.tail)) {
      const e = ends(cmds);
      for (const [which, spec] of [['head', l.head], ['tail', l.tail]]) {
        if (!e || !spec) continue;
        const a = arrowHead(spec.type, e[which].tip, e[which].from, l.width, spec.w, spec.len);
        if (a) out.push({ t: 'path', cmds: a.cmds, fill: a.filled ? { css: l.css, alpha: l.alpha } : null, line: a.filled ? null : { ...l, dash: null, head: undefined, tail: undefined } });
      }
    }
  }
  if (it.text && !isLine) {
    const box = shape.text;
    const ins = it.text.insets;
    const bx = box.x + ins.l;
    const by = box.y + ins.t;
    const bw = Math.max(0, box.w - ins.l - ins.r);
    const bh = Math.max(0, box.h - ins.t - ins.b);
    const turn = (it.text.vert === 'vert' || it.text.vert === 'eaVert' ? 90 : it.text.vert === 'vert270' ? 270 : 0) + (it.text.rot || 0);
    const side = turn % 180 ? [bh, bw] : [bw, bh];
    const tm = mul(mText, mul(mul(T(bx + bw / 2, by + bh / 2), Rot(turn)), T(-side[0] / 2, -side[1] / 2)));
    out.push({ t: 'text', m: tm, w: side[0], h: side[1], body: textBody(it, theme) });
  }
}

/**
 * Object `obj` (readObjects's) at `rect` on the sheet (drawobjects.js anchorRect): what it paints, in the sheet's
 * pixels, in the order painted.
 */
export function paintObject(obj, rect, theme) {
  const it = obj.item;
  if (!it) return [];
  const x = it.xfrm || { rot: 0 };
  const out = [];
  const f = frameOf(rect, x.rot || 0);
  paintItem(it, frame(f.x, f.y, f.w, f.h, x.rot || 0, x.flipH, x.flipV), f.w, f.h, frame(f.x, f.y, f.w, f.h, x.rot || 0, x.flipH, x.flipV, false), theme, out);
  return out;
}

/** An object's own frame from its anchor's rectangle: turned 45 to 135 degrees (or 225 to 315), the anchor holds the
 *  frame turned upright, its width and height swapped about its middle (Excel's PDF of drawings.xlsx: Arrow 9, at 45
 *  degrees over B18:D22, 184.6px across both ways, as the arrow of D's height by B:D's width turned is). */
export function frameOf(rect, rot) {
  const r = ((rot % 360) + 360) % 360;
  if (!((r >= 45 && r < 135) || (r >= 225 && r < 315))) return rect;
  return { x: rect.x + (rect.w - rect.h) / 2, y: rect.y + (rect.h - rect.w) / 2, w: rect.h, h: rect.w };
}

/** Whether an object draws as a line (its height or width may be 0). */
export const isLineObject = (obj) => obj.item?.type === 'shape' && (obj.item.connector || LINES.has(obj.item.geom?.prst));
