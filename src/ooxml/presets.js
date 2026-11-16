// The outlines of DrawingML's preset shapes (prstGeom), as ECMA-376 Part 1's presetShapeDefinitions.xml draws them,
// for the common ones: rect, roundRect, ellipse, triangle, rtTriangle, diamond, parallelogram, trapezoid, pentagon,
// hexagon, octagon, plus, star5, the block arrows (right, left, up, down, left-right, up-down), chevron, homePlate,
// the lines and connectors (line, straightConnector1, bentConnector2 and 3, curvedConnector3) and the flowchart shapes
// that are these (process, alternate process, decision, terminator, connector). Any other is drawn as its bounding
// box (`known: false`). And custom geometry (custGeom): its paths' moveTo, lnTo, cubicBezTo, quadBezTo, arcTo, close.
// Each outline is a list of subpaths of absolute commands, ['M', x, y], ['L', x, y], ['C', x1, y1, x2, y2, x, y] and
// ['Z'], in the shape's own pixels (0..w by 0..h), so that the screen (SVG), print and PDF export can each draw it,
// transformed as they need. Arcs are drawn as cubic Béziers.

const pin = (lo, v, hi) => Math.max(lo, Math.min(hi, v));

/** A quarter-or-less elliptical arc from angle a to b (radians, clockwise on screen) round (cx, cy), radii rx, ry, as
 *  Bézier commands (the start point is where the path is). */
function arcCmds(cx, cy, rx, ry, a, b) {
  const out = [];
  const n = Math.max(1, Math.ceil(Math.abs(b - a) / (Math.PI / 2) - 1e-9));
  const step = (b - a) / n;
  const k = (4 / 3) * Math.tan(step / 4);
  for (let i = 0; i < n; i++) {
    const t0 = a + i * step;
    const t1 = t0 + step;
    const [c0, s0, c1, s1] = [Math.cos(t0), Math.sin(t0), Math.cos(t1), Math.sin(t1)];
    out.push(['C', cx + rx * (c0 - k * s0), cy + ry * (s0 + k * c0), cx + rx * (c1 + k * s1), cy + ry * (s1 - k * c1), cx + rx * c1, cy + ry * s1]);
  }
  return out;
}

const poly = (pts) => [['M', ...pts[0]], ...pts.slice(1).map((p) => ['L', ...p]), ['Z']];
const ellipse = (x, y, w, h) => {
  const cx = x + w / 2;
  const cy = y + h / 2;
  return [['M', x + w, cy], ...arcCmds(cx, cy, w / 2, h / 2, 0, Math.PI * 2), ['Z']];
};
/** A rectangle with corners of radii rx, ry. */
const rounded = (w, h, rx, ry) => {
  if (rx <= 0 || ry <= 0) return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  const P = Math.PI;
  return [['M', rx, 0], ['L', w - rx, 0], ...arcCmds(w - rx, ry, rx, ry, -P / 2, 0), ['L', w, h - ry], ...arcCmds(w - rx, h - ry, rx, ry, 0, P / 2),
    ['L', rx, h], ...arcCmds(rx, h - ry, rx, ry, P / 2, P), ['L', 0, ry], ...arcCmds(rx, ry, rx, ry, P, 1.5 * P), ['Z']];
};

/** The shapes drawn as lines: no fill, their ends where arrowheads go. */
export const LINES = new Set(['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'curvedConnector3']);

/** The default adjust values (adj, or adj1, adj2...) by shape. */
const ADJ = {
  roundRect: { adj: 16667 }, triangle: { adj: 50000 }, parallelogram: { adj: 25000 }, trapezoid: { adj: 25000 }, hexagon: { adj: 25000, vf: 115470 },
  octagon: { adj: 29289 }, plus: { adj: 25000 }, star5: { adj: 19098, hf: 105146, vf: 110557 }, pentagon: { hf: 105146, vf: 110557 },
  rightArrow: { adj1: 50000, adj2: 50000 }, leftArrow: { adj1: 50000, adj2: 50000 }, upArrow: { adj1: 50000, adj2: 50000 }, downArrow: { adj1: 50000, adj2: 50000 },
  leftRightArrow: { adj1: 50000, adj2: 50000 }, upDownArrow: { adj1: 50000, adj2: 50000 }, chevron: { adj: 50000 }, homePlate: { adj: 50000 },
  bentConnector3: { adj1: 50000 }, curvedConnector3: { adj1: 50000 },
};

/**
 * Preset `prst` at w by h with adjust values `av` ({ adj: 25000, ... } as the file gives them, the rest the
 * defaults): { paths: [[commands]], line (a line, not filled), known (false: drawn as its box), text: { x, y, w, h }
 * (the box its text goes in, before the insets) }.
 */
export function presetPaths(prst, w, h, av = {}) {
  const a = { ...ADJ[prst], ...av };
  const ss = Math.min(w, h);
  const hc = w / 2;
  const vc = h / 2;
  const box = { x: 0, y: 0, w, h };
  const out = (paths, text = box, extra = {}) => ({ paths, line: LINES.has(prst), known: true, text, ...extra });
  switch (prst) {
    case 'rect': case 'flowChartProcess': return out([poly([[0, 0], [w, 0], [w, h], [0, h]])]);
    case 'roundRect': {
      const r = ss * pin(0, a.adj, 50000) / 100000;
      const il = r * 0.29289;
      return out([rounded(w, h, r, r)], { x: il, y: il, w: w - 2 * il, h: h - 2 * il });
    }
    case 'flowChartAlternateProcess': { const r = ss / 6; return out([rounded(w, h, r, r)]); }
    case 'flowChartTerminator': { const rx = w * 3475 / 21600; return out([rounded(w, h, rx, h / 2)], { x: w * 1018 / 21600, y: h * 3163 / 21600, w: w * (20582 - 1018) / 21600, h: h * (18437 - 3163) / 21600 }); }
    case 'ellipse': case 'flowChartConnector': {
      const ix = (w / 2) * (1 - Math.SQRT1_2);
      const iy = (h / 2) * (1 - Math.SQRT1_2);
      return out([ellipse(0, 0, w, h)], { x: ix, y: iy, w: w - 2 * ix, h: h - 2 * iy });
    }
    case 'triangle': {
      const x = w * pin(0, a.adj, 100000) / 100000;
      return out([poly([[0, h], [x, 0], [w, h]])], { x: x / 2, y: vc, w: (x + w) / 2 - x / 2, h: vc });
    }
    case 'rtTriangle': return out([poly([[0, h], [0, 0], [w, h]])], { x: w / 12, y: h * 7 / 12, w: w / 2 - w / 12, h: h * 11 / 12 - h * 7 / 12 });
    case 'diamond': case 'flowChartDecision': return out([poly([[hc, 0], [w, vc], [hc, h], [0, vc]])], { x: w / 4, y: h / 4, w: w / 2, h: h / 2 });
    case 'parallelogram': {
      const x2 = ss * pin(0, a.adj, 100000 * w / ss) / 100000;
      return out([poly([[0, h], [x2, 0], [w, 0], [w - x2, h]])]);
    }
    case 'trapezoid': {
      const x2 = ss * pin(0, a.adj, 50000 * w / ss) / 100000;
      return out([poly([[0, h], [x2, 0], [w - x2, 0], [w, h]])]);
    }
    case 'pentagon': {
      const swd2 = (w / 2) * a.hf / 100000;
      const shd2 = (h / 2) * a.vf / 100000;
      const svc = vc * a.vf / 100000;
      const r = (d) => (d * Math.PI) / 180;
      const dx1 = swd2 * Math.cos(r(18));
      const dx2 = swd2 * Math.cos(r(54));
      const y1 = svc - shd2 * Math.sin(r(18));
      const y2 = svc + shd2 * Math.sin(r(54));
      return out([poly([[hc - dx1, y1], [hc, 0], [hc + dx1, y1], [hc + dx2, y2], [hc - dx2, y2]])]);
    }
    case 'hexagon': {
      const x1 = ss * pin(0, a.adj, 50000 * w / ss) / 100000;
      const dy1 = (h / 2) * a.vf / 100000 * Math.sin(Math.PI / 3);
      return out([poly([[0, vc], [x1, vc - dy1], [w - x1, vc - dy1], [w, vc], [w - x1, vc + dy1], [x1, vc + dy1]])]);
    }
    case 'octagon': {
      const x1 = ss * pin(0, a.adj, 50000) / 100000;
      return out([poly([[0, x1], [x1, 0], [w - x1, 0], [w, x1], [w, h - x1], [w - x1, h], [x1, h], [0, h - x1]])]);
    }
    case 'plus': {
      const x1 = ss * pin(0, a.adj, 50000) / 100000;
      return out([poly([[0, x1], [x1, x1], [x1, 0], [w - x1, 0], [w - x1, x1], [w, x1], [w, h - x1], [w - x1, h - x1], [w - x1, h], [x1, h], [x1, h - x1], [0, h - x1]])]);
    }
    case 'star5': {
      const swd2 = (w / 2) * a.hf / 100000;
      const shd2 = (h / 2) * a.vf / 100000;
      const svc = vc * a.vf / 100000;
      const inner = pin(0, a.adj, 50000) / 50000;
      const pts = [];
      for (let i = 0; i < 10; i++) {
        const ang = -Math.PI / 2 + (i * Math.PI) / 5;
        const k = i % 2 ? inner : 1;
        pts.push([hc + swd2 * k * Math.cos(ang), svc + shd2 * k * Math.sin(ang)]);
      }
      return out([poly(pts)]);
    }
    case 'rightArrow': case 'leftArrow': {
      const dx1 = ss * pin(0, a.adj2, 100000 * w / ss) / 100000;
      const dy1 = h * pin(0, a.adj1, 100000) / 200000;
      const pts = [[0, vc - dy1], [w - dx1, vc - dy1], [w - dx1, 0], [w, vc], [w - dx1, h], [w - dx1, vc + dy1], [0, vc + dy1]];
      return out([poly(prst === 'leftArrow' ? pts.map(([x, y]) => [w - x, y]) : pts)]);
    }
    case 'downArrow': case 'upArrow': {
      const dy1 = ss * pin(0, a.adj2, 100000 * h / ss) / 100000;
      const dx1 = w * pin(0, a.adj1, 100000) / 200000;
      const pts = [[hc - dx1, 0], [hc + dx1, 0], [hc + dx1, h - dy1], [w, h - dy1], [hc, h], [0, h - dy1], [hc - dx1, h - dy1]];
      return out([poly(prst === 'upArrow' ? pts.map(([x, y]) => [x, h - y]) : pts)]);
    }
    case 'leftRightArrow': {
      const x2 = ss * pin(0, a.adj2, 50000 * w / ss) / 100000;
      const dy = h * pin(0, a.adj1, 100000) / 200000;
      return out([poly([[0, vc], [x2, 0], [x2, vc - dy], [w - x2, vc - dy], [w - x2, 0], [w, vc], [w - x2, h], [w - x2, vc + dy], [x2, vc + dy], [x2, h]])]);
    }
    case 'upDownArrow': {
      const y2 = ss * pin(0, a.adj2, 50000 * h / ss) / 100000;
      const dx = w * pin(0, a.adj1, 100000) / 200000;
      return out([poly([[hc, 0], [w, y2], [hc + dx, y2], [hc + dx, h - y2], [w, h - y2], [hc, h], [0, h - y2], [hc - dx, h - y2], [hc - dx, y2], [0, y2]])]);
    }
    case 'chevron': {
      const x1 = ss * pin(0, a.adj, 100000 * w / ss) / 100000;
      return out([poly([[0, 0], [w - x1, 0], [w, vc], [w - x1, h], [0, h], [x1, vc]])]);
    }
    case 'homePlate': {
      const x1 = w - ss * pin(0, a.adj, 100000 * w / ss) / 100000;
      return out([poly([[0, 0], [x1, 0], [w, vc], [x1, h], [0, h]])]);
    }
    case 'line': case 'straightConnector1': return out([[['M', 0, 0], ['L', w, h]]]);
    case 'bentConnector2': return out([[['M', 0, 0], ['L', w, 0], ['L', w, h]]]);
    case 'bentConnector3': { const x1 = w * a.adj1 / 100000; return out([[['M', 0, 0], ['L', x1, 0], ['L', x1, h], ['L', w, h]]]); }
    case 'curvedConnector3': {
      const x2 = w * a.adj1 / 100000;
      return out([[['M', 0, 0], ['C', x2 / 2, 0, x2, h / 4, x2, vc], ['C', x2, h * 3 / 4, (w + x2) / 2, h, w, h]]]);
    }
    default: return { paths: [poly([[0, 0], [w, 0], [w, h], [0, h]])], line: false, known: false, text: box };
  }
}

/**
 * Custom geometry's paths at w by h: `paths` as the file has them, [{ w, h, fill (false for fill="none"), stroke,
 * cmds: [{ op: 'moveTo'|'lnTo'|'cubicBezTo'|'quadBezTo'|'arcTo'|'close', pts: [[x, y]], wR, hR, stAng, swAng }] }],
 * their coordinates in each path's own w by h (the shape's when 0). Guides are not worked out: points that name one
 * are taken as 0. Returns [{ cmds (commands as presetPaths's), fill, stroke }].
 */
export function customPaths(paths, w, h) {
  return paths.map((p) => {
    const sx = p.w ? w / p.w : 1;
    const sy = p.h ? h / p.h : 1;
    const out = [];
    let cur = [0, 0];
    let start = [0, 0];
    const at = ([x, y]) => [x * sx, y * sy];
    for (const c of p.cmds) {
      if (c.op === 'moveTo') { cur = at(c.pts[0]); start = cur; out.push(['M', ...cur]); }
      else if (c.op === 'lnTo') { cur = at(c.pts[0]); out.push(['L', ...cur]); }
      else if (c.op === 'cubicBezTo') { const [a, b, e] = c.pts.map(at); out.push(['C', ...a, ...b, ...e]); cur = e; }
      else if (c.op === 'quadBezTo') {
        const [q, e] = c.pts.map(at);
        out.push(['C', cur[0] + (2 / 3) * (q[0] - cur[0]), cur[1] + (2 / 3) * (q[1] - cur[1]), e[0] + (2 / 3) * (q[0] - e[0]), e[1] + (2 / 3) * (q[1] - e[1]), ...e]);
        cur = e;
      } else if (c.op === 'arcTo') {
        // (From the current point, on the ellipse of radii wR, hR, from angle stAng through swAng, 60000ths of a degree.)
        const rx = c.wR * sx;
        const ry = c.hR * sy;
        const a0 = (c.stAng / 60000) * (Math.PI / 180);
        const sw = (c.swAng / 60000) * (Math.PI / 180);
        if (!rx || !ry || !sw) continue;
        const cx = cur[0] - rx * Math.cos(a0);
        const cy = cur[1] - ry * Math.sin(a0);
        const cmds = arcCmds(cx, cy, rx, ry, a0, a0 + sw);
        out.push(...cmds);
        const last = cmds[cmds.length - 1];
        cur = [last[5], last[6]];
      } else if (c.op === 'close') { out.push(['Z']); cur = start; }
    }
    return { cmds: out, fill: p.fill !== false, stroke: p.stroke !== false };
  });
}

/** Commands as an SVG path's `d`. */
export const pathData = (cmds) => cmds.map((c) => c[0] + (c.length > 1 ? ` ${c.slice(1).map((v) => +v.toFixed(2)).join(' ')}` : '')).join(' ');

/** Commands with each point put through `m` ([a, b, c, d, e, f]: x' = a x + c y + e, y' = b x + d y + f). */
export const transformCmds = (cmds, m) => cmds.map((c) => {
  if (c[0] === 'Z') return c;
  const out = [c[0]];
  for (let i = 1; i < c.length; i += 2) out.push(m[0] * c[i] + m[2] * c[i + 1] + m[4], m[1] * c[i] + m[3] * c[i + 1] + m[5]);
  return out;
});

/**
 * An arrowhead at the end of a line that arrives at `tip` from `from` (a point before it on the line), for a line of
 * width `lw`: `type` (triangle, stealth, diamond, oval: filled; arrow: two strokes), `w` and `len` ('sm', 'med', 'lg':
 * 2, 3 and 5 times the line's width, and no less than 2, 3 and 5 times 2.67px: Excel's PDF of drawings.xlsx has its
 * medium arrowheads 8px long and wide on lines of 1.33 and 2px). { cmds, filled }, or null for none.
 */
export function arrowHead(type, tip, from, lw, w = 'med', len = 'med') {
  if (!type || type === 'none') return null;
  const size = { sm: 2, med: 3, lg: 5 };
  const hw = (Math.max(lw, 8 / 3) * (size[w] || 3)) / 2;
  const hl = Math.max(lw, 8 / 3) * (size[len] || 3);
  let dx = tip[0] - from[0];
  let dy = tip[1] - from[1];
  const n = Math.hypot(dx, dy) || 1;
  dx /= n; dy /= n;
  const back = (d, side) => [tip[0] - dx * d - dy * side, tip[1] - dy * d + dx * side];
  if (type === 'arrow') return { cmds: [['M', ...back(hl, hw)], ['L', ...tip], ['L', ...back(hl, -hw)]], filled: false };
  if (type === 'stealth') return { cmds: poly([tip, back(hl, hw), back(hl * 0.6, 0), back(hl, -hw)]), filled: true };
  if (type === 'diamond') return { cmds: poly([back(-hl / 2, 0), back(0, hw), back(hl / 2, 0), back(0, -hw)]), filled: true };
  if (type === 'oval') return { cmds: ellipse(tip[0] - hw, tip[1] - hw, 2 * hw, 2 * hw), filled: true };
  return { cmds: poly([tip, back(hl, hw), back(hl, -hw)]), filled: true };
}
