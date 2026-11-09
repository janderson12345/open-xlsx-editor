// The filled rectangles and stroked lines of every PDF in a folder, as JSON beside them, in pixels at 96 dpi from the
// page's top left: fills as ['fill', page, '#rrggbb', x, y, w, h] (a filled path's bounding box), strokes as
// ['line', page, '#rrggbb', x1, y1, x2, y2, width, dashed]. Cell fills and borders are what a sheet draws this way, so
// Excel's PDFs and ours can be compared for them like the text (compare-shapes.py). Text is left out.
// usage: node tools/excel-compare/shapes.cjs <pdfDir> <outDir>
const fs = require('fs');
const path = require('path');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node shapes.cjs <pdfDir> <outDir>'); process.exit(1); }
fs.mkdirSync(outDir, { recursive: true });
const OPS = pdfjs.OPS;
const PX = 4 / 3;
const hex = (c) => (typeof c === 'string' ? c.toLowerCase() : '#' + [...c].slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, '0')).join(''));
const gray = (g) => hex([g * 255, g * 255, g * 255]);
const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

async function shapesOf(file) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(file)), disableFontFace: true, verbosity: 0 }).promise;
  const out = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const h = page.view[3];
    const list = await page.getOperatorList();
    let gs = { m: [1, 0, 0, 1, 0, 0], fill: '#000000', stroke: '#000000', lw: 1, dash: false };
    const stack = [];
    let pathPts = []; // [[x, y], ...] in page points, per subpath
    const toPx = ([x, y]) => [+(x * PX).toFixed(2), +((h - y) * PX).toFixed(2)];
    for (let i = 0; i < list.fnArray.length; i++) {
      const fn = list.fnArray[i];
      const a = list.argsArray[i];
      if (fn === OPS.save) stack.push({ ...gs });
      else if (fn === OPS.restore) gs = stack.pop() || gs;
      else if (fn === OPS.transform) gs.m = mul(gs.m, a);
      else if (fn === OPS.setFillRGBColor) gs.fill = hex(a.length === 1 ? a[0] : a);
      else if (fn === OPS.setStrokeRGBColor) gs.stroke = hex(a.length === 1 ? a[0] : a);
      else if (fn === OPS.setFillGray) gs.fill = gray(a[0]);
      // (A tiling pattern or shading, as Excel prints a hatched fill: reported as such, not as the last colour.)
      else if (fn === OPS.setFillColorN || fn === OPS.shadingFill) gs.fill = 'pattern';
      else if (fn === OPS.setStrokeGray) gs.stroke = gray(a[0]);
      else if (fn === OPS.setLineWidth) gs.lw = a[0];
      else if (fn === OPS.setDash) gs.dash = !!(a[0] && a[0].length);
      else if (fn === OPS.constructPath) {
        const [ops, coords] = a;
        let k = 0;
        let sub = null;
        for (const op of ops) {
          if (op === OPS.rectangle) {
            const [x, y, w, hh] = coords.slice(k, k + 4); k += 4;
            pathPts.push([[x, y], [x + w, y], [x + w, y + hh], [x, y + hh]].map(([px, py]) => apply(gs.m, px, py)));
          } else if (op === OPS.moveTo) { sub = [apply(gs.m, coords[k], coords[k + 1])]; pathPts.push(sub); k += 2; }
          else if (op === OPS.lineTo) { (sub || pathPts[pathPts.length - 1] || []).push(apply(gs.m, coords[k], coords[k + 1])); k += 2; }
          else if (op === OPS.curveTo) { (sub || []).push(apply(gs.m, coords[k + 4], coords[k + 5])); k += 6; }
          else if (op === OPS.curveTo2 || op === OPS.curveTo3) { (sub || []).push(apply(gs.m, coords[k + 2], coords[k + 3])); k += 4; }
        }
      } else if (fn === OPS.fill || fn === OPS.eoFill || fn === OPS.fillStroke || fn === OPS.eoFillStroke || fn === OPS.closeFillStroke || fn === OPS.closeEOFillStroke) {
        for (const sub of pathPts) {
          if (sub.length < 3) continue;
          const pts = sub.map(toPx);
          const xs = pts.map((q) => q[0]);
          const ys = pts.map((q) => q[1]);
          const x0 = Math.min(...xs); const y0 = Math.min(...ys);
          out.push(['fill', p, gs.fill, x0, y0, +(Math.max(...xs) - x0).toFixed(2), +(Math.max(...ys) - y0).toFixed(2)]);
        }
        pathPts = [];
      } else if (fn === OPS.stroke || fn === OPS.closeStroke) {
        const scale = Math.hypot(gs.m[0], gs.m[1]);
        for (const sub of pathPts) {
          for (let j = 1; j < sub.length; j++) {
            const [x1, y1] = toPx(sub[j - 1]);
            const [x2, y2] = toPx(sub[j]);
            out.push(['line', p, gs.stroke, x1, y1, x2, y2, +(gs.lw * scale * PX).toFixed(2), gs.dash]);
          }
        }
        pathPts = [];
      } else if (fn === OPS.endPath) pathPts = [];
    }
  }
  return out;
}

(async () => {
  for (const f of fs.readdirSync(dir).filter((x) => /\.pdf$/i.test(x))) {
    try {
      fs.writeFileSync(path.join(outDir, f.replace(/\.pdf$/i, '.json')), JSON.stringify(await shapesOf(path.join(dir, f))));
    } catch (e) { console.log('FAILED', f, e.message); }
  }
})();
