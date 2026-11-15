// The images painted on each page of a PDF, one JSON line each: { page, x, y (top), w, h } in pixels at 96 dpi from the
// page's top left. Used for header and footer pictures (docs/PRINT.md); Excel's PDFs and ours go through it alike.
// usage: node tools/excel-compare/images.cjs <file.pdf>
const fs = require('fs');
// (pdf.js ships as an ES module only: loaded with import(), awaited before use.)
(async () => {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(process.argv[2])) }).promise;
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const vp = page.getViewport({ scale: 96 / 72 });
    const ops = await page.getOperatorList();
    const stack = []; let m = [1, 0, 0, 1, 0, 0];
    const mul = (a, b) => [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i], args = ops.argsArray[i];
      if (fn === pdfjs.OPS.save) stack.push(m); else if (fn === pdfjs.OPS.restore) m = stack.pop() || m;
      else if (fn === pdfjs.OPS.transform) m = mul(m, args);
      else if (fn === pdfjs.OPS.paintImageXObject || fn === pdfjs.OPS.paintInlineImageXObject || fn === pdfjs.OPS.paintJpegXObject) {
        const t = mul(vp.transform, m);
        const x = t[4], y = t[5], w = t[0], h = t[3];
        console.log(JSON.stringify({ page: p, x: +x.toFixed(2), y: +(y + h).toFixed(2), w: +w.toFixed(2), h: +(-h).toFixed(2) }));
      }
    }
  }
})();
