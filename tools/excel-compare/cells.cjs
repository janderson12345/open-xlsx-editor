// The text of every PDF in a folder, as JSON beside them: [[key, page, x, baseline y, width], ...] in pixels at 96 dpi
// from the page's top left, one entry per piece of cell text. Pieces on one baseline that touch or are a space apart
// (a gap of -0.5pt to 0.3 em) are joined, since a PDF may split one cell's text into several runs; the key is the text without spaces, to 40
// characters (what entries are matched by). Excel's PDFs and ours go through this alike.
// usage: node tools/excel-compare/cells.cjs <pdfDir> <outDir>
const fs = require('fs');
const path = require('path');

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node cells.cjs <pdfDir> <outDir>'); process.exit(1); }
fs.mkdirSync(outDir, { recursive: true });
const PX = 4 / 3; // pixels per point
// Spaces and zero-width spaces (U+200B) don't count in a key.
const SPACES = new RegExp('[' + String.fromCharCode(92) + 's' + String.fromCharCode(0x200b) + ']+', 'g');

// (pdf.js ships as an ES module only: loaded with import(), awaited before use.)
(async () => {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  for (const f of fs.readdirSync(dir).filter((x) => /\.pdf$/i.test(x))) {
    try {
      const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(path.join(dir, f))), disableFontFace: true, verbosity: 0 }).promise;
      const out = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const h = page.view[3];
        const items = (await page.getTextContent()).items
          .filter((it) => it.str)
          .map((it) => ({ x: it.transform[4], y: h - it.transform[5], w: it.width, s: it.str, size: Math.hypot(it.transform[0], it.transform[1]) }))
          .sort((a, b) => a.y - b.y || a.x - b.x);
        const joined = [];
        let last = null;
        for (const it of items) {
          // (Excel writes the gap between two cells' text as a run of spaces, wide: a run of only spaces wider than half an
          // em ends a piece. A single space, as Chromium writes at a wrapped line's break, doesn't.)
          if (!it.s.trim()) { if (it.w > 0.5 * it.size) last = null; continue; }
          const gap = last ? it.x - (last.x + last.w) : 0;
          // (Touching or a space apart, not overlapping: text clipped by a narrow column overlaps the next cell's, and
          // isn't joined. Chromium writes wrapped words as runs of their own and leaves the space out, a gap of about
          // 0.23 em in Calibri; the gap between two cells' text is wider than 0.3 em.)
          if (last && Math.abs(last.y - it.y) <= 0.6 && gap > -0.5 && gap < Math.max(1.5, 0.3 * it.size)) {
            last.s += it.s;
            last.w = it.x + it.w - last.x;
          } else joined.push((last = { ...it }));
        }
        for (const it of joined) {
          const key = it.s.replace(SPACES, '').slice(0, 40);
          if (key) out.push([key, p, +(it.x * PX).toFixed(1), +(it.y * PX).toFixed(1), +(it.w * PX).toFixed(1)]);
        }
      }
      fs.writeFileSync(path.join(outDir, f.replace(/\.pdf$/i, '.json')), JSON.stringify(out));
    } catch (e) { console.log('FAILED', f, e.message); }
  }
})();
