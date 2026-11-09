// The text of every PDF in a folder, as JSON beside them: [[key, page, x, baseline y, width], ...] in pixels at 96 dpi
// from the page's top left, one entry per piece of cell text. Pieces on one baseline that touch (a gap of -0.5 to
// 1.5pt) are joined, since a PDF may split one cell's text into several runs; the key is the text without spaces, to 40
// characters (what entries are matched by). Excel's PDFs and ours go through this alike.
// usage: node tools/excel-compare/cells.cjs <pdfDir> <outDir>
const fs = require('fs');
const path = require('path');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node cells.cjs <pdfDir> <outDir>'); process.exit(1); }
fs.mkdirSync(outDir, { recursive: true });
const PX = 4 / 3; // pixels per point
// Spaces and zero-width spaces (U+200B) don't count in a key.
const SPACES = new RegExp('[' + String.fromCharCode(92) + 's' + String.fromCharCode(0x200b) + ']+', 'g');

(async () => {
  for (const f of fs.readdirSync(dir).filter((x) => /\.pdf$/i.test(x))) {
    try {
      const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(path.join(dir, f))), disableFontFace: true, verbosity: 0 }).promise;
      const out = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const h = page.view[3];
        const items = (await page.getTextContent()).items
          .filter((it) => it.str)
          .map((it) => ({ x: it.transform[4], y: h - it.transform[5], w: it.width, s: it.str }))
          .sort((a, b) => a.y - b.y || a.x - b.x);
        const joined = [];
        let last = null;
        for (const it of items) {
          // (Excel writes the gap between two cells' text as a run of spaces: a run of only spaces ends a piece.)
          if (!it.s.trim()) { last = null; continue; }
          const gap = last ? it.x - (last.x + last.w) : 0;
          // (Touching, not overlapping: text clipped by a narrow column overlaps the next cell's, and isn't joined.)
          if (last && Math.abs(last.y - it.y) <= 0.6 && gap > -0.5 && gap < 1.5) {
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
