// The pictures painted on each page of a PDF, with their pixels: Excel prints each sparkline as a bitmap of its cell
// (docs/SPARKLINES.md), so its drawing is compared pixel by pixel. Writes <outDir>/<n>.png for each picture and
// <outDir>/images.json: [{ n, page, x, y, w, h (pixels at 96 dpi from the page's top left), width, height (the
// bitmap's own pixels) }], in the order painted.
// usage: node tools/excel-compare/sparkimages.cjs <file.pdf> <outDir>
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const [file, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node sparkimages.cjs <file.pdf> <outDir>'); process.exit(1); }
fs.mkdirSync(outDir, { recursive: true });

/** A PNG of RGBA pixels (no filter, one zlib stream). */
function png(width, height, rgba) {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]);
    const c = Buffer.alloc(4); c.writeUInt32BE(crc(td));
    return Buffer.concat([len, td, c]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) Buffer.from(rgba.buffer, rgba.byteOffset + y * width * 4, width * 4).copy(raw, y * (width * 4 + 1) + 1);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

/** pdf.js's image as RGBA (its kinds: 1 grayscale 1bpp, 2 RGB 24bpp, 3 RGBA 32bpp). */
function rgbaOf(img) {
  const { width, height, data, kind } = img;
  if (kind === 3) return new Uint8Array(data);
  const out = new Uint8Array(width * height * 4);
  if (kind === 2) {
    for (let i = 0, j = 0; i < width * height; i++, j += 3) { out[i * 4] = data[j]; out[i * 4 + 1] = data[j + 1]; out[i * 4 + 2] = data[j + 2]; out[i * 4 + 3] = 255; }
    return out;
  }
  const row = Math.ceil(width / 8);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const v = (data[y * row + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0;
    out.set([v, v, v, 255], (y * width + x) * 4);
  }
  return out;
}

(async () => {
  // (pdf.js ships as an ES module only: loaded with import(), awaited before use.)
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(file)), verbosity: 0 }).promise;
  const list = [];
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
      else if (fn === pdfjs.OPS.paintImageXObject) {
        const t = mul(vp.transform, m);
        const objs = args[0].startsWith('g_') ? page.commonObjs : page.objs; // (a picture used more than once is kept in the document's shared objects)
        const img = await new Promise((res) => objs.get(args[0], res));
        const n = list.length;
        fs.writeFileSync(path.join(outDir, `${n}.png`), png(img.width, img.height, rgbaOf(img)));
        list.push({ n, page: p, x: +t[4].toFixed(2), y: +(t[5] + t[3]).toFixed(2), w: +t[0].toFixed(2), h: +(-t[3]).toFixed(2), width: img.width, height: img.height });
      }
    }
  }
  fs.writeFileSync(path.join(outDir, 'images.json'), JSON.stringify(list, null, 1));
  console.log(`${list.length} images`);
})();
