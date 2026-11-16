// PDF export drawn from the editor's own print layout (pdf-lib, MIT): each printed page becomes a PDF page of the same
// paper size, and on it, at the places the layout put them, the fills, the borders and the cells' text, as real text
// (selectable and searchable), each piece clipped to its cell as the page clips it; header and footer pictures.
// Separate from the main bundle: import { exportPdf } from 'open-xlsx-editor/pdf'. Adapted from open-docx-editor's
// pdf.js.
//
// Fonts, the first of the cell's font stack that has the face: the fonts installed on the computer, read with the
// browser's Local Font Access (Chromium browsers: it asks the user once), and the open-metric-fonts twins the page serves
// (`fonts.omf`: pdfomf.js), embedded as subsets (fontkit); else the PDF standard fonts (Helvetica, Times, Courier and
// their bold and italic faces), each piece of text drawn at the width the layout measured, so the columns and pages are
// the layout's in any font. A standard font draws characters it lacks as "?" (it covers Windows-1252's only).
//
// Pictures, shapes and icons (editor/drawview.js) over the cells: their paths as vector paths, their pictures as
// images (PNG and JPEG as they are; an SVG icon by its PNG fallback; GIF, BMP and WebP, and an SVG with none, turned to
// PNG by the browser), clipped to their frames and to the page's cells, and their text as real text.
import { PDFDocument, StandardFonts, rgb, PDFOperator, PDFOperatorNames, PDFNumber, pushGraphicsState, popGraphicsState, beginText, endText, setFontAndSize, setTextMatrix, showText, setFillingRgbColor, setStrokingRgbColor, setLineWidth, setDashPattern, setLineCap, setLineJoin, setGraphicsState, LineCapStyle, LineJoinStyle, moveTo, lineTo, appendBezierCurve, closePath, fill, stroke, fillAndStroke, concatTransformationMatrix, drawObject, rectangle, clip, endPath } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { twinFor, twinSegments } from './pdfomf.js';

const PT = 0.75; // CSS px to PDF points

/**
 * The workbook open in `editor` (mountXlsxEditor's API) printed to a PDF: the sheet shown, or (`sheets: 'all'`) every
 * worksheet that can be shown, in tab order, as Excel's "Print Entire Workbook". Returns the PDF's bytes. Options:
 * `title` and `author` (document information); `localFonts` false not to ask for the installed fonts.
 */
export async function exportPdf(editor, { sheets = 'shown', title, author, localFonts = true } = {}) {
  const wb = editor.workbook;
  if (!wb) throw new Error('no workbook open');
  const root = editor.grid.el.closest('.oxe-root');
  const shown = editor.sheetIndex;
  const list = sheets === 'all' ? [...wb.sheets.keys()].filter((i) => wb.canShow(i) && wb.sheets[i].kind === 'worksheet') : [shown];
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  if (title) pdf.setTitle(title);
  if (author) pdf.setAuthor(author);
  pdf.setProducer('open-xlsx-editor');
  const sources = await fontSources(localFonts);
  // (The twins' folder, as mountXlsxEditor's `fonts.omf` gave it: index.js.)
  sources.omf = root.dataset.omf || '';
  const fonts = {};
  // (The print pages shown off-screen while they're measured, in print's colours: editor.css, .oxe-exporting.)
  root.classList.add('oxe-exporting');
  try {
    for (const i of list) {
      if (i !== editor.sheetIndex) editor.showSheet(i);
      editor.preparePrint();
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      for (const pageEl of root.querySelectorAll(':scope > .oxe-print > .oxe-page')) await drawPage(pdf, pageEl, { sources, fonts });
    }
  } finally {
    root.classList.remove('oxe-exporting');
    if (editor.sheetIndex !== shown) editor.showSheet(shown);
  }
  if (!pdf.getPageCount()) pdf.addPage([8.5 * 72, 11 * 72]); // (an empty sheet prints one blank page, as Excel says "nothing to print")
  return pdf.save();
}

/** The installed fonts by "family|bold|italic" in lower case, when the browser has Local Font Access and it's allowed. */
async function fontSources(local) {
  const key = (family, bold, italic) => `${family.toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
  const installed = new Map();
  if (local && typeof window !== 'undefined' && typeof window.queryLocalFonts === 'function') {
    try {
      for (const f of await window.queryLocalFonts()) {
        const s = (f.style || '').toLowerCase();
        const b = /bold|black|heavy|semibold|demibold/.test(s), i = /italic|oblique/.test(s);
        const k = key(f.family, b, i);
        if (!installed.has(k) || /^(regular|bold|italic|bold italic)$/.test(s)) installed.set(k, f);
        const named = key((f.fullName || '').replace(/\s+(bold|italic|oblique|regular)\b/gi, '').trim(), b, i);
        if (f.fullName && !installed.has(named)) installed.set(named, f);
      }
    } catch { /* not allowed: the standard fonts */ }
  }
  return { key, installed };
}

/** The font for a computed style: an installed one embedded as a subset, or a served twin (`{ twin }`: pdfomf.js),
 *  else a standard font of its kind. */
async function fontFor(pdf, cs, { sources, fonts }) {
  const bold = (parseInt(cs.fontWeight, 10) || 400) >= 600;
  const italic = /italic|oblique/.test(cs.fontStyle);
  for (const fam of cs.fontFamily.split(',').map((f) => f.trim().replace(/^["']|["']$/g, ''))) {
    const k = sources.key(fam, bold, italic);
    if (fonts[k] !== undefined) { if (fonts[k]) return fonts[k]; continue; }
    fonts[k] = null;
    const twin = await twinFor(pdf, fam, bold, italic, sources.omf);
    if (twin) return (fonts[k] = { twin });
    const f = sources.installed.get(k);
    if (!f) continue;
    try {
      const font = await pdf.embedFont(fromCollection(new Uint8Array(await (await f.blob()).arrayBuffer()), f.postscriptName), { subset: true });
      fonts[k] = { font, chars: new Set(font.getCharacterSet()) };
      return fonts[k];
    } catch { /* a font fontkit can't read: the next family */ }
  }
  const first = cs.fontFamily.toLowerCase().split(',')[0];
  const mono = /courier|consol|mono/.test(first);
  const serif = !mono && /times|cambria|georgia|garamond|book|serif|palatino|constantia/.test(first) && !/sans/.test(first);
  const name = (mono ? 'Courier' : serif ? 'TimesRoman' : 'Helvetica') + (bold && italic ? (serif ? 'BoldItalic' : 'BoldOblique') : bold ? 'Bold' : italic ? (serif ? 'Italic' : 'Oblique') : '');
  if (!fonts[name]) {
    const font = await pdf.embedFont(StandardFonts[name]);
    fonts[name] = { font, chars: new Set(font.getCharacterSet()) };
  }
  return fonts[name];
}

/** One printed page: its fills and borders, then its cells' text. */
async function drawPage(pdf, pageEl, ctx) {
  const pr = pageEl.getBoundingClientRect();
  const page = pdf.addPage([pr.width * PT, pr.height * PT]);
  const X = (x) => (x - pr.left) * PT;
  const Y = (y) => (pr.bottom - y) * PT;
  const areaEl = pageEl.querySelector('.oxe-page-area');
  const area = areaEl.getBoundingClientRect();
  const clipTo = (r) => [pushGraphicsState(), rectangle(X(r.left), Y(r.bottom), r.width * PT, r.height * PT), clip(), endPath()];
  // 1. Fills and border edges (decor.js) and cells' own backgrounds, inside the printable area, in the order the page
  //    paints them (the DOM's: fills, cells, edges).
  page.pushOperators(...clipTo(area));
  for (const el of areaEl.querySelectorAll('*')) {
    if (el.closest('.oxe-page-objclip')) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility !== 'visible' || cs.display === 'none') continue;
    const r = el.getBoundingClientRect();
    if (!r.width || !r.height || r.right <= area.left || r.left >= area.right || r.bottom <= area.top || r.top >= area.bottom) continue;
    const bg = color(cs.backgroundColor);
    if (bg) page.drawRectangle({ x: X(r.left), y: Y(r.bottom), width: r.width * PT, height: r.height * PT, color: bg.rgb, opacity: bg.a });
    for (const s of ['Top', 'Right', 'Bottom', 'Left']) {
      const w = parseFloat(cs[`border${s}Width`]) || 0;
      const c = color(cs[`border${s}Color`]);
      const style = cs[`border${s}Style`];
      if (!w || !c || style === 'none' || style === 'hidden') continue;
      const t = w * PT;
      const dash = style === 'dashed' ? [3 * t, 2 * t] : style === 'dotted' ? [t, t] : undefined;
      const line = (x1, y1, x2, y2) => page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: t, color: c.rgb, opacity: c.a, dashArray: dash });
      if (s === 'Top') line(X(r.left), Y(r.top + w / 2), X(r.right), Y(r.top + w / 2));
      if (s === 'Bottom') line(X(r.left), Y(r.bottom - w / 2), X(r.right), Y(r.bottom - w / 2));
      if (s === 'Left') line(X(r.left + w / 2), Y(r.top), X(r.left + w / 2), Y(r.bottom));
      if (s === 'Right') line(X(r.right - w / 2), Y(r.top), X(r.right - w / 2), Y(r.bottom));
    }
  }
  // Sparklines (sparkview.js): their SVG paths, at the cell's place and scale, clipped to the cell.
  for (const svg of areaEl.querySelectorAll('svg.oxe-spark')) {
    const r = svg.getBoundingClientRect();
    const vb = svg.viewBox.baseVal;
    if (!r.width || !r.height || !vb?.width) continue;
    const k = (r.width / vb.width) * PT;
    page.pushOperators(...clipTo(r));
    for (const p of svg.querySelectorAll('path')) {
      const fill = hexColor(p.getAttribute('fill'));
      const stroke = hexColor(p.getAttribute('stroke'));
      if (!fill && !stroke) continue;
      page.drawSvgPath(p.getAttribute('d'), { x: X(r.left), y: Y(r.top), scale: k, ...(fill ? { color: fill } : { borderColor: stroke, borderWidth: (parseFloat(p.getAttribute('stroke-width')) || 1) * k }) });
    }
    page.pushOperators(popGraphicsState());
  }
  page.pushOperators(popGraphicsState());
  // 2. Text: each cell's text node, line by line, clipped to the cell (and the cell to the printable area); headers
  //    and footers.
  const range = document.createRange();
  const canvas = document.createElement('canvas').getContext('2d');
  // (Measured once per font: a sheet has thousands of cells in a few.)
  ctx.ascents ||= new Map();
  // (In a zoomed box, the print scale, the computed font size is the unzoomed one: `z` is the element's own zoom.)
  const ascent = (cs, z = 1) => {
    const f = `${cs.fontStyle} ${cs.fontWeight} ${(parseFloat(cs.fontSize) || 16) * z}px ${cs.fontFamily}`;
    if (!ctx.ascents.has(f)) {
      canvas.font = f;
      ctx.ascents.set(f, canvas.measureText('Hg').fontBoundingBoxAscent ?? parseFloat(cs.fontSize) * 0.8);
    }
    return ctx.ascents.get(f);
  };
  // (The page's text: its cells', and its header's and footer's outside the printable area.)
  /** One text node, line by line, in its font, clipped as its kind of text is (`within`: to that rectangle). */
  const drawText = async (t, within = null) => {
    const el = t.parentElement;
    const cs = getComputedStyle(el);
    if (cs.visibility !== 'visible') return;
    const cell = el.closest('.oxe-cell');
    // (A cell's text clipped to its cell and the printable area; a header's or footer's to the page; an object's to
    // the page's cells, `within`.)
    const bound = within || (cell ? area : pr);
    const cr = cell ? cell.getBoundingClientRect() : bound;
    const box = { left: Math.max(cr.left, bound.left), top: Math.max(cr.top, bound.top), right: Math.min(cr.right, bound.right), bottom: Math.min(cr.bottom, bound.bottom) };
    if (box.right <= box.left || box.bottom <= box.top) return;
    const got = await fontFor(pdf, cs, ctx);
    const z = el.currentCSSZoom || 1;
    const size = (parseFloat(cs.fontSize) || 16) * z * PT;
    const c = color(cs.color) || { rgb: rgb(0, 0, 0) };
    // (Each font's key in this page's resources, once.)
    const keys = new Map();
    const keyOf = (font) => { if (!keys.has(font)) keys.set(font, page.node.newFontDictionary(font.name, font.ref)); return keys.get(font); };
    for (const piece of pieces(t, range)) {
      const plain = piece.text.replace(/[\u00a0\u2000-\u200a\u202f\u205f]/g, ' ').replace(/[\u200b-\u200d\u2060\ufeff\u00ad]/g, '').replace(/[\u2010\u2011\u2012\u2212]/g, '-');
      if (!plain.trim()) continue;
      // (The piece in its font, or a twin's alphabets in theirs: pdfomf.js.)
      const segs = got.twin ? await twinSegments(got.twin, plain)
        : [{ font: got.font, text: [...plain].map((ch) => (got.chars.has(ch.codePointAt(0)) ? ch : ch.trim() ? '?' : ' ')).join('') }];
      const natural = segs.reduce((w, sg) => w + sg.font.widthOfTextAtSize(sg.text, size), 0);
      const want = piece.r.width * PT;
      const ops = [...clipTo({ left: box.left, bottom: box.bottom, width: box.right - box.left, height: box.bottom - box.top }), beginText(), setFillingRgbColor(c.rgb.red, c.rgb.green, c.rgb.blue)];
      // (Scaled to the width the layout measured in the page's font: another font's letters fill the same room.)
      if (natural > 0 && Math.abs(want - natural) > 0.1) ops.push(PDFOperator.of(PDFOperatorNames.SetTextHorizontalScaling, [PDFNumber.of(+(100 * want / natural).toFixed(2))]));
      ops.push(setTextMatrix(1, 0, 0, 1, X(piece.r.left), Y(piece.r.top + ascent(cs, z))));
      for (const sg of segs) ops.push(setFontAndSize(keyOf(sg.font), size), showText(sg.font.encodeText(sg.text)));
      ops.push(endText(), popGraphicsState());
      page.pushOperators(...ops);
      // Underline and strikethrough, from the cell text's own decoration.
      const lines = cs.textDecorationLine || '';
      const y = Y(piece.r.top + ascent(cs, z));
      const thick = Math.max(0.5, size / 18);
      if (lines.includes('underline')) page.drawLine({ start: { x: X(piece.r.left), y: y - size * 0.12 }, end: { x: X(piece.r.left) + want, y: y - size * 0.12 }, thickness: thick, color: c.rgb });
      if (lines.includes('line-through')) page.drawLine({ start: { x: X(piece.r.left), y: y + size * 0.3 }, end: { x: X(piece.r.left) + want, y: y + size * 0.3 }, thickness: thick, color: c.rgb });
    }
  };
  const walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT, { acceptNode: (t) => (t.data.trim() && !t.parentElement.closest('.oxe-page-objclip') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT) });
  for (let t = walker.nextNode(); t; t = walker.nextNode()) await drawText(t);
  // 2b. Pictures, shapes and icons, over the cells and their text, in the drawing's order, each with its text.
  await drawObjects(pdf, page, pageEl, { pr, drawText, ctx, area });
  // 3. Header and footer pictures (&G), PNG or JPEG, where the layout put them.
  for (const img of pageEl.querySelectorAll('.oxe-hf-pic')) {
    const m = /^data:image\/(png|jpeg);base64,(.*)$/.exec(img.src);
    if (!m) continue;
    const bytes = Uint8Array.from(atob(m[2]), (ch) => ch.charCodeAt(0));
    ctx.images ||= new Map();
    if (!ctx.images.has(img.src)) ctx.images.set(img.src, await (m[1] === 'png' ? pdf.embedPng(bytes) : pdf.embedJpg(bytes)));
    const r = img.getBoundingClientRect();
    page.drawImage(ctx.images.get(img.src), { x: X(r.left), y: Y(r.bottom), width: r.width * PT, height: r.height * PT });
  }
}

/** A page's pictures and shapes (print.js drawObjectsBlock): each object's paint (drawpaint.js, kept on its box by
 *  drawview.js) mapped from the sheet's pixels to the page, clipped to its block of cells. */
async function drawObjects(pdf, page, pageEl, { pr, drawText, ctx, area }) {
  // (The sheet's pixels to the page's points: the object's box is where the page has it, at its zoom.)
  const toPdf = (r, box, z) => [z * PT, 0, 0, -z * PT, (r.left - box.x * z - pr.left) * PT, (pr.bottom - r.top + box.y * z) * PT];
  const pathOps = (cmds, q) => cmds.map((c) => {
    const p = (i) => [q[0] * c[i] + q[2] * c[i + 1] + q[4], q[1] * c[i] + q[3] * c[i + 1] + q[5]];
    if (c[0] === 'M') return moveTo(...p(1));
    if (c[0] === 'L') return lineTo(...p(1));
    if (c[0] === 'C') return appendBezierCurve(...p(1), ...p(3), ...p(5));
    return closePath();
  });
  const rgbOf = (css) => [1, 3, 5].map((i) => parseInt(css.slice(i, i + 2), 16) / 255);
  ctx.images ||= new Map();
  for (const clipEl of pageEl.querySelectorAll('.oxe-page-objclip')) {
    const cr = clipEl.getBoundingClientRect();
    const bound = { left: Math.max(cr.left, area.left), top: Math.max(cr.top, area.top), right: Math.min(cr.right, area.right), bottom: Math.min(cr.bottom, area.bottom) };
    if (bound.right <= bound.left || bound.bottom <= bound.top) continue;
    const clipOps = () => [rectangle((bound.left - pr.left) * PT, (pr.bottom - bound.bottom) * PT, (bound.right - bound.left) * PT, (bound.bottom - bound.top) * PT), clip(), endPath()];
    for (const objEl of clipEl.querySelectorAll('.oxe-obj')) {
      const ops = objEl._oxePaint || [];
      const box = objEl._oxeBox;
      const z = objEl.currentCSSZoom || 1;
      const q = toPdf(objEl.getBoundingClientRect(), box, z);
      const texts = [...objEl.querySelectorAll(':scope > .oxe-obj-text')];
      let n = 0;
      for (const op of ops) {
        if (op.t === 'path') {
          const out = [pushGraphicsState(), ...clipOps()];
          const f = op.fill;
          const l = op.line;
          if ((f && f.alpha < 1) || (l && l.alpha < 1)) out.push(setGraphicsState(page.maybeEmbedGraphicsState({ opacity: f?.alpha ?? 1, borderOpacity: l?.alpha ?? 1 })));
          if (f) out.push(setFillingRgbColor(...rgbOf(f.css)));
          if (l) {
            const w = l.width * z * PT;
            out.push(setStrokingRgbColor(...rgbOf(l.css)), setLineWidth(w), setLineJoin(LineJoinStyle.Round), setLineCap(l.cap === 'round' ? LineCapStyle.Round : l.cap === 'square' ? LineCapStyle.Projecting : LineCapStyle.Butt));
            if (l.dash) out.push(setDashPattern(l.dash.map((d) => d * w), 0));
          }
          out.push(...pathOps(op.cmds, q), f && l ? fillAndStroke() : f ? fill() : stroke(), popGraphicsState());
          page.pushOperators(...out);
        } else if (op.t === 'image') {
          const image = await embedPicture(pdf, op, ctx.images);
          if (!image) continue;
          const name = page.node.newXObject('Image', image.ref);
          const m = op.m;
          // (The picture's unit square, its top at 1 in PDF's space, to the sheet, then to the page.)
          const u = [m[0], m[1], -m[2], -m[3], m[2] + m[4], m[3] + m[5]];
          const t = [q[0] * u[0] + q[2] * u[1], q[1] * u[0] + q[3] * u[1], q[0] * u[2] + q[2] * u[3], q[1] * u[2] + q[3] * u[3], q[0] * u[4] + q[2] * u[5] + q[4], q[1] * u[4] + q[3] * u[5] + q[5]];
          page.pushOperators(pushGraphicsState(), ...clipOps(), ...pathOps(op.clip, q), clip(), endPath(), concatTransformationMatrix(...t), drawObject(name), popGraphicsState());
        } else if (op.t === 'text') {
          const el = texts[n++];
          // (Text in a box that isn't turned: placed where the page has it. A turned one isn't exported yet.)
          if (!el || Math.abs(op.m[1]) > 1e-6 || Math.abs(op.m[2]) > 1e-6 || op.m[0] < 0 || op.m[3] < 0) continue;
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (t) => (t.data.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT) });
          for (let t = walker.nextNode(); t; t = walker.nextNode()) await drawText(t, bound);
        }
      }
    }
  }
}

/** A picture's image, embedded once: PNG and JPEG as they are, others (an SVG with no PNG beside it among them) drawn
 *  by the browser into a PNG. */
async function embedPicture(pdf, op, cache) {
  const src = op.src;
  if (!src?.bytes) return null;
  if (cache.has(src.bytes)) return cache.get(src.bytes);
  let image = null;
  try {
    if (src.mime === 'image/png') image = await pdf.embedPng(src.bytes);
    else if (src.mime === 'image/jpeg') image = await pdf.embedJpg(src.bytes);
    else {
      const img = new Image();
      img.src = src.url;
      await img.decode();
      const k = src.mime === 'image/svg+xml' ? 4 : 1;
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round((img.naturalWidth || 300) * k));
      canvas.height = Math.max(1, Math.round((img.naturalHeight || 150) * k));
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
      image = blob ? await pdf.embedPng(new Uint8Array(await blob.arrayBuffer())) : null;
    }
  } catch { image = null; }
  cache.set(src.bytes, image);
  return image;
}

/** '#RRGGBB' as a pdf-lib colour, or null. */
const hexColor = (h) => (/^#[0-9a-f]{6}$/i.test(h || '') ? rgb(...[1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255)) : null);

/** A text node's pieces, one per line it's on (wrapped cell text): [{ text, r }]. */
function pieces(t, rg) {
  const out = [];
  const n = t.data.length;
  let from = 0;
  while (from < n) {
    rg.setStart(t, from); rg.setEnd(t, from + 1);
    const first = rg.getClientRects()[0];
    if (!first) { from++; continue; }
    let lo = from, hi = n - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      rg.setStart(t, mid); rg.setEnd(t, mid + 1);
      const q = rg.getClientRects()[0];
      if (q && Math.abs(q.top - first.top) < 2) lo = mid; else hi = mid - 1;
    }
    let end = lo + 1;
    while (end > from + 1 && t.data[end - 1] === ' ' && end < n) end--;
    rg.setStart(t, from); rg.setEnd(t, end);
    const r = rg.getBoundingClientRect();
    if (r.width > 0) out.push({ text: t.data.slice(from, end), r: { left: r.left, top: r.top, width: r.width, height: r.height } });
    from = lo + 1;
  }
  return out;
}

/** A font file of its own from a collection (.ttc): the font whose PostScript name is `ps` (else the first). */
function fromCollection(bytes, ps) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 12 || dv.getUint32(0) !== 0x74746366) return bytes; // 'ttcf'
  const offsets = Array.from({ length: dv.getUint32(8) }, (_, k) => dv.getUint32(12 + 4 * k));
  const tables = (at) => Array.from({ length: dv.getUint16(at + 4) }, (_, k) => {
    const r = at + 12 + 16 * k;
    return { tag: dv.getUint32(r), sum: dv.getUint32(r + 4), off: dv.getUint32(r + 8), len: dv.getUint32(r + 12) };
  });
  const psName = (at) => {
    const name = tables(at).find((t) => t.tag === 0x6e616d65); // 'name'
    if (!name) return null;
    const count = dv.getUint16(name.off + 2), strings = name.off + dv.getUint16(name.off + 4);
    for (let k = 0; k < count; k++) {
      const r = name.off + 6 + 12 * k;
      if (dv.getUint16(r + 6) !== 6) continue;
      const platform = dv.getUint16(r), len = dv.getUint16(r + 8), off = strings + dv.getUint16(r + 10);
      let s = '';
      if (platform === 3 || platform === 0) for (let j = 0; j < len; j += 2) s += String.fromCharCode(dv.getUint16(off + j));
      else for (let j = 0; j < len; j++) s += String.fromCharCode(bytes[off + j]);
      return s;
    }
    return null;
  };
  const at = offsets.find((o) => psName(o) === ps) ?? offsets[0];
  const list = tables(at);
  const head = 12 + 16 * list.length;
  const out = new Uint8Array(head + list.reduce((t, x) => t + ((x.len + 3) & ~3), 0));
  const ov = new DataView(out.buffer);
  ov.setUint32(0, dv.getUint32(at));
  for (let k = 4; k < 12; k += 2) ov.setUint16(k, dv.getUint16(at + k));
  let pos = head;
  list.forEach((t, k) => {
    const r = 12 + 16 * k;
    ov.setUint32(r, t.tag); ov.setUint32(r + 4, t.sum); ov.setUint32(r + 8, pos); ov.setUint32(r + 12, t.len);
    out.set(bytes.subarray(t.off, t.off + t.len), pos);
    pos += (t.len + 3) & ~3;
  });
  return out;
}

/** CSS color to pdf-lib's, with its alpha; null when transparent. */
function color(css) {
  const m = /rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/.exec(css || '');
  if (!m) return null;
  const a = m[4] == null ? 1 : +m[4];
  if (a < 0.02) return null;
  return { rgb: rgb(+m[1] / 255, +m[2] / 255, +m[3] / 255), a };
}
