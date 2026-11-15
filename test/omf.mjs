// The open-metric-fonts twins served by the page (`fonts: { omf }`, the demo's ?fonts): a cell in a Microsoft font the
// computer lacks drawn in its twin, only that face's latin file fetched; a font the computer has fetches nothing; and
// exportPdf embeds the twin from the same folder. The twins' files are answered here with Liberation Sans (pdfjs-dist's
// copy, SIL OFL), so the suite needs no copy of the bundle: it checks which files are asked for and that what they hold
// is drawn and embedded, not the twins' own widths (measured with the real bundle: README, Fonts).
// Run after `npm run build`: node test/omf.mjs
import fs from 'node:fs';
import { PDFDocument, PDFName, PDFDict } from 'pdf-lib';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { openDemo, openFixture, checker } from './browser.mjs';

const t = checker();
const stand = fs.readFileSync('node_modules/pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf');
const asked = [];
const route = ['**/fonts/omf/**', (r) => { asked.push(new URL(r.request().url()).pathname); r.fulfill({ status: 200, contentType: 'font/woff2', body: stand }); }];
const { page, errors, close } = await openDemo({ query: '?fonts', route });
await openFixture(page, 'basic.xlsx');
// (basic.xlsx's own font is Aptos Narrow, which no computer has: its twin's regular and bold latin files.)
t.check(asked.length && asked.every((p) => /\/morrow-sans\/files\/morrowsans-narrow-(bold-)?latin\.woff2$/.test(p)),
  'only the files of the fonts shown that the computer lacks: Aptos Narrow’s twin, regular and bold, latin', asked.join(', '));

// (Seaford, a cloud font of Office's, isn't installed on Windows or Linux: its twin is open-metric-fonts' Tidewater.)
const installed = (fam) => page.evaluate((f) => {
  const ctx = document.createElement('canvas').getContext('2d');
  const w = (s) => { ctx.font = `40px ${s}`; return ctx.measureText('mmmmmmmmmmlli1WQ').width; };
  return w(`'${f}', monospace`) !== w('monospace');
}, fam);
const seafordHere = await installed('Seaford');
await page.evaluate(() => {
  const o = window.oxe;
  o.enter('H2', 'Twin text'); o.select('H2'); o.command('format', { font: { name: 'Seaford' } });
  o.enter('H3', 'Installed'); o.select('H3'); o.command('format', { font: { name: 'Tahoma' } });
});
await page.waitForFunction(() => [...document.fonts].some((f) => f.family === 'oxe-omf Seaford' && f.status === 'loaded'), null, { timeout: 10000 }).catch(() => {});
await page.evaluate(() => document.fonts.ready);
const cell = (row, col) => page.evaluate(([r, c]) => {
  // (The grid's indexes count its headers: row 2 is aria-rowindex 3.)
  const txt = document.querySelector(`.oxe-row[aria-rowindex="${r + 1}"] .oxe-cell[aria-colindex="${c + 1}"] .oxe-tx`);
  return txt ? { family: getComputedStyle(txt).fontFamily, width: txt.getBoundingClientRect().width } : null;
}, [row, col]);
const h2 = await cell(2, 8);
t.check(/^Seaford, "oxe-omf Seaford", /.test(h2?.family || ''), 'a Microsoft font’s stack: the font, then its twin', h2?.family);
const tidewater = asked.filter((p) => /\/tidewater\/files\//.test(p));
t.check(seafordHere || (tidewater.length === 1 && /\/tidewater-regular-latin\.woff2$/.test(tidewater[0])),
  'the computer lacking it: only its regular face’s latin file fetched', asked.join(', '));
// (Drawn in what the file holds: the text as wide as Liberation Sans draws it.)
const libWidth = await page.evaluate(async () => {
  const f = new FontFace('check-liberation', `url(${new URL('fonts/omf/x.woff2', location.href)})`);
  await f.load(); document.fonts.add(f);
  const ctx = document.createElement('canvas').getContext('2d');
  ctx.font = `${getComputedStyle(document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="9"] .oxe-tx')).fontSize} check-liberation`;
  return ctx.measureText('Twin text').width;
});
t.check(seafordHere || Math.abs(h2.width - libWidth) < 0.5, 'its text drawn and measured in the twin', `${h2?.width} / ${libWidth}`);
const tahomaHere = await installed('Tahoma');
t.check(!tahomaHere || !asked.some((p) => /\/tamarack\//.test(p)), 'a font the computer has (Tahoma, where installed): its twin never fetched', asked.join(', '));

// exportPdf: the twin embedded from the same folder, as a subset, its text real text.
const bytes = await page.evaluate(async () => {
  const { exportPdf } = await import('/dist/pdf.js');
  return [...await exportPdf(window.oxe, { localFonts: false })];
});
const doc = await PDFDocument.load(new Uint8Array(bytes));
const names = new Set();
for (const p of doc.getPages()) {
  const fonts = p.node.Resources()?.lookup(PDFName.of('Font'), PDFDict);
  for (const [, ref] of fonts?.entries() || []) names.add(doc.context.lookup(ref).get(PDFName.of('BaseFont'))?.toString());
}
t.check(seafordHere || [...names].some((n) => /LiberationSans/.test(n)), 'the PDF: the served twin embedded (the file the folder answered)', [...names].join(' '));
t.check(bytes.length < 400000, 'embedded as a subset', `${bytes.length} bytes`);
const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes), disableFontFace: true, verbosity: 0 }).promise;
let words = '';
for (let i = 1; i <= pdf.numPages; i++) words += (await (await pdf.getPage(i)).getTextContent()).items.map((x) => x.str).join(' ');
t.check(words.includes('Twin text') && words.includes('Installed'), 'its text real text', words.slice(-300));

t.check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
