// Prints every .xlsx in a folder to PDF through the editor's print layout, as a person printing from the browser would,
// for comparison with Excel's PDFs of the same files. Several at a time; each has a time limit, so one bad file can't
// hang the run; workbooks already printed are skipped, so a rerun picks up where it stopped. The page must be in
// standards mode (LESSONS.md section 4), or the run stops.
// usage: node tools/excel-compare/print-ours.mjs <xlsxDir> <outDir>   (from the repository, after `node build.mjs`)
//   BROWSER_CHANNEL=msedge (default) | chrome; CHROMIUM_PATH=<executable> instead; OXE_JOBS=4
//   OXE_EXPORT=1: the PDF made by src/pdf.js exportPdf (the installed fonts allowed) instead of the browser's print
//   OXE_FONTS=1: the demo with Open Metric Fonts' Morrow Sans served (demo/fonts/omf/morrow-sans; README, Fonts)
// Use the browser your users print from: its fonts and font rendering are part of what's measured (Edge on Windows).
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node print-ours.mjs <xlsxDir> <outDir>'); process.exit(1); }
const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const url = `http://localhost:${server.address().port}/${process.env.OXE_FONTS ? '?fonts' : ''}`;
fs.mkdirSync(outDir, { recursive: true });
const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : { channel: process.env.BROWSER_CHANNEL || 'msedge' });
const pdfName = (f) => path.join(outDir, f.replace(/\.xls[xm]$/i, '.pdf'));
const todo = fs.readdirSync(dir).filter((f) => /\.xls[xm]$/i.test(f) && !fs.existsSync(pdfName(f)));
let done = 0;
let quirks = false;
const one = async (f) => {
  const context = await browser.newContext({ viewport: { width: 1300, height: 1000 } });
  if (process.env.OXE_EXPORT) await context.grantPermissions(['local-fonts']).catch(() => {});
  const page = await context.newPage();
  try {
    await Promise.race([(async () => {
      await page.goto(url);
      if ((await page.evaluate(() => document.compatMode)) !== 'CSS1Compat') { quirks = true; throw new Error('demo page not in standards mode'); }
      await page.setInputFiles('input[type=file]', path.join(dir, f));
      await page.waitForFunction((n) => window.oxe.fileName === n || /Could not open/.test(document.getElementById('status').textContent), f);
      const pages = await page.evaluate(() => window.oxe.preparePrint());
      if (!pages) throw new Error(await page.evaluate(() => document.getElementById('status').textContent));
      await page.evaluate(() => document.fonts.ready);
      if (process.env.OXE_EXPORT) {
        const b64 = await page.evaluate(async () => {
          const { exportPdf } = await import('/dist/pdf.js');
          const bytes = await exportPdf(window.oxe);
          let s = '';
          for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
          return btoa(s);
        });
        fs.writeFileSync(pdfName(f), Buffer.from(b64, 'base64'));
      } else {
        await page.emulateMedia({ media: 'print' });
        await page.pdf({ path: pdfName(f), preferCSSPageSize: true, printBackground: true });
      }
    })(), new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 150000))]);
    done++;
  } catch (e) { console.log('FAILED', f, e.message); }
  await context.close();
};
const jobs = Number(process.env.OXE_JOBS || 4);
await Promise.all(Array.from({ length: jobs }, async () => { while (todo.length && !quirks) await one(todo.shift()); }));
await browser.close();
server.close();
console.log(`printed ${done}`);
if (quirks) process.exit(1);
