// What the browser suites share: a static server for demo/, a Chromium page on it with errors collected, and the
// standards-mode check every page must pass (LESSONS.md section 4: quirks mode once hid a 15-point layout gap).
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.map': 'application/json' };

export async function openDemo({ width = 1200, height = 900, colorScheme = 'light', query = '', route = null } = {}) {
  const root = path.resolve('demo');
  const server = http.createServer((req, res) => {
    const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
    if (req.url === '/favicon.ico') { res.writeHead(204); return res.end(); }
    if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { 'content-type': TYPES[path.extname(p)] || 'application/octet-stream' });
    fs.createReadStream(p).pipe(res);
  }).listen(0);
  const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
  const page = await browser.newPage({ viewport: { width, height }, colorScheme });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
  // (`route`: requests the suite answers itself, set up before the page loads: [pattern, handler].)
  if (route) await page.route(...route);
  await page.goto(`http://localhost:${server.address().port}/${query}`);
  const mode = await page.evaluate(() => document.compatMode);
  if (mode !== 'CSS1Compat') throw new Error(`demo page in ${mode}, not standards mode: it needs <!doctype html>`);
  const close = async () => { await browser.close(); server.close(); };
  return { page, errors, close };
}

/** Open a fixture through the page's file input and wait until it is shown. */
export async function openFixture(page, name) {
  await page.setInputFiles('input[type=file]', `test/fixtures/${name}`);
  await page.waitForFunction((n) => window.oxe.fileName === n, name);
}

/** Report one check and count failures. */
export function checker() {
  let failed = 0;
  const check = (ok, what, detail = '') => {
    console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}${!ok && detail ? `\n     ${detail}` : ''}`);
    if (!ok) failed++;
    return ok;
  };
  return { check, get failed() { return failed; } };
}
