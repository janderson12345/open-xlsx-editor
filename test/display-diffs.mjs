// Browser check of three display rules measured on Excel: gridlines left out under text that runs on into empty
// cells (on screen, read off a screenshot, and in print, against Excel's PDF: test/runover-excel.json), column widths
// by the Normal style's maximum digit width (test/digitwidth-excel.json), and the East Asian built-in date formats
// by locale (test/eastasia-excel.json).
// Run after `npm run build`: node test/display-diffs.mjs [outDir]
import fs from 'node:fs';
import { zipSync, strToU8 } from 'fflate';
import { openDemo, checker } from './browser.mjs';

const outDir = process.argv[2] || 'test/out';
fs.mkdirSync(outDir, { recursive: true });
const { page, errors, close } = await openDemo({ width: 1100, height: 900 });
const t = checker();
const { check } = t;
const RUNOVER = JSON.parse(fs.readFileSync('test/runover-excel.json', 'utf8'));
const DIGITS = JSON.parse(fs.readFileSync('test/digitwidth-excel.json', 'utf8')).fonts;

/** A one-sheet workbook: Normal style font `font` (name, size), cells [[ref, value, numFmtId]], cols [[n, width]]. */
function book({ font = ['Calibri', 11], cells = [], cols = [] }) {
  const ids = [...new Set(cells.map((c) => c[2] || 0))];
  const xfs = ids.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('');
  const rows = new Map();
  for (const [ref, v, id] of cells) {
    const r = Number(ref.slice(1));
    if (!rows.has(r)) rows.set(r, []);
    rows.get(r).push(`<c r="${ref}" s="${ids.indexOf(id || 0)}"><v>${v}</v></c>`);
  }
  const colXml = cols.length ? `<cols>${cols.map(([n, w]) => `<col min="${n}" max="${n}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
  const files = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="${font[1]}"/><name val="${font[0]}"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${ids.length}">${xfs}</cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${colXml}<sheetData>${[...rows].sort((a, b) => a[0] - b[0]).map(([r, cs]) => `<row r="${r}">${cs.join('')}</row>`).join('')}</sheetData></worksheet>`,
  };
  return Buffer.from(zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)]))));
}
const open = async (name, buffer) => {
  await page.setInputFiles('input[type=file]', { name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer });
  await page.waitForFunction((n) => window.oxe.fileName === n, name);
};

// 1. Gridlines under text running on: Excel's own runover.xlsx (Aptos Narrow 11, 8.43 columns).
await open('runover-excel.xlsx', fs.readFileSync('test/runover-excel.xlsx'));
const want = new Set(RUNOVER.hidden.map(([r, k]) => `${r},${k}`));
// On screen: a screenshot of the sheet, each column edge's gridline pixel read in each row.
// (Read 2px below each row's top, above the text's glyphs, with the selection moved off the sheet's corner.)
await page.evaluate(() => window.oxe.select('J30'));
const where = await page.evaluate(() => {
  const g = window.oxe.grid;
  const v = g.el.getBoundingClientRect();
  const { rows, cols, rowHeadW, colHeadH } = g.geometry();
  return { x: v.left + rowHeadW - g.el.scrollLeft, y: v.top + colHeadH - g.el.scrollTop, starts: Array.from({ length: 10 }, (_, k) => cols.start(k)), mids: Array.from({ length: 34 }, (_, r) => rows.start(r) + 2) };
});
const shot = await page.screenshot({ path: `${outDir}/display-runover.png` });
const pixels = await page.evaluate(async ({ png, where, rows, edges }) => {
  const img = new Image();
  img.src = `data:image/png;base64,${png}`;
  await img.decode();
  const cv = document.createElement('canvas');
  cv.width = img.width;
  cv.height = img.height;
  const ctx = cv.getContext('2d');
  ctx.drawImage(img, 0, 0);
  const out = [];
  for (let r = 0; r < rows; r++) {
    for (let k = 1; k < edges; k++) {
      const d = ctx.getImageData(Math.round(where.x + where.starts[k] - 1), Math.round(where.y + where.mids[r]), 1, 1).data;
      // (A gridline pixel is light grey: lighter than text, darker than the white sheet.)
      out.push([r, k, d[0] < 245 && d[0] > 150 && Math.abs(d[0] - d[1]) < 12 && Math.abs(d[1] - d[2]) < 12]);
    }
  }
  return out;
}, { png: shot.toString('base64'), where, rows: RUNOVER.rows, edges: RUNOVER.edges - 1 });
const screenWrong = pixels.filter(([r, k, line]) => line === want.has(`${r},${k}`)).map(([r, k, line]) => `row ${r + 1} edge ${k} ${line ? 'shown' : 'hidden'}`);
check(!screenWrong.length && pixels.length === RUNOVER.rows * (RUNOVER.edges - 2), `on screen: the ${want.size} gridlines Excel leaves out under running-on text are left out, the other ${pixels.length - want.size} drawn (screenshot pixels)`, screenWrong.join('; '));
// In print (Gridlines: Print is on in the file): the printed vertical lines' rows, as Excel's PDF has them.
await page.evaluate(() => window.oxe.preparePrint());
const printed = await page.evaluate(({ rows }) => {
  const pg = document.querySelector('.oxe-page');
  const block = [...pg.querySelectorAll('.oxe-page-block')].find((b) => b.querySelector('.oxe-print-grid-v'));
  const vs = [...block.querySelectorAll('.oxe-print-grid-v')].map((e) => ({ x: parseFloat(e.style.getPropertyValue('--oxe-gx')), t: parseFloat(e.style.getPropertyValue('--oxe-gt') || '0'), h: parseFloat(e.style.getPropertyValue('--oxe-gl')) }));
  const hs = [...block.querySelectorAll('.oxe-print-grid-h')].map((e) => parseFloat(e.style.getPropertyValue('--oxe-gy'))).sort((a, b) => a - b);
  const xs = [...new Set(vs.map((v) => v.x))].sort((a, b) => a - b);
  const out = [];
  for (let r = 0; r < rows; r++) {
    const mid = (hs[r] + hs[r + 1]) / 2;
    xs.forEach((x, k) => { if (!vs.some((v) => v.x === x && v.t <= mid && v.t + v.h >= mid)) out.push(`${r},${k}`); });
  }
  return out;
}, { rows: RUNOVER.rows });
const printWrong = [...new Set([...printed, ...want])].filter((x) => printed.includes(x) !== want.has(x));
check(!printWrong.length, `in print: the same ${want.size} gridlines left out as in Excel's PDF, the rest printed`, `differ: ${printWrong.join(' ')}`);

// 2. Column widths by the Normal font's maximum digit width: columns of 1, 2, 5, 10 and 20 characters, as Excel
// saved them, at Excel's pixels, and the default column, for fonts with digit widths 6 to 16.
const fonts = DIGITS.filter((f) => ['Arial 12', 'Times New Roman 10', 'Courier New 10', 'Calibri 20', 'Calibri 24', 'Verdana 12', 'Calibri 11'].includes(`${f.font} ${f.size}`));
const wrongWidths = [];
for (const f of fonts) {
  await open(`mdw-${f.font}-${f.size}.xlsx`.replace(/ /g, ''), book({ font: [f.font, f.size], cells: [['A1', 1]], cols: f.stored.map((w, k) => [k + 2, w]) }));
  const got = await page.evaluate(() => {
    const heads = [...document.querySelectorAll('.oxe-colhead')].slice(0, 6).map((h) => Math.round(h.getBoundingClientRect().width));
    return { heads, mdw: window.oxe.workbook.mdw };
  });
  const px = [f.defaultPt, ...f.pt].map((p) => Math.round(p / 0.75));
  if (got.heads.join() !== px.join()) wrongWidths.push(`${f.font} ${f.size}: ${got.heads.join(' ')} for Excel's ${px.join(' ')} (mdw ${got.mdw})`);
}
check(!wrongWidths.length && fonts.length === 7, `column headings as wide as Excel's columns for ${fonts.length} Normal fonts (mdw 6 to 16)`, wrongWidths.join('; '));
await page.screenshot({ path: `${outDir}/display-mdw.png` });
// A typeface and size Excel's table lacks (Georgia 13): the browser's measure of 0-9 at 17.33px, rounded, is used.
await open('mdw-georgia.xlsx', book({ font: ['Georgia', 13], cells: [['A1', 1]] }));
const derived = await page.evaluate(() => {
  const ctx = document.createElement('canvas').getContext('2d');
  ctx.font = `${(13 * 96) / 72}px "Georgia", sans-serif`;
  const widest = Math.max(...'0123456789'.split('').map((d) => ctx.measureText(d).width));
  return { mdw: window.oxe.workbook.mdw, canvas: Math.round(widest), head: Math.round(document.querySelector('.oxe-colhead').getBoundingClientRect().width) };
});
const pad = 2 * Math.ceil(derived.mdw / 4) + 1;
check(derived.mdw === derived.canvas && derived.head === Math.ceil((8 * derived.mdw + pad) / 8) * 8, 'a font Excel\'s table lacks: its digits measured in the browser and rounded, the default column 8 of them plus padding', JSON.stringify(derived));

// 3. East Asian built-in dates: numFmtId 28 and 34 in ja-JP, zh-CN, zh-TW, ko-KR and en-US.
await open('eastasia.xlsx', book({ cells: [['A1', 43586, 28], ['A2', 32515, 27], ['A3', 0.5625, 34], ['A4', 45356, 31]], cols: [[1, 30]] }));
const shown = async (tag) => page.evaluate((tag) => {
  window.oxe.setLocale(tag);
  return [1, 2, 3, 4].map((r) => document.querySelector(`.oxe-row[aria-rowindex="${r + 1}"] .oxe-cell[aria-colindex="2"]`)?.textContent ?? '');
}, tag);
const results = {};
for (const tag of ['ja-JP', 'zh-CN', 'zh-TW', 'ko-KR', 'en-US']) results[tag] = (await shown(tag)).join(' | ');
const expect = {
  'ja-JP': '令和1年5月1日 | S64.1.7 | 1900年1月 | 2024年3月5日',
  'zh-CN': '5月1日 | 1989年1月 | 下午1时30分 | 2024年3月5日',
  'zh-TW': '2019年5月1日 | 1989/1/7 | 下午01時30分 | 2024年3月5日',
  'ko-KR': '05-01 | 1989年 01月 07日 | 1900-01-00 | 2024년 03월 05일',
  'en-US': '5/1/2019 | 1/7/1989 | 13:30:00 | 3/5/2024',
};
const eaWrong = Object.keys(expect).filter((k) => results[k] !== expect[k]).map((k) => `${k}: ${results[k]}`);
check(!eaWrong.length, 'built-in East Asian dates shown by locale as Excel draws their codes (ja-JP era, zh 上午/下午, en-US the short date)', eaWrong.join('; '));
await page.screenshot({ path: `${outDir}/display-eastasia.png` });

check(!errors.length, 'no page errors', errors.join('\n'));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
