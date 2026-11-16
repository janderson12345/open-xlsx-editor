// What Excel wrote for links to another workbook (dst.xlsx from excel-extlinks.ps1, kept outside the repo: its parts
// hold the computer's paths), reduced to what the tests compare and nothing of the computer's: each cell's formula
// as the file has it, its value and type; the defined names; the link part's sheet names, defined names and the
// values it keeps, sheet by sheet; the link's relative file name.
// usage: node tools/excel-compare/extlinks-extract.mjs <dst.xlsx> test/extlinks-excel.json
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';

const [from, to] = process.argv.slice(2);
const files = unzipSync(new Uint8Array(fs.readFileSync(from)));
const text = (p) => strFromU8(files[p]);
const all = (re, s) => [...s.matchAll(re)];
const unesc = (s) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const sheet = text('xl/worksheets/sheet1.xml');
const cells = {};
for (const m of all(/<c r="([A-Z]+\d+)"([^>]*)>(?:<f>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?<\/c>/g, sheet)) {
  const t = /t="(\w+)"/.exec(m[2])?.[1] || 'n';
  cells[m[1]] = { formula: m[3] != null ? unesc(m[3]) : null, value: m[4] != null ? unesc(m[4]) : null, t };
}
const book = text('xl/workbook.xml');
const names = Object.fromEntries(all(/<definedName name="([^"]+)"([^>]*)>([^<]*)<\/definedName>/g, book).map((m) => [m[1] + (/localSheetId="(\d+)"/.exec(m[2]) ? `@${/localSheetId="(\d+)"/.exec(m[2])[1]}` : ''), unesc(m[3])]));
const part = text('xl/externalLinks/externalLink1.xml');
const link = {
  sheetNames: all(/<sheetName val="([^"]*)"\/>/g, part).map((m) => unesc(m[1])),
  names: all(/<definedName name="([^"]*)" refersTo="([^"]*)"\/>/g, part).map((m) => ({ name: m[1], refersTo: unesc(m[2]) })),
  sheetData: all(/<sheetData sheetId="(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/sheetData>)/g, part).map((m) => ({
    sheetId: Number(m[1]), refreshError: /refreshError="1"/.test(m[2]),
    cells: all(/<cell r="([A-Z]+\d+)"(?: t="(\w+)")?><v>([^<]*)<\/v><\/cell>/g, m[3] || '').map((c) => ({ r: c[1], t: c[2] || 'n', v: unesc(c[3]) })),
  })),
  target: /Id="rId1"[^>]*Target="([^"]*)"/.exec(text('xl/externalLinks/_rels/externalLink1.xml.rels'))?.[1] || null,
};
fs.writeFileSync(to, `${JSON.stringify({ cells, names, link }, null, 2)}\n`);
console.log(`${Object.keys(cells).length} cells, ${Object.keys(names).length} names`);
