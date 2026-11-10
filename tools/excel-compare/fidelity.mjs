// The save-fidelity benchmark, first half: each workbook of a folder opened by the editor, one cell typed into (an empty
// cell two rows under the first sheet's data), saved, and its package compared with the original part by part: byte
// for byte the same, changed, or gone. excel-fidelity.ps1 then has Excel read both and compare what it sees.
// usage: node tools/excel-compare/fidelity.mjs <folder of .xlsx> <out folder>   (writes <out>/parts.json)
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync } from 'fflate';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';

const { DOMParser, XMLSerializer } = await import('@xmldom/xmldom');
setXmlImpl({ DOMParser, XMLSerializer });

const [src, out] = process.argv.slice(2);
if (!src || !out) { console.error('usage: node fidelity.mjs <folder of .xlsx> <out folder>'); process.exit(2); }
fs.mkdirSync(out, { recursive: true });
const same = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
const results = [];
for (const name of fs.readdirSync(src).filter((n) => /\.xlsx$/i.test(n)).sort()) {
  const bytes = new Uint8Array(fs.readFileSync(path.join(src, name)));
  const wb = Workbook.fromBytes(bytes);
  const si = wb.sheets.findIndex((s, i) => wb.canShow(i));
  const sheet = wb.sheet(si);
  const edit = { sheet: wb.sheets[si].name, r: Math.max(sheet.maxRow, 0) + 3, c: 0 };
  wb.apply(si, [{ r: edit.r, c: edit.c, value: 'oxe edit' }]);
  const saved = wb.toBytes();
  // (Its own name: Excel can't have two workbooks of one name open, to compare them.)
  fs.writeFileSync(path.join(out, `saved-${name}`), saved);
  const a = unzipSync(bytes);
  const b = unzipSync(saved);
  const parts = Object.keys(a).filter((k) => !k.endsWith('/'));
  const kept = parts.filter((k) => b[k] && same(a[k], b[k]));
  const changed = parts.filter((k) => b[k] && !same(a[k], b[k]));
  const gone = parts.filter((k) => !b[k]);
  const added = Object.keys(b).filter((k) => !k.endsWith('/') && !a[k]);
  results.push({ name, edit: `${edit.sheet}!A${edit.r + 1}`, parts: parts.length, kept: kept.length, changed, gone, added });
  console.log(`${name.padEnd(22)} parts ${String(parts.length).padStart(3)}  kept byte for byte ${String(kept.length).padStart(3)}  changed ${changed.join(', ') || '-'}${gone.length ? `  gone ${gone.join(', ')}` : ''}${added.length ? `  added ${added.join(', ')}` : ''}`);
}
const total = results.reduce((t, x) => ({ parts: t.parts + x.parts, kept: t.kept + x.kept }), { parts: 0, kept: 0 });
console.log(`all: ${total.kept} of ${total.parts} parts kept byte for byte; the rest are the sheets edited (and what an edit must touch)`);
fs.writeFileSync(path.join(out, 'parts.json'), JSON.stringify(results, null, 1));
