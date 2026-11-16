// Reduces excel-filelinks.ps1's filelinks.json to what the tests need, without the computer's paths: for each link
// given as a name relative to the workbook (or a web address), what was given, Excel's Address after the save, its
// SubAddress, and the relationship's Target and TargetMode as Excel wrote them; and the full path in the workbook's
// own folder (A2), saved as a name relative to it.
// usage: node tools/excel-compare/filelinks-extract.mjs <filelinks.json> test/filelinks-excel.json
import { readFileSync, writeFileSync } from 'node:fs';

const [from, to] = process.argv.slice(2);
const j = JSON.parse(readFileSync(from, 'utf8').replace(/^\uFEFF/, ''));
const unxml = (s) => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
const rels = Object.fromEntries([...j.rels.matchAll(/<Relationship ([^>]*)\/>/g)].map(([, a]) => {
  const at = (n) => (new RegExp(`${n}="([^"]*)"`).exec(a) || [])[1];
  return [at('Id'), { target: unxml(at('Target')), mode: at('TargetMode') || null }];
}));
const links = Object.fromEntries([...j.hyperlinks.matchAll(/<hyperlink ([^>]*)\/>/g)].map(([, a]) => {
  const at = (n) => (new RegExp(`${n}="([^"]*)"`).exec(a) || [])[1];
  return [at('ref'), { rel: rels[at('r:id')], location: at('location') ? unxml(at('location')) : null }];
}));
const out = [];
for (const [ref, c] of Object.entries(j.after)) {
  if (['A3', 'A4'].includes(ref)) continue; // (paths of the computer's own folders)
  const l = links[ref];
  out.push({ ref, given: ref === 'A2' ? '<folder>\\New Book.xlsx' : c.given, address: c.address, subAddress: c.subAddress, target: l.rel.target, mode: l.rel.mode, location: l.location });
}
writeFileSync(to, `${JSON.stringify(out, null, 1)}\n`);
