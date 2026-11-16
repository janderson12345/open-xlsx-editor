// Excel's results from excel-refmoves.ps1 cut down to what the tests compare, with each case's step:
// test/refmoves-excel.json. (No ids, no names: the parts' revision ids and Excel's
// user name stay out.)
// usage: node tools/excel-compare/refmove-results.mjs refmoves <cases.json> <excel out.json> > test/refmoves-excel.json
import fs from 'node:fs';

const [, casesPath, outPath] = process.argv.slice(2);
const cases = JSON.parse(fs.readFileSync(casesPath, 'utf8')).cases;
const results = JSON.parse(fs.readFileSync(outPath, 'utf8').replace(/^\uFEFF/, ''));
const unescape = (s) => s.replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const list = (x) => (x == null ? [] : Array.isArray(x) ? x : [x]);

/** What the refmoves tests compare: sparklines (cell <- data, in the file's order), rules ("sqref | formula1 |
 *  formula2", sorted: Excel's own order of them shifts), notes, threads, the notes' boxes, the ovals' cells. */
function refmove(c, k) {
  const x = c.xml;
  const formulas = (body, tag) => [...body.matchAll(new RegExp(`<${tag}(\\d)>(?:<xm:f>)?([^<]*)`, 'g'))].map((m) => unescape(m[2]));
  const rules = [...(x.dataValidations || '').matchAll(/<dataValidation\b[^>]*\ssqref="([^"]*)"[^>]*>([\s\S]*?)<\/dataValidation>/g)]
    .map((m) => [m[1], ...formulas(m[2], 'formula')].join(' | ')).sort();
  const x14 = [...(x.x14DataValidations || '').matchAll(/<x14:dataValidation\b[^>]*>([\s\S]*?)<\/x14:dataValidation>/g)]
    .map((m) => [/<xm:sqref>([^<]*)/.exec(m[1])[1], ...formulas(m[1], 'x14:formula')].join(' | ')).sort();
  return {
    name: c.name, step: cases[k].step, refused: c.refused,
    sparklines: [...(x.sparklines || '').matchAll(/<xm:f>([^<]*)<\/xm:f><xm:sqref>([^<]*)<\/xm:sqref>/g)].map((m) => `${m[2]} <- ${unescape(m[1])}`),
    rules, x14,
    notes: list(x.commentRefs).sort(), threads: list(x.threadRefs).sort(), boxes: list(x.vmlNotes).sort(),
    ovals: list(c.ovals).map((o) => o.split(' ')[0]),
  };
}

const out = results.map((c, k) => refmove(c, k));
console.log(JSON.stringify(out, null, 1));
