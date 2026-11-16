// Puts the answers of one or more excel-formulas.ps1 runs over the same formulas into one test file:
// { formulas: [...], <name>: [answer, ...] } with an answer a number, a string, a boolean, null for blank or
// { e: '#VALUE!' } for an error (a string keeps half an emoji as JSON's \ud83d).
// usage: node tools/excel-compare/formula-results.mjs <out.json> <name>=<run.json> ...
import fs from 'node:fs';

const [out, ...runs] = process.argv.slice(2);
const result = {};
for (const run of runs) {
  const [name, file] = run.split('=');
  const entries = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));
  if (!result.formulas) result.formulas = entries.map((e) => e.f);
  else if (entries.some((e, i) => e.f !== result.formulas[i])) throw new Error(`${file}: not the same formulas`);
  result[name] = entries.map((e) => (e.t === 'n' ? Number(e.v) : e.t === 's' ? String.fromCharCode(...e.u)
    : e.t === 'b' ? e.v : e.t === 'e' ? { e: e.v } : null));
}
fs.writeFileSync(out, `${JSON.stringify(result, null, 0).replace(/\],"/g, '],\n"')}\n`);
console.log(`${result.formulas.length} formulas, ${runs.length} runs`);
