// Random lists for measuring Excel's approximate lookups (docs/CALC.md): 400 unsorted lists of 1 to 12 whole
// numbers, every fourth with blanks and text among them, and a value to look up in each. The same seed gives the same
// lists. Excel answers them with excel-lookups.ps1.
// usage: node tools/excel-compare/lookup-cases.mjs <cases.json>
import fs from 'node:fs';

let seed = 12345;
const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
const cases = [];
for (let k = 0; k < 400; k++) {
  const n = 1 + Math.floor(rnd() * 12);
  const mixed = k % 4 === 3;
  const list = Array.from({ length: n }, () => (mixed && rnd() < 0.25 ? (rnd() < 0.5 ? null : 't') : Math.floor(rnd() * 10)));
  cases.push({ list, x: Math.floor(rnd() * 12) - 1 });
}
fs.writeFileSync(process.argv[2], JSON.stringify(cases));
console.log(`${cases.length} cases`);
