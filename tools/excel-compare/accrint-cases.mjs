// ACCRINT cases for measuring Excel's rules (docs/CALC.md), 4251 formulas in four sets, the same each time (fixed
// seeds):
// 1. 2000 near the ends of Excel's dates: issue, first_interest and settlement as serials near 1900-01-01 (across the
//    1900 leap-year bug, serial 60 the February 29 that wasn't) or near 9999-12-31 (serial 2958465), in every order,
//    each frequency, basis 0 to 4 and calc_method left out, TRUE or FALSE;
// 2. 1500 more the same way from another seed, held out;
// 3. 451 sweeps by one day at a time near 9999-12-31: first_interest, then settlement, then issue;
// 4. 300 in the 2000s with issue exactly on a coupon date, held out.
// Excel works them out with excel-formulas.ps1 in a new workbook; formula-results.mjs makes test/accrint-excel.json.
// usage: node tools/excel-compare/accrint-cases.mjs <formulas.json>
import fs from 'node:fs';

const LAST = 2958465;
let seed;
const rnd = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return Math.floor(seed / 65536) % n; };
// (Low dates: the first days, either side of serial 60, the first years; high: the last days and years.)
const low = () => [1, 2, 3, 28, 29, 30, 31, 32, 58, 59, 60, 61, 62, 90, 91, 120][rnd(16)] + [0, 0, 0, 91, 182, 365, 366, 400, 730, 1500][rnd(10)];
const high = () => LAST - [0, 1, 2, 5, 10, 15, 25, 30, 31, 45, 59, 60, 61, 89, 90, 91, 92, 120, 150, 181, 182, 183, 200, 270, 365, 366, 400, 730, 1500][rnd(29)];
function ends(count, s) {
  seed = s;
  const out = [];
  for (let i = 0; i < count; i++) {
    const pick = rnd(2) ? low : high;
    const d = [pick(), pick(), pick()].sort((x, y) => x - y);
    // (Mostly issue < first_interest < settlement; also first_interest before issue or after settlement.)
    const order = rnd(6);
    const [issue, first, settle] = order < 3 ? d : order < 5 ? [d[0], d[2], d[1]] : [d[1], d[0], d[2]];
    const freq = [1, 2, 4][rnd(3)];
    const basis = rnd(5);
    const method = ['', ',TRUE', ',FALSE'][rnd(3)];
    out.push(`=ACCRINT(${issue},${first},${settle},0.1,1000,${freq},${basis}${method})`);
  }
  return out;
}
const out = [...ends(2000, 4242), ...ends(1500, 777)];
for (let f = LAST - 200; f <= LAST; f++) out.push(`=ACCRINT(${LAST - 400},${f},${LAST - 1},0.1,1000,2,0,TRUE)`);
for (let s = LAST - 99; s <= LAST; s++) out.push(`=ACCRINT(${LAST - 400},${LAST - 100},${s},0.1,1000,2,0,TRUE)`);
for (let i = LAST - 400; i < LAST - 100; i += 2) out.push(`=ACCRINT(${i},${LAST - 100},${LAST - 1},0.1,1000,2,0,TRUE)`);
seed = 99;
for (let i = 0; i < 300; i++) {
  const freq = [1, 2, 4][rnd(3)];
  const y = 2005 + rnd(10); const m = 1 + rnd(12); const d = 1 + rnd(28); const k = 2 + rnd(6);
  // (Issue k coupons before first_interest, on the same day of the month.)
  const im = m - k * 12 / freq;
  const iy = y + Math.floor((im - 1) / 12); const imm = (((im - 1) % 12) + 12) % 12 + 1;
  const settle = `DATE(${y + rnd(3) - 1},${1 + rnd(12)},${1 + rnd(28)})`;
  out.push(`=ACCRINT(DATE(${iy},${imm},${d}),DATE(${y},${m},${d}),${settle},0.1,1000,${freq},${rnd(5)},${rnd(2) ? 'FALSE' : 'TRUE'})`);
}
fs.writeFileSync(process.argv[2], JSON.stringify(out));
console.log(`${out.length} formulas`);
