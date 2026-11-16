// Date functions around Excel's February 29, 1900 (serial 60, the leap day that wasn't), for docs/CALC.md: a fixed list
// (DATE, EDATE, EOMONTH, DAYS360, the coupon dates, YEARFRAC), 150 random DATEs with months and days out of range in
// 1900 and 1901, 250 random DAYS360 over serials 1 to 130 by both methods, and the date functions over serials 58 to
// 62. The same seed gives the same cases. Excel works them out with excel-formulas.ps1 in a new workbook;
// formula-results.mjs makes test/date1900-excel.json.
// usage: node tools/excel-compare/date1900-cases.mjs <formulas.json>
import fs from 'node:fs';

const out = ['=DATE(1900,2,29)', '=DATE(1900,2,30)', '=DATE(1900,3,0)', '=DATE(1900,1,60)', '=DATE(1900,1,61)', '=EDATE(31,1)',
  '=EDATE(29,1)', '=EDATE(30,1)', '=EOMONTH(31,0)', '=EOMONTH(40,0)', '=EOMONTH(1,1)', '=EDATE(60,1)', '=EDATE(60,12)',
  '=EDATE(91,-1)', '=EDATE(90,-1)', '=EOMONTH(91,-1)', '=EDATE(244,-6)', '=EDATE(243,-6)', '=DAYS360(59,61)',
  '=DAYS360(60,61)', '=DAYS360(60,91)', '=DAYS360(59,91)', '=DAYS360(30,60)', '=DAYS360(31,60,TRUE)',
  '=COUPPCD(100,244,2,0)', '=COUPPCD(70,244,2,0)', '=COUPNCD(50,244,2,0)', '=COUPPCD(100,243,2,0)',
  '=COUPNCD(45,243,2,0)', '=COUPDAYS(100,244,2,1)', '=COUPDAYS(50,244,2,1)', '=COUPDAYS(50,243,2,1)',
  '=YEARFRAC(59,61,1)', '=YEARFRAC(1,367,1)', '=MONTH(60)', '=DAY(61)', '=WEEKDAY(60)', '=COUPPCD(61,1000,4,0)',
  '=COUPNCD(31,1000,4,0)', '=COUPPCD(61,881,4,0)', '=COUPNCD(50,881,4,0)', '=COUPDAYS(50,881,4,0)',
  '=COUPDAYBS(50,881,4,0)'];
let seed = 5;
const r = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return Math.floor(seed / 65536) % n; };
for (let i = 0; i < 150; i++) out.push(`=DATE(${1900 + r(2)},${r(15) - 1},${r(80) - 10})`);
for (let i = 0; i < 250; i++) { const a = 1 + r(130); const b = 1 + r(130); out.push(`=DAYS360(${a},${b}${['', ',TRUE', ',FALSE'][r(3)]})`); }
for (const s of [58, 59, 60, 61, 62]) {
  out.push(`=YEAR(${s})`, `=MONTH(${s})`, `=DAY(${s})`, `=EOMONTH(${s},0)`, `=EOMONTH(${s},1)`, `=EDATE(${s},12)`, `=WEEKDAY(${s})`,
    `=TEXT(${s},"yyyy-mm-dd")`, '=DATEVALUE("1900-02-29")', `=YEARFRAC(${s},400,0)`, `=YEARFRAC(${s},400,1)`,
    `=NETWORKDAYS(${s},100)`, `=WORKDAY(${s},5)`, `=DATEDIF(${s},400,"md")`);
}
fs.writeFileSync(process.argv[2], JSON.stringify(out));
console.log(`${out.length} formulas`);
