// The fill handle's Auto Fill Options for dates: Fill Days, Fill Weekdays, Fill Months and Fill Years, as Excel fills
// (tools/excel-compare/excel-steps.ps1 on fill2-cases.json, AutoFill's xlFillDays to xlFillYears; test/fill2-excel.json):
// one date goes on by one unit a row (a day; the next weekday, Saturday and Sunday left out; a month or a year, the day
// kept or the month's last); dates the same number of units apart go on by it (two years apart: by two); dates whose
// steps differ, or don't move, are copied as a block, but for two dates the same month (year) apart, or not both
// weekdays, which each go on by one month (year, week) a round of the block. Fill Days and Weekdays drop the time.
import { dateParts } from './format.js';
import { dateSerial } from './input.js';

export const DATE_FILLS = ['days', 'weekdays', 'months', 'years'];

const weekday = (serial) => ((Math.floor(serial) % 7) + 6) % 7; // 0 Sunday ... 6 Saturday (serial 1 is a Sunday)
const isWorkday = (serial) => weekday(serial) !== 0 && weekday(serial) !== 6;
/** Serial moved by `n` weekdays (Saturdays and Sundays not counted). */
function addWorkdays(serial, n) {
  let s = Math.floor(serial);
  const step = Math.sign(n);
  for (let left = Math.abs(n); left > 0;) { s += step; if (isWorkday(s)) left--; }
  return s;
}
/** Weekdays from a to b (both weekdays), signed. */
function workdaysBetween(a, b) {
  let n = 0;
  const step = Math.sign(b - a);
  for (let s = Math.floor(a); s !== Math.floor(b); s += step) if (isWorkday(s + step)) n += step;
  return n;
}
/** Serial moved by `n` months, the day kept or the month's last, its time kept. */
function addMonths(serial, n, date1904) {
  const p = dateParts(serial, date1904);
  const total = p.y * 12 + (p.m - 1) + n;
  const y = Math.floor(total / 12);
  const m = total - y * 12 + 1;
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  const day = dateSerial(y, m, Math.min(p.d, last), date1904);
  return day == null ? null : day + (serial - Math.floor(serial));
}
const monthsBetween = (a, b, date1904) => { const p = dateParts(a, date1904); const q = dateParts(b, date1904); return (q.y - p.y) * 12 + (q.m - p.m); };

/**
 * The values `count` cells after the dates `seeds` (serials) take with Excel's Fill `unit` ('days', 'weekdays',
 * 'months', 'years'), or null when a seed isn't a number (the block then copied).
 */
export function fillDates(seeds, unit, count, date1904 = false) {
  if (!seeds.length || seeds.some((v) => typeof v !== 'number')) return null;
  const n = seeds.length;
  const whole = unit === 'days' || unit === 'weekdays';
  const months = unit === 'months' ? 1 : 12;
  const move = (v, k) => (unit === 'days' ? Math.floor(v) + k : unit === 'weekdays' ? addWorkdays(v, k) : addMonths(v, k * months, date1904));
  const out = [];
  if (n === 1) {
    for (let i = 1; i <= count; i++) out.push(move(seeds[0], i));
    return out;
  }
  // (The steps between the dates, in the unit: null where a step isn't whole.)
  const steps = seeds.slice(1).map((v, i) => {
    const a = seeds[i];
    if (unit === 'days') return Math.floor(v) - Math.floor(a);
    if (unit === 'weekdays') return isWorkday(a) && isWorkday(v) ? workdaysBetween(a, v) : null;
    const m = monthsBetween(a, v, date1904);
    return unit === 'years' ? (m % 12 === 0 ? m / 12 : null) : m;
  });
  const even = steps.every((s) => s != null && s === steps[0]) && steps[0] !== 0;
  for (let i = 0; i < count; i++) {
    const k = n + i;
    if (even) out.push(move(seeds[n - 1], steps[0] * (k - n + 1)));
    else {
      // (A block: its dates repeated, two of them going on by a unit (a week for weekdays) each round.)
      const round = Math.floor(k / n);
      const seed = seeds[k % n];
      const by = n === 2 ? round : 0;
      out.push(unit === 'weekdays' ? (by ? addWorkdays(seed, 5 * by) : Math.floor(seed)) : unit === 'days' ? Math.floor(seed) : move(seed, by));
    }
  }
  return whole ? out.map(Math.floor) : out;
}
