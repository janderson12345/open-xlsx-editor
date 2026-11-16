// Excel's number format codes (ECMA-376 Part 1, 18.8.31): sections by sign or condition, digit placeholders,
// thousands and scaling commas, percent, scientific, fractions, dates and times, text, literals, colours, padding and
// fill. Checked against Excel's own PDF of test/fixtures/numfmt.xlsx (57 formats x 14 values, test/numfmt-excel.json).
import { decimal, roundTo, formatGeneral } from './numfmt.js';
import { EN_US, localDecimal, isSpaceGroup, currentLocale } from './locale.js';
import { EAST_ASIAN_IDS, eastAsianTable } from './eastasia.js';

/** Built-in formats by id, as Excel shows them in an en-US locale (ids 5-8, 41-44 are locale currency and accounting
 *  formats; the East Asian ids 27-36 and 50-58 are builtinFormat's). 37-40 pad positives with _) as Excel does (its
 *  PDF of numfmt.xlsx leaves the width of a parenthesis after them), where ECMA-376's table shows a space. */
export const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm',
  21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)', 41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)', 43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0',
  48: '##0.0E+0', 49: '@',
};

const EAST_ASIAN = new Set(EAST_ASIAN_IDS);
const TIME_IDS = new Set([32, 33, 34, 35]);

/** The code built-in format `id` means in `locale`: an East Asian id (27-36, 50-58) is its locale's code from
 *  ECMA-376's table in Japanese, Chinese and Korean, and elsewhere what Excel 365 en-US reads it as (measured over
 *  COM: dates as format 14, m/d/yyyy, the short date; 32-35 as h:mm:ss); others are BUILTIN_FORMATS'. */
export function builtinFormat(id, locale = currentLocale()) {
  if (!EAST_ASIAN.has(Number(id))) return BUILTIN_FORMATS[id];
  const table = eastAsianTable(locale?.tag);
  if (table) return table[id];
  return TIME_IDS.has(Number(id)) ? 'h:mm:ss' : 'm/d/yyyy';
}

// Japanese eras by first day (Windows' calendar, as Excel's [$-411] codes show them: 1989-01-07 is Showa 64,
// 1989-01-08 Heisei 1), with g, gg and ggg names.
const ERAS = [
  [20190501, 2018, 'R', '令', '令和'], [19890108, 1988, 'H', '平', '平成'],
  [19261225, 1925, 'S', '昭', '昭和'], [19120730, 1911, 'T', '大', '大正'],
  [0, 1867, 'M', '明', '明治'],
];
const AMPM_ZH = '上午/下午';

const COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00' };
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Split a code into sections at unquoted, unescaped semicolons, and each section into tokens:
 *  { t: 'lit', s } literal text, { t: 'pad', s } (_x: space as wide as x), { t: 'fill', s } (*x: x repeated to fill),
 *  { t: 'd', s } a digit placeholder 0 # ?, { t: '.' }, { t: ',' }, { t: '%' }, { t: 'E', s: '+' | '-' },
 *  { t: '/' }, { t: 'date', s } (y.. m.. d.. h.. s.. AM/PM A/P, elapsed [h] [m] [s], .0 fractions), { t: '@' },
 *  { t: 'general' }; with the section's colour and condition. */
function parseCode(code) {
  const sections = [];
  let cur = { tokens: [], color: null, cond: null, ja: false, gannen: false };
  const lit = (s) => {
    const last = cur.tokens[cur.tokens.length - 1];
    if (last && last.t === 'lit') last.s += s;
    else cur.tokens.push({ t: 'lit', s });
  };
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    const rest = code.slice(i);
    if (ch === ';') { sections.push(cur); cur = { tokens: [], color: null, cond: null, ja: false, gannen: false }; continue; }
    if (ch === '"') { const j = code.indexOf('"', i + 1); lit(code.slice(i + 1, j < 0 ? code.length : j)); i = j < 0 ? code.length : j; continue; }
    if (ch === '\\') { lit(code[i + 1] ?? ''); i++; continue; }
    if (ch === '_') { cur.tokens.push({ t: 'pad', s: code[i + 1] ?? ' ' }); i++; continue; }
    if (ch === '*') { cur.tokens.push({ t: 'fill', s: code[i + 1] ?? ' ' }); i++; continue; }
    if (ch === '[') {
      const j = code.indexOf(']', i);
      const body = code.slice(i + 1, j < 0 ? code.length : j);
      i = j < 0 ? code.length : j;
      const low = body.toLowerCase();
      if (COLORS[low]) cur.color = COLORS[low];
      else if (/^color\s*\d+$/.test(low)) cur.color = { indexed: Number(low.replace(/\D/g, '')) + 7 };
      else if (/^(<=|>=|<>|<|>|=)/.test(body)) { const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)/.exec(body); if (m) cur.cond = { op: m[1], v: Number(m[2]) }; }
      else if (/^(h+|m+|s+)$/.test(low)) cur.tokens.push({ t: 'date', s: `[${low}]` });
      else if (body.startsWith('$')) {
        const sym = body.slice(1).split('-')[0];
        if (sym) lit(sym);
        // A locale after the dash: Japanese ([$-411], [$-ja-JP], a calendar byte before 411) counts eras.
        const loc = body.slice(body.indexOf('-') + 1).toLowerCase();
        if (body.includes('-') && (/^ja(-|,|$)/.test(loc) || /^[0-9a-f]*0411(,|$)|^411(,|$)/.test(loc))) cur.ja = true;
        if (/-x-gannen/.test(loc)) cur.gannen = true;
      }
      continue;
    }
    if (/^general/i.test(rest)) { cur.tokens.push({ t: 'general' }); i += 6; continue; }
    if (rest.startsWith(AMPM_ZH)) { cur.tokens.push({ t: 'date', s: AMPM_ZH }); i += AMPM_ZH.length - 1; continue; }
    if (/^(am\/pm|a\/p)/i.test(rest)) { const n = /^am\/pm/i.test(rest) ? 5 : 3; cur.tokens.push({ t: 'date', s: rest.slice(0, n).toUpperCase() === 'AM/PM' ? 'AM/PM' : 'A/P' }); i += n - 1; continue; }
    if (/[yYmMdDhHsSgG]/.test(ch) || ((ch === 'e' || ch === 'E') && code[i + 1] !== '+' && code[i + 1] !== '-')) {
      let j = i;
      while (j < code.length && code[j].toLowerCase() === ch.toLowerCase()) j++;
      cur.tokens.push({ t: 'date', s: code.slice(i, j).toLowerCase() });
      i = j - 1;
      continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') { cur.tokens.push({ t: 'd', s: ch }); continue; }
    if (ch === '.') { cur.tokens.push({ t: '.' }); continue; }
    if (ch === ',') { cur.tokens.push({ t: ',' }); continue; }
    if (ch === '%') { cur.tokens.push({ t: '%' }); continue; }
    if ((ch === 'E' || ch === 'e') && (code[i + 1] === '+' || code[i + 1] === '-')) { cur.tokens.push({ t: 'E', s: code[i + 1] }); i++; continue; }
    if (ch === '/') { cur.tokens.push({ t: '/' }); continue; }
    if (ch === '@') { cur.tokens.push({ t: '@' }); continue; }
    if (/[1-9]/.test(ch) && cur.tokens.some((t) => t.t === '/')) { cur.tokens.push({ t: 'den', s: ch }); continue; }
    lit(ch);
  }
  sections.push(cur);
  for (const sec of sections) {
    // A .0 after seconds is fractions of a second; an "m" after hours or before seconds is minutes.
    const toks = sec.tokens;
    for (let k = 0; k < toks.length; k++) {
      if (toks[k].t === '.' && toks.slice(0, k).some((t) => t.t === 'date' && /^\[?s/.test(t.s))) {
        let n = 0;
        while (toks[k + 1 + n] && toks[k + 1 + n].t === 'd' && toks[k + 1 + n].s === '0') n++;
        if (n) { toks.splice(k, 1 + n, { t: 'date', s: '.' + '0'.repeat(n) }); }
      }
    }
    const dates = toks.filter((t) => t.t === 'date');
    dates.forEach((t, k) => {
      if (!/^m{1,2}$/.test(t.s)) return;
      const prev = dates.slice(0, k).reverse().find((x) => !/^(AM\/PM|A\/P|\.0+)$/.test(x.s));
      const next = dates.slice(k + 1).find((x) => !/^(AM\/PM|A\/P)$/.test(x.s));
      if ((prev && /^\[?h/.test(prev.s)) || (next && /^\[?s/.test(next.s))) t.s = t.s === 'm' ? 'M' : 'MM'; // minutes
    });
    sec.isDate = dates.length > 0;
    sec.isText = toks.some((t) => t.t === '@');
  }
  return sections;
}

const cache = new Map();
/** A format code parsed once. */
export function compileFormat(code) {
  if (!cache.has(code)) cache.set(code, parseCode(code));
  return cache.get(code);
}

/** Whether a format code shows dates or times (for editing and alignment later). */
export const isDateFormat = (code) => compileFormat(code).some((s) => s.isDate);

const test = (c, v) => ({ '<': v < c.v, '>': v > c.v, '=': v === c.v, '<=': v <= c.v, '>=': v >= c.v, '<>': v !== c.v })[c.op];

/** The section a number takes, and whether its sign is shown by the section rather than a minus. */
function pickSection(sections, v) {
  const nums = sections.filter((s, k) => k < 3);
  if (nums.some((s) => s.cond)) {
    // Explicit conditions: the first section whose condition holds; the next one without a condition otherwise.
    for (const s of nums) if (s.cond && test(s.cond, v)) return { sec: s, abs: s !== nums[0] && v < 0 };
    const rest = nums.find((s) => !s.cond);
    return { sec: rest || nums[nums.length - 1], abs: !!rest && v < 0 && rest !== nums[0] };
  }
  if (v < 0 && nums.length >= 2) return { sec: nums[1], abs: true };
  if (v === 0 && nums.length >= 3) return { sec: nums[2], abs: true };
  return { sec: nums[0], abs: false };
}

/** Digits of |a| rounded to `decimals` places, half up on Excel's 15 digits: { int: '1234', frac: '57' }. */
function digitsOf(a, decimals) {
  if (a === 0) return { int: '0', frac: '0'.repeat(decimals) };
  const v = decimal(a);
  const keep = v.exp + 1 + decimals;
  if (keep <= 0) {
    // Rounds to 0, or up to one unit of the last place.
    const up = keep === 0 && v.digits.charCodeAt(0) >= 53;
    if (up && !decimals) return { int: '1', frac: '' };
    return { int: '0', frac: up ? '0'.repeat(decimals - 1) + '1' : '0'.repeat(decimals) };
  }
  const { digits, carried } = roundTo(v.digits, keep);
  const exp = v.exp + (carried ? 1 : 0);
  const all = digits.padEnd(exp + 1 + decimals, '0');
  const intLen = exp + 1;
  return intLen > 0 ? { int: all.slice(0, intLen).replace(/^0+(?=\d)/, '') || '0', frac: all.slice(intLen, intLen + decimals) } : { int: '0', frac: ('0'.repeat(-intLen) + all).slice(0, decimals) };
}

/** Place integer digits into a run of placeholders and literals, right to left; extra digits go before the first
 *  placeholder. `zeroShows`: a lone zero integer shows in '#' places too (as in an exponent's mantissa of 0). */
function fillInt(tokens, int, grouping, zeroShows) {
  let digits = int === '0' && !zeroShows ? '' : int;
  const places = tokens.filter((t) => t.t === 'd').length;
  const minDigits = tokens.filter((t) => t.t === 'd' && t.s === '0').length;
  if (zeroShows) digits = digits.padStart(places, '0');
  if (grouping && digits.length) {
    // Separators go between groups of the digits shown, zero padding included.
    const padded = digits.padStart(Math.max(digits.length, minDigits), '0');
    digits = padded.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    if (grouping !== true && grouping !== ',') digits = digits.replace(/,/g, grouping);
  }
  const out = [];
  let k = digits.length;
  let placesLeft = places;
  for (let i = tokens.length - 1; i >= 0; i--) {
    const t = tokens[i];
    if (t.t !== 'd') { out.unshift(t.t === 'lit' ? t.s : t.t === 'pad' ? { pad: t.s } : t.t === 'fill' ? { fill: t.s } : ''); continue; }
    placesLeft--;
    let s;
    if (k > 0) {
      // Commas from grouping ride with the digit before them.
      let take = 1;
      while (k - take > 0 && /\D/.test(digits[k - take])) take++;
      s = digits.slice(k - take, k);
      if (placesLeft === 0) s = digits.slice(0, k);
      k -= placesLeft === 0 ? k : take;
    } else s = t.s === '0' ? '0' : t.s === '?' ? { pad: '0' } : '';
    out.unshift(s);
  }
  return out;
}

/** Place decimal digits left to right; trailing zeros in '#' places go, in '?' places become spaces. */
function fillFrac(tokens, frac) {
  const places = tokens.filter((t) => t.t === 'd');
  let shown = frac.split('');
  for (let i = shown.length - 1; i >= 0 && shown[i] === '0' && places[i] && places[i].s !== '0'; i--) shown[i] = places[i].s === '?' ? { pad: '0' } : '';
  let k = 0;
  return tokens.map((t) => (t.t === 'd' ? shown[k++] ?? '' : t.t === 'lit' ? t.s : t.t === 'pad' ? { pad: t.s } : t.t === 'fill' ? { fill: t.s } : ''));
}

/** The closest fraction to x (0 <= x < 1) with a denominator of at most `maxDen` (continued fractions). */
function bestFraction(x, maxDen) {
  let [h0, h1, k0, k1] = [0, 1, 1, 0];
  let v = x;
  let best = [0, 1];
  for (let n = 0; n < 64; n++) {
    const a = Math.floor(v);
    const h2 = a * h1 + h0;
    const k2 = a * k1 + k0;
    if (k2 > maxDen) {
      // The best semiconvergent within the limit, if it beats the last convergent.
      const t = Math.floor((maxDen - k0) / k1);
      const cand = [t * h1 + h0, t * k1 + k0];
      if (Math.abs(cand[0] / cand[1] - x) < Math.abs(best[0] / best[1] - x)) best = cand;
      break;
    }
    [h0, h1, k0, k1] = [h1, h2, k1, k2];
    best = [h1, k1];
    if (v - a < 1e-12) break;
    v = 1 / (v - a);
  }
  return best;
}

/** Format the numeric part of a section (digit placeholders, literals, percent, scaling, exponent, fraction). */
function formatNumber(tokens, a, loc = EN_US) {
  let v = a;
  // Percent multiplies; commas right after the last digit placeholder (before the point or at the end) divide by 1000.
  v *= 100 ** tokens.filter((t) => t.t === '%').length;
  const lastDigit = tokens.reduce((k, t, i) => (t.t === 'd' ? i : k), -1);
  const point = tokens.findIndex((t) => t.t === '.');
  let scaling = 0;
  const intEnd = point >= 0 ? point : lastDigit + 1;
  if (point >= 0) for (let i = point - 1; i >= 0 && tokens[i].t === ','; i--) scaling++;
  for (let i = lastDigit + 1; i < tokens.length && tokens[i].t === ','; i++) scaling++;
  v /= 1000 ** scaling;
  const grouping = tokens.some((t, i) => t.t === ',' && i < intEnd && tokens.slice(0, i).some((x) => x.t === 'd') && tokens.slice(i + 1, intEnd).some((x) => x.t === 'd'));
  const clean = tokens.filter((t) => t.t !== ',' && t.t !== '%' || t.t === '%');
  const E = clean.findIndex((t) => t.t === 'E');
  const slash = clean.findIndex((t) => t.t === '/');
  const pct = (list) => list.map((t) => (t.t === '%' ? { t: 'lit', s: '%' } : t)).filter((t) => t.t !== ',');
  if (E >= 0) return formatScientific(pct(clean), E, v, loc);
  if (slash >= 0) return formatFraction(pct(clean), slash, v);
  const list = pct(clean);
  const p = list.findIndex((t) => t.t === '.');
  const intToks = p >= 0 ? list.slice(0, p) : list;
  const fracToks = p >= 0 ? list.slice(p + 1) : [];
  const decimals = fracToks.filter((t) => t.t === 'd').length;
  const { int, frac } = digitsOf(v, decimals);
  // (A no-break space as the thousands separator is shown as a plain space: Excel's Range.Text with it set, measured.)
  const group = isSpaceGroup(loc.group) ? ' ' : loc.group;
  return [...fillInt(intToks, int, grouping && group, false), ...(p >= 0 ? [loc.decimal] : []), ...fillFrac(fracToks, frac)];
}

function formatScientific(list, E, v, loc) {
  const mant = list.slice(0, E);
  const expToks = list.slice(E + 1);
  const p = mant.findIndex((t) => t.t === '.');
  const intToks = p >= 0 ? mant.slice(0, p) : mant;
  const fracToks = p >= 0 ? mant.slice(p + 1) : [];
  const intPlaces = intToks.filter((t) => t.t === 'd').length;
  const decimals = fracToks.filter((t) => t.t === 'd').length;
  // Engineering notation when there are several integer places: exponents in steps of that many.
  const step = intPlaces > 1 ? intPlaces : 1;
  let exp = v === 0 ? 0 : Math.floor(Math.log10(v));
  exp = Math.floor(exp / step) * step;
  let m = v === 0 ? 0 : v / 10 ** exp;
  let { int, frac } = digitsOf(m, decimals);
  if (int.length > step) {
    exp += step;
    m = v / 10 ** exp;
    ({ int, frac } = digitsOf(m, decimals));
  }
  const expDigits = expToks.filter((t) => t.t === 'd').length;
  const sign = exp < 0 ? '-' : list[E].s === '+' ? '+' : '';
  // (Excel shows a zero mantissa with every integer place filled: 0 in ##0.0E+0 is 000.0E+0.)
  return [...fillInt(intToks, int, false, v === 0), ...(p >= 0 ? [loc.decimal] : []), ...fillFrac(fracToks, frac), 'E', sign, String(Math.abs(exp)).padStart(expDigits, '0'),
    ...expToks.filter((t) => t.t !== 'd').map((t) => (t.t === 'lit' ? t.s : ''))];
}

function formatFraction(list, slash, v) {
  // [integer part, space] numerator / denominator: the integer part is the digit run before the first space-literal.
  const numStart = (() => {
    let i = slash - 1;
    while (i >= 0 && list[i].t === 'd') i--;
    return i + 1;
  })();
  const head = list.slice(0, numStart);
  const hasInt = head.some((t) => t.t === 'd');
  const numToks = list.slice(numStart, slash);
  const after = list.slice(slash + 1);
  const denToks = after.filter((t) => t.t === 'd' || t.t === 'den');
  const tail = after.filter((t) => t.t !== 'd' && t.t !== 'den');
  const fixed = denToks.some((t) => t.t === 'den') ? Number(denToks.map((t) => t.s).join('')) : null;
  const maxDen = fixed || 10 ** denToks.length - 1;
  let int = hasInt ? Math.floor(v) : 0;
  let x = v - int;
  let [n, d] = fixed ? [Math.round(x * fixed), fixed] : bestFraction(x, maxDen);
  if (n === d) { int += 1; n = 0; }
  if (!hasInt) n += int * d;
  const out = [];
  const headLits = head.filter((t) => t.t !== 'd').map((t) => (t.t === 'lit' ? t.s : ''));
  if (hasInt) {
    const intPart = fillInt(head.filter((t) => t.t === 'd'), String(int), false, false).map((s) => (s === '' && !n ? '0' : s));
    // With no fraction part the integer keeps its place: the fraction's room stays blank, as wide as the space before
    // it, its digits and its slash (Excel's PDF of numfmt.xlsx: 0 in # ?/? leaves 24.25px, # ??/?? 39.25px).
    if (!n) out.push(...intPart, ...headLits.filter(Boolean).map((s) => ({ pad: s })), ...numToks.map(() => ({ pad: '0' })), { pad: '/' }, ...denToks.map(() => ({ pad: '0' })));
    else out.push(...(int ? intPart : []), ...headLits);
  }
  // (Digits the numerator and denominator don't use stay blank, a digit wide: 1/2 in ??/?? leaves 7.65px after it.)
  if (n) out.push(...Array(Math.max(0, numToks.length - String(n).length)).fill({ pad: '0' }), String(n), '/', String(d), ...Array(Math.max(0, denToks.length - String(d).length)).fill({ pad: '0' }));
  else if (!hasInt) out.push('0');
  out.push(...tail.map((t) => (t.t === 'lit' ? t.s : '')));
  return out;
}

/** A serial date-time's parts: 1900 system (serial 0 is "January 0, 1900", 60 the 29 February 1900 Excel keeps
 *  from Lotus 1-2-3) or 1904. Time rounded to whole seconds, or to `fracDigits` decimals of a second. */
export function dateParts(serial, date1904, fracDigits = 0) {
  const unit = 10 ** fracDigits;
  let total = Math.round(serial * 86400 * unit);
  let day = Math.floor(total / (86400 * unit));
  let secs = (total - day * 86400 * unit) / unit;
  let y; let m; let d;
  if (date1904) {
    const t = new Date(Date.UTC(1904, 0, 1 + day));
    [y, m, d] = [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()];
  } else if (day === 0) [y, m, d] = [1900, 1, 0];
  else if (day === 60) [y, m, d] = [1900, 2, 29];
  else {
    const t = new Date(Date.UTC(1899, 11, 31 + (day > 60 ? day - 1 : day)));
    [y, m, d] = [t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()];
  }
  const weekday = date1904 ? (day + 5) % 7 : (day + 6) % 7;
  const whole = Math.floor(secs);
  return { y, m, d, weekday, H: Math.floor(whole / 3600), M: Math.floor((whole % 3600) / 60), S: whole % 60, frac: secs - whole, total: total / unit, day };
}

function formatDate(tokens, v, date1904, loc = EN_US, ja = { ja: false, gannen: false }) {
  const fracTok = tokens.find((t) => t.t === 'date' && t.s.startsWith('.'));
  const p = dateParts(v, date1904, fracTok ? fracTok.s.length - 1 : 0);
  const ampm = tokens.some((t) => t.t === 'date' && (t.s === 'AM/PM' || t.s === 'A/P' || t.s === AMPM_ZH));
  // Japanese era and year in it (e is the plain year outside a Japanese code).
  const ymd = p.y * 10000 + p.m * 100 + p.d;
  const era = ERAS.find((x) => ymd >= x[0]);
  const eraYear = (n) => (ja.gannen && p.y - era[1] === 1 ? '元' : String(p.y - era[1]).padStart(n, '0'));
  const h12 = p.H % 12 === 0 ? 12 : p.H % 12;
  const two = (n) => String(n).padStart(2, '0');
  return tokens.map((t) => {
    if (t.t === 'lit') return t.s;
    if (t.t === 'pad') return { pad: t.s };
    if (t.t === 'fill') return { fill: t.s };
    if (t.t === 'd') return t.s === '0' ? '0' : '';
    if (t.t !== 'date') return { '.': '.', '/': '/', '%': '%', ',': ',' }[t.t] ?? '';
    const s = t.s;
    switch (s) {
      case 'y': case 'yy': return two(p.y % 100);
      case 'yyy': case 'yyyy': return String(p.y);
      case 'm': return String(p.m);
      case 'mm': return two(p.m);
      case 'mmm': return MONTHS[p.m - 1].slice(0, 3);
      case 'mmmmm': return MONTHS[p.m - 1][0];
      case 'd': return String(p.d);
      case 'dd': return two(p.d);
      case 'ddd': return DAYS[p.weekday].slice(0, 3);
      case 'h': return String(ampm ? h12 : p.H);
      case 'hh': return two(ampm ? h12 : p.H);
      case 'M': return String(p.M);
      case 'MM': return two(p.M);
      case 's': return String(p.S);
      case 'ss': return two(p.S);
      case 'AM/PM': return p.H < 12 ? 'AM' : 'PM';
      case 'A/P': return p.H < 12 ? 'A' : 'P';
      case AMPM_ZH: return AMPM_ZH.split('/')[p.H < 12 ? 0 : 1];
      case 'g': return ja.ja ? era[2] : '';
      case 'gg': return ja.ja ? era[3] : '';
      case 'e': return ja.ja ? eraYear(1) : String(p.y);
      case 'ee': return ja.ja ? eraYear(2) : String(p.y);
      case '[h]': case '[hh]': return String(Math.floor(p.total / 3600)).padStart(s.length - 2, '0');
      case '[m]': case '[mm]': return String(Math.floor(p.total / 60)).padStart(s.length - 2, '0');
      case '[s]': case '[ss]': return String(Math.floor(p.total)).padStart(s.length - 2, '0');
      default:
        if (s.startsWith('.')) return loc.decimal + String(Math.round(p.frac * 10 ** (s.length - 1))).padStart(s.length - 1, '0');
        if (/^m{4,}$/.test(s)) return MONTHS[p.m - 1];
        if (/^d{4,}$/.test(s)) return DAYS[p.weekday];
        if (/^y+$/.test(s)) return String(p.y);
        if (/^h+$/.test(s)) return two(ampm ? h12 : p.H);
        if (/^s+$/.test(s)) return two(p.S);
        if (/^g+$/.test(s)) return ja.ja ? era[4] : '';
        if (/^e+$/.test(s)) return ja.ja ? eraYear(2) : String(p.y);
        return '';
    }
  });
}

/** Join formatted pieces: text, with padding and fill kept apart for the view ({ pad } is a space the width of a
 *  character, { fill } a character repeated to fill the cell). */
function finish(pieces, color) {
  const parts = [];
  for (const piece of pieces.flat()) {
    if (piece === '' || piece == null) continue;
    if (typeof piece === 'string') {
      if (parts.length && typeof parts[parts.length - 1] === 'string') parts[parts.length - 1] += piece;
      else parts.push(piece);
    } else parts.push(piece);
  }
  const text = parts.map((x) => (typeof x === 'string' ? x : x.pad ? ' ' : '')).join('');
  return { text, parts, color };
}

/** A value in a format code: { text, parts, color, hashes }. `text` with padding as spaces and fill left out; `parts`
 *  for the view; `color` from the section ('#RRGGBB' or { indexed }); `hashes` when the cell must show #s (a negative
 *  date). `width` is the cell's room in characters, for General inside a code. `locale` (locale.js) gives the decimal
 *  and thousands separators the number is drawn with (the code's own are en-US's, as in the file); en-US unless given,
 *  as TEXT() works. */
export function formatValue(v, code, { date1904 = false, width = Infinity, locale = EN_US } = {}) {
  const sections = compileFormat(code || 'General');
  if (typeof v === 'string') {
    const sec = sections[3] || (sections.length === 1 && sections[0].isText ? sections[0] : null);
    if (!sec) return { text: v, parts: [v], color: null };
    return finish(sec.tokens.map((t) => (t.t === '@' ? v : t.t === 'lit' ? t.s : t.t === 'pad' ? { pad: t.s } : t.t === 'fill' ? { fill: t.s } : '')), sec.color);
  }
  if (typeof v !== 'number' || !Number.isFinite(v)) return { text: String(v), parts: [String(v)], color: null };
  const { sec, abs } = pickSection(sections, v);
  const neg = v < 0 && !abs;
  const a = Math.abs(v);
  const toks = sec.tokens;
  if (sec.isDate) {
    if (v < 0) return { text: '', parts: [], color: sec.color, hashes: true };
    return finish(formatDate(toks, a, date1904, locale, sec), sec.color);
  }
  if (toks.some((t) => t.t === 'general') || (sec.isText && !toks.some((t) => t.t === 'd'))) {
    // General inside a code, or a text format (@) given a number: the number as General, literals around it.
    const g = localDecimal(formatGeneral(a, width - (neg ? 1 : 0)), locale);
    const body = toks.some((t) => t.t === 'general') ? toks.map((t) => (t.t === 'general' ? g : t.t === 'lit' ? t.s : t.t === 'pad' ? { pad: t.s } : t.t === 'fill' ? { fill: t.s } : '')) : [g];
    return finish([neg ? '-' : '', ...body], sec.color);
  }
  if (!toks.some((t) => t.t === 'd')) {
    // Only literals ("big", or an empty section: ;;; hides the value).
    return finish([neg ? '-' : '', ...toks.map((t) => (t.t === 'lit' ? t.s : t.t === 'pad' ? { pad: t.s } : t.t === 'fill' ? { fill: t.s } : ''))], sec.color);
  }
  const pieces = formatNumber(toks, a, locale);
  // A minus goes first, before any literal ("-Total: 1", "-$1.00"), and not at all when the digits shown are all zero
  // (-1 in #,##0.0,,"M" is 0.0M; General keeps it: -0.12 in two characters is -0).
  const shown = pieces.flat().filter((x) => typeof x === 'string').join('');
  return finish([neg && /[1-9]/.test(shown) ? '-' : '', ...pieces], sec.color);
}
