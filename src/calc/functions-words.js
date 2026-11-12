// ENCODEURL, BAHTTEXT and PHONETIC as Excel 365 (English) gives them (measured over COM, docs/CALC.md).
import { Area, E } from './values.js';
import { def, lifted, num, text, roundDec } from './functions.js';

// ---- ENCODEURL: UTF-8, every byte but a letter, a digit, - _ and . as %XX (~ too, unlike encodeURIComponent).
lifted('ENCODEURL', (a, ctx) => {
  let out = '';
  for (const b of new TextEncoder().encode(text(a[0], ctx))) {
    out += /[A-Za-z0-9\-_.]/.test(String.fromCharCode(b)) && b < 0x80 ? String.fromCharCode(b) : `%${b.toString(16).toUpperCase().padStart(2, '0')}`;
  }
  return out;
});

// ---- BAHTTEXT: a number in Thai words, baht and satang.
const DIGITS = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
const PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
const MILLION = 'ล้าน';

/** Up to six digits in words; `more`: a higher digit comes before them, so a 1 in the units place is เอ็ด. */
function group(n, more) {
  const ds = String(n).split('').map(Number).reverse();
  let out = '';
  for (let p = ds.length - 1; p >= 0; p--) {
    const d = ds[p];
    if (!d) continue;
    if (p === 1) out += (d === 1 ? '' : d === 2 ? 'ยี่' : DIGITS[d]) + PLACES[1];
    else if (p === 0) out += d === 1 && (more || ds.length > 1) ? 'เอ็ด' : DIGITS[d];
    else out += DIGITS[d] + PLACES[p];
  }
  return out;
}

/** A whole number in words, by groups of a million (ล้าน after each but the last). */
function words(n) {
  const s = BigInt(n).toString();
  const groups = [];
  for (let end = s.length; end > 0; end -= 6) groups.unshift(Number(s.slice(Math.max(0, end - 6), end)));
  let out = '';
  groups.forEach((g, i) => {
    const higher = groups.slice(0, i).some(Boolean);
    if (g) out += group(g, higher);
    if (i < groups.length - 1 && (g || out)) out += MILLION;
  });
  return out;
}

lifted('BAHTTEXT', (a, ctx) => {
  const v = roundDec(num(a[0], ctx), 2);
  const neg = v < 0;
  const abs = Math.abs(v);
  const baht = Math.floor(abs);
  const satang = Math.round((abs - baht) * 100);
  let out = neg ? 'ลบ' : '';
  if (baht) out += `${words(baht)}บาท`;
  if (satang) out += `${group(satang, false)}สตางค์`;
  else out += baht ? 'ถ้วน' : 'ศูนย์บาทถ้วน';
  return out;
});

// ---- PHONETIC: a cell's phonetic reading (furigana). Excel gives #N/A without an East Asian editing language, even for
// a cell whose text has one saved (measured), so #N/A here.
def('PHONETIC', (a) => { if (!(a[0] instanceof Area)) throw E.value; throw E.na; });
