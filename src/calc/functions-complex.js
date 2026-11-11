// The complex-number functions: COMPLEX, IMREAL, IMAGINARY, IMABS, IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT,
// IMDIV, IMPOWER, IMSQRT, IMEXP, IMLN, IMLOG10, IMLOG2, and the circular and hyperbolic ones (IMSIN, IMCOS, IMTAN,
// IMCOT, IMSEC, IMCSC, IMSINH, IMCOSH, IMSECH, IMCSCH). A complex number is text, "3+4i" or "3+4j". As Excel reads and
// writes them, measured against Excel with the formula probe (test/formulas-excel.json) and Application.Evaluate.
// Registered into functions.js's table; evaluate.js imports this file.
import { Area, Arr, Union, E, XlError, MISSING, textToNumber, complexPartText } from './values.js';
import { def, lifted, given, one, cellsIn } from './functions.js';
import { sin as xsin, cos as xcos } from './trig.js';

/** Excel's smallest normal number: text reading as less (1E-400) is #NUM!, as is text reading as more than the largest. */
const TINY = 2.2250738585072014e-308;
/** Excel's circular functions give #NUM! from 2^27 on (=SIN(134217728)). */
const TRIG_LIMIT = 2 ** 27;

/** A part of a complex number as text: a number as Excel reads one (" 4", "$2", "(2)", "2%", "1,000", "1e3"). */
function part(s) {
  const n = textToNumber(s);
  if (n == null || (n !== 0 && Math.abs(n) < TINY) || (n === 0 && /[1-9]/.test(s.split(/[eE]/)[0]))) throw E.num;
  return n + 0;
}
/** The coefficient of i: nothing or a sign alone is 1 or -1. */
const coefficient = (s) => (s === '' || s === '+' ? 1 : s === '-' ? -1 : part(s));

/** Text as a complex number { re, im, unit }: the real part, then the imaginary one from the first sign that isn't the
 *  first character or an exponent's ("3+-4i" is 3-4i, "++i" #NUM!); a space may come before that sign, not after it
 *  alone ("3 +i", but not "3+ i"). `unit` is "i" or "j" when the text ends with one, else "". */
function parse(text) {
  const s = text.trim();
  if (s === '') throw E.num;
  const last = s[s.length - 1];
  const unit = last === 'i' || last === 'j' ? last : '';
  const body = unit ? s.slice(0, -1) : s;
  let k = -1;
  for (let i = 1; i < body.length; i++) if ((body[i] === '+' || body[i] === '-') && !/[eE]/.test(body[i - 1])) { k = i; break; }
  if (!unit) {
    if (k >= 0) throw E.num;
    return { re: part(body), im: 0, unit };
  }
  if (k < 0) return { re: 0, im: coefficient(body), unit };
  const rest = body.slice(k + 1);
  const im = rest === '' ? 1 : part(rest);
  return { re: part(body.slice(0, k)), im: (body[k] === '-' ? -im : im) + 0, unit };
}

/** A value as a complex number: a number is its real part, an empty cell 0; TRUE and FALSE are #VALUE!. */
function complex(v) {
  if (v instanceof XlError) throw v;
  if (typeof v === 'number') return { re: v, im: 0, unit: '' };
  if (v == null || v === MISSING) return { re: 0, im: 0, unit: '' };
  if (typeof v === 'string') return parse(v);
  throw E.value;
}
const arg = (v, ctx) => complex(one(v, ctx));

/** A complex number as Excel writes one: "3+4i", "-i", "2", "0"; each part to 15 significant digits. */
function write(re, im, unit = 'i') {
  if (!Number.isFinite(re) || !Number.isFinite(im)) throw E.num;
  const r = re === 0 ? '' : complexPartText(re);
  if (im === 0) return r || '0';
  const t = complexPartText(im);
  const c = t === '1' ? '' : t === '-1' ? '-' : t;
  return `${r}${r && im > 0 ? '+' : ''}${c}${unit || 'i'}`;
}
const out = (z, unit) => write(z.re, z.im, unit);

/** The unit two numbers share: "i" and "j" together are an error (#VALUE! or #NUM!, as the function has it). */
function unitOf(zs, err) {
  let u = '';
  for (const z of zs) { if (z.unit) { if (u && z.unit !== u) throw err; u = z.unit; } }
  return u;
}

// ---- Building and taking apart

lifted('COMPLEX', (a, ctx) => {
  const real = (v) => { const x = one(v, ctx); if (typeof x === 'boolean') throw E.value; return x === MISSING || x == null ? 0 : x; };
  const num = (v) => { const x = real(v); if (typeof x === 'number') return x; const n = textToNumber(x); if (n == null) throw E.value; return n; };
  const re = num(a[0]);
  const im = num(a[1]);
  const unit = given(a[2]) ? one(a[2], ctx) : 'i';
  if (unit !== 'i' && unit !== 'j' && unit !== '') throw E.value;
  return write(re, im, unit);
});
lifted('IMREAL', (a, ctx) => arg(a[0], ctx).re);
lifted('IMAGINARY', (a, ctx) => arg(a[0], ctx).im);
/** |z|; #NUM! where |re| + |im| is past the largest number (IMABS("1E308+1E308i")), as Excel's. */
function modulus(z) {
  if (!Number.isFinite(Math.abs(z.re) + Math.abs(z.im))) throw E.num;
  return Math.hypot(z.re, z.im);
}
lifted('IMABS', (a, ctx) => modulus(arg(a[0], ctx)));
lifted('IMARGUMENT', (a, ctx) => {
  const z = arg(a[0], ctx);
  if (z.re === 0 && z.im === 0) throw E.div0;
  return Math.atan2(z.im, z.re);
});
lifted('IMCONJUGATE', (a, ctx) => { const z = arg(a[0], ctx); return write(z.re, -z.im, z.unit); });

// ---- Arithmetic

/** Every value of IMSUM's or IMPRODUCT's arguments as a complex number, an empty cell of a range as 0. */
function* allOf(args, ctx) {
  const area = function* (x) {
    let seen = 0;
    for (const [, , v] of cellsIn(x, ctx)) { seen++; yield complex(v); }
    if (seen < (x.r1 - x.r0 + 1) * (x.c1 - x.c0 + 1)) yield { re: 0, im: 0, unit: '' };
  };
  for (const a of args) {
    if (a instanceof Area) yield* area(a);
    else if (a instanceof Union) for (const x of a.areas) yield* area(x);
    else if (a instanceof Arr) for (const row of a.rows) for (const v of row) yield complex(v);
    else yield complex(a);
  }
}
def('IMSUM', (args, ctx) => {
  const zs = [...allOf(args, ctx)];
  const unit = unitOf(zs, E.value);
  return write(zs.reduce((s, z) => s + z.re, 0), zs.reduce((s, z) => s + z.im, 0), unit);
});
def('IMPRODUCT', (args, ctx) => {
  const zs = [...allOf(args, ctx)];
  const unit = unitOf(zs, E.value);
  let re = 1;
  let im = 0;
  for (const z of zs) [re, im] = [re * z.re - im * z.im, re * z.im + im * z.re];
  // (A product past the largest number is #VALUE!, as Excel's.)
  if (!Number.isFinite(re) || !Number.isFinite(im)) throw E.value;
  return write(re, im, unit);
});
lifted('IMSUB', (a, ctx) => {
  const x = arg(a[0], ctx);
  const y = arg(a[1], ctx);
  return write(x.re - y.re, x.im - y.im, unitOf([x, y], E.num));
});
/** x / y, scaled (Smith's) so that a small divisor doesn't vanish on the way: 1/(1E-300+1E-300i) is 5E+299-5E+299i. */
function divide(x, y) {
  if (y.re === 0 && y.im === 0) throw E.num;
  if (Math.abs(y.re) >= Math.abs(y.im)) {
    const r = y.im / y.re;
    const d = y.re + y.im * r;
    return { re: (x.re + x.im * r) / d, im: (x.im - x.re * r) / d };
  }
  const r = y.re / y.im;
  const d = y.re * r + y.im;
  return { re: (x.re * r + x.im) / d, im: (x.im * r - x.re) / d };
}
lifted('IMDIV', (a, ctx) => {
  const x = arg(a[0], ctx);
  const y = arg(a[1], ctx);
  const unit = unitOf([x, y], E.num);
  return out(divide(x, y), unit);
});

// ---- Powers, roots and logarithms (by the polar form: IMPOWER("i",2) is -1+1.22514845490862E-16i, as Excel's)

/** r^n (cos nθ + i sin nθ). */
function power(z, n) {
  if (z.re === 0 && z.im === 0) { if (n > 0) return { re: 0, im: 0 }; throw E.num; }
  const rn = modulus(z) ** n;
  const t = n * Math.atan2(z.im, z.re);
  if (Math.abs(t) >= TRIG_LIMIT || !Number.isFinite(rn)) throw E.num;
  return { re: rn * xcos(t), im: rn * xsin(t) };
}
lifted('IMPOWER', (a, ctx) => {
  const z = arg(a[0], ctx);
  const x = one(a[1], ctx);
  if (typeof x === 'boolean') throw E.value;
  const n = typeof x === 'number' ? x : x == null ? 0 : textToNumber(x);
  if (n == null) throw E.value;
  return out(power(z, n), z.unit);
});
lifted('IMSQRT', (a, ctx) => { const z = arg(a[0], ctx); return out(power(z, 0.5), z.unit); });
function ln(z) {
  if (z.re === 0 && z.im === 0) throw E.num;
  return { re: Math.log(modulus(z)), im: Math.atan2(z.im, z.re) };
}
const logIn = (names, base) => lifted(names, (a, ctx) => {
  const z = arg(a[0], ctx);
  const l = ln(z);
  return write(l.re / base, l.im / base, z.unit);
});
logIn('IMLN', 1);
logIn('IMLOG10', Math.LN10);
logIn('IMLOG2', Math.LN2);

// ---- Exponential, circular and hyperbolic functions

/** cosh and sinh of x as e^x makes them: #NUM! once e^|x| is past the largest number (IMCOSH("710")), as Excel's. */
function hyper(x) {
  if (!Number.isFinite(Math.exp(Math.abs(x)))) return null;
  return { cosh: Math.cosh(x), sinh: Math.sinh(x) };
}
const circular = (x) => { if (Math.abs(x) >= TRIG_LIMIT) throw E.num; return { cos: xcos(x), sin: xsin(x) }; };
lifted('IMEXP', (a, ctx) => {
  const z = arg(a[0], ctx);
  const { cos, sin } = circular(z.im);
  const e = Math.exp(z.re);
  if (!Number.isFinite(e)) throw E.num;
  return write(e * cos, e * sin, z.unit);
});
/** sin z and cos z, or null where cosh(im z) is past the largest number. */
function sinCos(z) {
  const { cos, sin } = circular(z.re);
  const h = hyper(z.im);
  if (!h) return null;
  return { sin: { re: sin * h.cosh, im: cos * h.sinh }, cos: { re: cos * h.cosh, im: -sin * h.sinh } };
}
/** sinh z and cosh z, or null where cosh(re z) is past the largest number. */
function sinhCosh(z) {
  const { cos, sin } = circular(z.im);
  const h = hyper(z.re);
  if (!h) return null;
  return { sinh: { re: h.sinh * cos, im: h.cosh * sin }, cosh: { re: h.cosh * cos, im: h.sinh * sin } };
}
const ZERO = { re: 0, im: 0 };
/** A function of z from its sine and cosine (sinCos or sinhCosh); #NUM! where they are past the largest number. */
const unary = (names, of, pair) => lifted(names, (a, ctx) => {
  const z = arg(a[0], ctx);
  const p = pair(z);
  if (!p) throw E.num;
  return out(of(p), z.unit);
});
unary('IMSIN', (p) => p.sin, sinCos);
unary('IMCOS', (p) => p.cos, sinCos);
unary('IMSINH', (p) => p.sinh, sinhCosh);
unary('IMCOSH', (p) => p.cosh, sinhCosh);

// The tangent and the reciprocals by the double angle, as Excel's last digits have them: tan(x+iy) is
// (sin 2x + i sinh 2y) / (cos 2x + cosh 2y), sec(x+iy) 2(cos x cosh y + i sin x sinh y) / (cos 2x + cosh 2y). A real
// number has the real functions' (IMTAN("1.5707963267949") is TAN's -286413440886091). Where cosh 2y is past the
// largest number, the tangent is ±i and a reciprocal 1 over the function itself, or 0.
/** The parts at x (the circular one) and y (the hyperbolic one): null where cosh 2y is past the largest number. */
function doubled(x, y) {
  const h2 = hyper(2 * y);
  if (!h2) return null;
  return { c2: circular(2 * x), h2, c: circular(x), h: hyper(y) };
}
const sign = (v) => (v < 0 ? -1 : 1);
const over = (re, im, den) => { if (den === 0) throw E.num; return { re: re / den, im: im / den }; };
/** 1 over a function of z where the double angle is too big (IMSECH("400") is 3.83033919342801E-174), else 0. */
const recip = (fn, part) => (z) => { const p = fn(z); return p ? divide({ re: 1, im: 0 }, p[part]) : ZERO; };
/** A function of z by the double angle; x and y are the real part and the imaginary one, the reverse for a hyperbolic
 *  function. `real` its value at a real x (y 0) from sin x and cos x; `big` its value where cosh 2y is too big. */
const angled = (names, hyperbolic, of, real, big) => lifted(names, (a, ctx) => {
  const z = arg(a[0], ctx);
  const [x, y] = hyperbolic ? [z.im, z.re] : [z.re, z.im];
  if (y === 0 && real) return out(real(circular(x)), z.unit);
  const d = doubled(x, y);
  return out(d ? of(d) : big(z, y), z.unit);
});
const ratio = (num, den) => (c) => over(c[num], 0, c[den]);
angled('IMTAN', false, ({ c2, h2 }) => over(c2.sin, h2.sinh, c2.cos + h2.cosh), ratio('sin', 'cos'), (z, y) => ({ re: 0, im: sign(y) }));
angled('IMCOT', false, ({ c2, h2 }) => over(c2.sin, -h2.sinh, h2.cosh - c2.cos), ratio('cos', 'sin'), (z, y) => ({ re: 0, im: -sign(y) }));
angled('IMSEC', false, ({ c2, h2, c, h }) => over(2 * c.cos * h.cosh, 2 * c.sin * h.sinh, c2.cos + h2.cosh), null, recip(sinCos, 'cos'));
angled('IMCSC', false, ({ c2, h2, c, h }) => over(2 * c.sin * h.cosh, -2 * c.cos * h.sinh, h2.cosh - c2.cos), null, recip(sinCos, 'sin'));
angled('IMSECH', true, ({ c2, h2, c, h }) => over(2 * h.cosh * c.cos, -2 * h.sinh * c.sin, h2.cosh + c2.cos), null, recip(sinhCosh, 'cosh'));
angled('IMCSCH', true, ({ c2, h2, c, h }) => over(2 * h.sinh * c.cos, -2 * h.cosh * c.sin, h2.cosh - c2.cos), null, recip(sinhCosh, 'sinh'));

export { parse as parseComplex, write as complexText };
