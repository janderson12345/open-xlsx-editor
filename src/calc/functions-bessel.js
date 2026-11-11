// BESSELJ, BESSELY, BESSELI and BESSELK as Excel works them out: not the exact functions but the polynomial and
// rational approximations of Numerical Recipes, good to about 1e-8 (=BESSELJ(0,0) is 1.00000000283141 in Excel), with
// recurrences for higher orders: up from orders 0 and 1 (J where x is past n, Y, K), down by Miller's method (J, I).
// Measured against Excel with cell values in a hidden Excel and the formula probe. Registered into functions.js's
// table; evaluate.js imports this file.
import { E, Unsupported } from './values.js';
import { lifted, num, one } from './functions.js';
import { sin, cos, TRIG_LIMIT } from './trig.js';

const TWO_OVER_PI = 0.636619772;

function j0(x) {
  const ax = Math.abs(x);
  if (ax < 8) {
    const y = x * x;
    const a = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * -184.9052456))));
    const b = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y))));
    return a / b;
  }
  const { p, q, z, xx } = far0(ax);
  return Math.sqrt(TWO_OVER_PI / ax) * (cos(xx) * p - z * sin(xx) * q);
}
/** The asymptotic parts for order 0 past 8 (the last coefficient as Excel has it: 0.934945152e-7, which Y0 shows). */
function far0(ax) {
  const z = 8 / ax;
  const y = z * z;
  return {
    z, xx: ax - 0.785398164,
    p: 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6))),
    q: -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934945152e-7))),
  };
}
/** The asymptotic parts for order 1 past 8. */
function far1(ax) {
  const z = 8 / ax;
  const y = z * z;
  return {
    z, xx: ax - 2.356194491,
    p: 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6))),
    q: 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6))),
  };
}
function j1(x) {
  const ax = Math.abs(x);
  if (ax < 8) {
    const y = x * x;
    const a = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.48260 + y * -30.16036606)))));
    const b = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y))));
    return a / b;
  }
  const { p, q, z, xx } = far1(ax);
  const v = Math.sqrt(TWO_OVER_PI / ax) * (cos(xx) * p - z * sin(xx) * q);
  return x < 0 ? -v : v;
}
function y0(x) {
  if (x < 8) {
    const y = x * x;
    const a = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
    const b = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.26470 + y * (226.1030244 + y))));
    return a / b + TWO_OVER_PI * j0(x) * Math.log(x);
  }
  const { p, q, z, xx } = far0(x);
  return Math.sqrt(TWO_OVER_PI / x) * (sin(xx) * p + z * cos(xx) * q);
}
function y1(x) {
  if (x < 8) {
    const y = x * x;
    const a = x * (-0.4900604943e13 + y * (0.1275274390e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
    const b = 0.2499580570e14 + y * (0.4244419664e12 + y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.1020426050e6 + y * (0.3549632885e3 + y)))));
    return a / b + TWO_OVER_PI * (j1(x) * Math.log(x) - 1 / x);
  }
  const { p, q, z, xx } = far1(x);
  return Math.sqrt(TWO_OVER_PI / x) * (sin(xx) * p + z * cos(xx) * q);
}
function i0(x) {
  const ax = Math.abs(x);
  if (ax < 3.75) {
    const y = (x / 3.75) ** 2;
    return 1.0 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))));
  }
  const y = 3.75 / ax;
  return (Math.exp(ax) / Math.sqrt(ax)) * (0.39894228 + y * (0.1328592e-1 + y * (0.225319e-2 + y * (-0.157565e-2 + y * (0.916281e-2
    + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))));
}
function i1(x) {
  const ax = Math.abs(x);
  let v;
  if (ax < 3.75) {
    const y = (x / 3.75) ** 2;
    v = ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))));
  } else {
    const y = 3.75 / ax;
    const t = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2));
    v = (0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * t))))) * (Math.exp(ax) / Math.sqrt(ax));
  }
  return x < 0 ? -v : v;
}
function k0(x) {
  if (x <= 2) {
    const y = x * x / 4;
    return -Math.log(x / 2) * i0(x) + (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))));
  }
  const y = 2 / x;
  return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))));
}
function k1(x) {
  if (x <= 2) {
    const y = x * x / 4;
    return Math.log(x / 2) * i1(x) + (1 / x) * (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4))))));
  }
  const y = 2 / x;
  return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (0.23498619 + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))));
}

/** Up from orders 0 and 1: f(n+1) = 2n/x f(n) - f(n-1) (J and Y), or + f(n-1) (K, `plus`). */
function upward(n, x, f0, f1, plus) {
  if (n === 0) return f0;
  const tox = 2 / x;
  let prev = f0;
  let cur = f1;
  for (let j = 1; j < n; j++) [prev, cur] = [cur, plus ? prev + j * tox * cur : j * tox * cur - prev];
  return cur;
}
const ACC = 40;
const BIG = 1e10;
/** J for n of 2 or more, as Numerical Recipes' bessj: upward where |x| is past n, else Miller's downward recurrence. */
function jn(n, x) {
  const ax = Math.abs(x);
  if (ax === 0) return 0;
  let ans;
  if (ax > n) ans = upward(n, ax, j0(ax), j1(ax), false);
  else {
    const tox = 2 / ax;
    const m = 2 * Math.trunc((n + Math.trunc(Math.sqrt(ACC * n))) / 2);
    let even = false;
    let bjp = 0;
    let sum = 0;
    let bj = 1;
    ans = 0;
    for (let j = m; j > 0; j--) {
      const bjm = j * tox * bj - bjp;
      bjp = bj;
      bj = bjm;
      if (Math.abs(bj) > BIG) { bj /= BIG; bjp /= BIG; ans /= BIG; sum /= BIG; }
      if (even) sum += bj;
      even = !even;
      if (j === n) ans = bjp;
    }
    ans /= 2 * sum - bj;
  }
  return x < 0 && n % 2 === 1 ? -ans : ans;
}
/** I for n of 2 or more, as Numerical Recipes' bessi: Miller's downward recurrence, scaled by I0. */
function iN(n, x) {
  if (x === 0) return 0;
  const tox = 2 / Math.abs(x);
  let bip = 0;
  let bi = 1;
  let ans = 0;
  for (let j = 2 * (n + Math.trunc(Math.sqrt(ACC * n))); j > 0; j--) {
    const bim = bip + j * tox * bi;
    bip = bi;
    bi = bim;
    if (Math.abs(bi) > BIG) { ans /= BIG; bi /= BIG; bip /= BIG; }
    if (j === n) ans = bip;
  }
  ans *= i0(x) / bi;
  return x < 0 && n % 2 === 1 ? -ans : ans;
}

/** x and n: n made whole, below 0 #NUM!; TRUE and FALSE #VALUE!. */
function xn(a, ctx) {
  if (typeof one(a[0], ctx) === 'boolean' || typeof one(a[1], ctx) === 'boolean') throw E.value;
  const x = num(a[0], ctx);
  const n = Math.trunc(num(a[1], ctx));
  if (n < 0) throw E.num;
  return { x, n };
}
const finite = (v) => { if (!Number.isFinite(v)) throw E.num; return v; };
lifted('BESSELJ', (a, ctx) => {
  const { x, n } = xn(a, ctx);
  if (Math.abs(x) >= TRIG_LIMIT) throw E.num;
  // (From 8 on, Excel's J0 and J1 aren't Numerical Recipes': J1(8) is 0.2346283865070744, 8e-6 from the function's
  // value. Not yet found; the cell keeps Excel's value.)
  if (Math.abs(x) >= 8 && !(n >= 2 && Math.abs(x) <= n)) throw new Unsupported('BESSELJ from 8 on');
  return finite(n === 0 ? j0(x) : n === 1 ? j1(x) : jn(n, x));
});
lifted('BESSELY', (a, ctx) => {
  const { x, n } = xn(a, ctx);
  if (x <= 0 || x >= TRIG_LIMIT) throw E.num;
  return finite(upward(n, x, y0(x), y1(x), false));
});
lifted('BESSELI', (a, ctx) => {
  const { x, n } = xn(a, ctx);
  return finite(n === 0 ? i0(x) : n === 1 ? i1(x) : iN(n, x));
});
lifted('BESSELK', (a, ctx) => {
  const { x, n } = xn(a, ctx);
  if (x <= 0) throw E.num;
  return finite(upward(n, x, k0(x), k1(x), true));
});
