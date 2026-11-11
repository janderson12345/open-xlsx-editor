// Sine and cosine as Excel works them out: the angle brought within π/4 of a multiple of π/2 by the 64-bit π of the
// x87's extended precision (0xC90FDAA22168C235 × 2^-62), not the true π. So =SIN(PI()) is 1.22514845490862E-16, not
// 1.22464679914735E-16, and =COS(PI()/2) is 6.1257422745431E-17: measured with Application.Evaluate. Elsewhere the
// two agree to the last digit. SIN, COS, TAN, SEC, CSC, COT and the complex-number functions use these.

const PI_L = 0xC90FDAA22168C235n;
// (π/2 in three parts of 26, 26 and 12 bits, so that k times each, for k below 2^27, is exact.)
const P1 = Number(PI_L >> 38n) * 2 ** -25;
const P2 = Number((PI_L >> 12n) & 0x3FFFFFFn) * 2 ** -51;
const P3 = Number(PI_L & 0xFFFn) * 2 ** -63;
const HALF_PI = P1 + P2 + P3;

/** x as k quarter turns and what is left, r, within about π/4 of 0. */
function reduce(x) {
  const k = Math.round(x / HALF_PI);
  return { q: ((k % 4) + 4) % 4, r: ((x - k * P1) - k * P2) - k * P3 };
}
/** sin x, as Excel's. */
export function sin(x) {
  const { q, r } = reduce(x);
  return q === 0 ? Math.sin(r) : q === 1 ? Math.cos(r) : q === 2 ? -Math.sin(r) : -Math.cos(r);
}
/** cos x, as Excel's. */
export function cos(x) {
  const { q, r } = reduce(x);
  return q === 0 ? Math.cos(r) : q === 1 ? -Math.sin(r) : q === 2 ? -Math.cos(r) : Math.sin(r);
}
/** tan x as sin x over cos x, so =TAN(PI()/2) is 1.63245522776191E+16, as Excel's. */
export const tan = (x) => sin(x) / cos(x);
/** Excel's circular functions give #NUM! from 2^27 on (=SIN(134217728), =TAN(-134217728)). */
export const TRIG_LIMIT = 2 ** 27;
