// The largest and smallest of a list by a loop. Math.max(...list) passes every item as an argument, which V8 refuses
// past about 125,000 of them ("Maximum call stack size exceeded"): =MAX(A:A) over 130,000 numbers, a chart of that many
// points, an outline over that many rows. Each item is taken as a number and a NaN makes the answer NaN, as Math.max
// and Math.min do; empty: -Infinity and Infinity, as Math.max() and Math.min() give.

/** The largest number of `list` (-Infinity when empty). */
export function maxOf(list) {
  let m = -Infinity;
  for (const x of list) {
    const v = +x;
    if (Number.isNaN(v)) return NaN;
    if (v > m) m = v;
  }
  return m;
}

/** The smallest number of `list` (Infinity when empty). */
export function minOf(list) {
  let m = Infinity;
  for (const x of list) {
    const v = +x;
    if (Number.isNaN(v)) return NaN;
    if (v < m) m = v;
  }
  return m;
}
