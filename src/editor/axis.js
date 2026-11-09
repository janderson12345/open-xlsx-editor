// One axis of the grid (its rows or its columns): each index's size in whole pixels and where it starts, so the
// grid can find what is visible from a scroll position without walking every row.

export class Axis {
  /** @param {number} count  how many rows or columns the grid spans
   *  @param {number} size   the default size in pixels
   *  @param {Map<number, number>} sizes  indexes with a size of their own (0 for hidden) */
  constructor(count, size, sizes = new Map()) {
    this.count = count;
    this.offsets = new Float64Array(count + 1);
    for (let i = 0; i < count; i++) this.offsets[i + 1] = this.offsets[i] + (sizes.has(i) ? sizes.get(i) : size);
  }
  /** Where index i starts. */
  start(i) {
    return this.offsets[Math.max(0, Math.min(i, this.count))];
  }
  /** Index i's size (0 when hidden). */
  size(i) {
    return this.offsets[i + 1] - this.offsets[i];
  }
  /** The whole length. */
  get end() {
    return this.offsets[this.count];
  }
  /** The index under a position: the last one starting at or before it (hidden ones, of no size, are passed over). */
  indexAt(pos) {
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offsets[mid] <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}
