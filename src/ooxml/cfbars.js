// What data bars and icon sets give the cells of their ranges, by Excel's documented rules (Excel draws them but COM
// can't read a bar's length or a cell's icon back, so these aren't measured): a bar's ends from its min and max (lowest
// or highest value; automatic: 0 if that is further; a number, percent or percentile), its length between the rule's
// min and max lengths; with negative numbers and an automatic axis, the axis where 0 falls and bars either side of it.
// An icon set's thresholds (percent of the span, percentile, number) give each number the last icon it reaches.

/** The numbers of a rule's cells: [[r, c, v]] and their sorted values. */
export function numbersOf(cells) {
  const nums = cells.filter(([, , v]) => typeof v === 'number');
  return { nums, sorted: nums.map(([, , v]) => v).sort((a, b) => a - b) };
}
/** Excel's PERCENTILE.INC of sorted values. */
const percentile = (s, p) => { const k = (s.length - 1) * p / 100; const i = Math.floor(k); return s[i] + ((s[Math.min(i + 1, s.length - 1)] ?? s[i]) - s[i]) * (k - i); };
/** A cfvo's value against sorted numbers (`end` 'min' or 'max' for the automatic ones). */
function point(v, s, end) {
  const lo = s[0];
  const hi = s[s.length - 1];
  const n = Number(v?.val);
  switch (v?.type) {
    case 'min': return lo;
    case 'max': return hi;
    case 'autoMin': return Math.min(0, lo);
    case 'autoMax': return Math.max(0, hi);
    case 'num': return Number.isFinite(n) ? n : end === 'min' ? lo : hi;
    case 'percent': return lo + (hi - lo) * (Number.isFinite(n) ? n : 0) / 100;
    case 'percentile': return percentile(s, Number.isFinite(n) ? n : 0);
    default: return end === 'min' ? lo : hi;
  }
}

/** Each number's bar: Map("r,c" => { from, to (fractions of the cell's width), negative, axis (a fraction or null) }). */
export function barSpans(cells, bar) {
  const { nums, sorted } = numbersOf(cells);
  const out = new Map();
  if (!sorted.length) return out;
  let min = point(bar.min, sorted, 'min');
  let max = point(bar.max, sorted, 'max');
  if (max < min) [min, max] = [max, min];
  const lenMin = Math.max(0, Math.min(100, bar.minLength)) / 100;
  const lenMax = Math.max(lenMin, Math.min(100, bar.maxLength)) / 100;
  const axisAt = bar.axis === 'middle' ? 0.5 : bar.axis === 'automatic' && min < 0 && max > 0 ? -min / (max - min) : bar.axis === 'automatic' && max <= 0 && min < 0 ? 1 : null;
  for (const [r, c, v] of nums) {
    const x = Math.max(min, Math.min(max, v));
    if (axisAt == null) {
      // (No axis: from the left, the min length at min, the max length at max.)
      const f = max === min ? 1 : (x - min) / (max - min);
      out.set(`${r},${c}`, { from: 0, to: lenMin + f * (lenMax - lenMin), negative: false, axis: null });
    } else if (x >= 0) {
      const room = 1 - axisAt;
      const f = max > 0 ? x / max : 0;
      out.set(`${r},${c}`, { from: axisAt, to: axisAt + room * (lenMin + f * (lenMax - lenMin)), negative: false, axis: axisAt });
    } else {
      const f = min < 0 ? x / min : 0;
      out.set(`${r},${c}`, { from: axisAt - axisAt * (lenMin + f * (lenMax - lenMin)), to: axisAt, negative: true, axis: axisAt });
    }
  }
  return out;
}

/** How many icons each set has. */
export const ICON_COUNTS = { '3Arrows': 3, '3ArrowsGray': 3, '3Flags': 3, '3TrafficLights1': 3, '3TrafficLights2': 3, '3Signs': 3, '3Symbols': 3, '3Symbols2': 3,
  '3Stars': 3, '3Triangles': 3, '4Arrows': 4, '4ArrowsGray': 4, '4RedToBlack': 4, '4Rating': 4, '4TrafficLights': 4, '5Arrows': 5, '5ArrowsGray': 5,
  '5Rating': 5, '5Quarters': 5, '5Boxes': 5 };

/** Each number's icon: Map("r,c" => index, 0 the lowest), reversed when the set is. */
export function iconIndexes(cells, icons) {
  const { nums, sorted } = numbersOf(cells);
  const out = new Map();
  if (!sorted.length) return out;
  const n = ICON_COUNTS[icons.set] || icons.cfvo.length || 3;
  const steps = icons.cfvo.slice(1, n).map((v) => ({ at: point(v, sorted, 'min'), gte: v.gte !== false }));
  for (const [r, c, v] of nums) {
    let k = 0;
    steps.forEach((s, i) => { if (s.gte ? v >= s.at : v > s.at) k = i + 1; });
    out.set(`${r},${c}`, icons.reverse ? n - 1 - k : k);
  }
  return out;
}
