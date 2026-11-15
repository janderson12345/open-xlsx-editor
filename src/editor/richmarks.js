// Formatting part of a cell's text while it is edited, as Excel lets you: select some of the text in the cell editor
// and press Ctrl+B, Ctrl+I, Ctrl+U or Ctrl+5 (or the Font group's Bold, Italic, Underline, Strikethrough), and that part
// takes it when the edit is entered, as rich text (strings.js richRunsXml). The marks follow the text as it is edited
// on either side of them. (The editor is a plain text box: the formatting shows once the cell is entered.)

const KEYS = { b: 'b', i: 'i', u: 'u', s: 'strike', strike: 'strike' };

/** Marks [{ start, end, prop }] from a cell's runs (each run's bold, italic, underline and strikethrough). */
export function marksFromRuns(runs) {
  const out = [];
  let at = 0;
  for (const run of runs || []) {
    const end = at + run.text.length;
    for (const prop of ['b', 'i', 'u', 'strike']) if (run.font?.[prop]) out.push({ start: at, end, prop });
    at = end;
  }
  return out;
}

/** `marks` after `before` became `after` (one edit: the common start and end kept, what lay between replaced). */
export function shiftMarks(marks, before, after) {
  let p = 0;
  while (p < before.length && p < after.length && before[p] === after[p]) p++;
  let s = 0;
  while (s < before.length - p && s < after.length - p && before[before.length - 1 - s] === after[after.length - 1 - s]) s++;
  const oldEnd = before.length - s;
  const delta = after.length - before.length;
  const move = (x) => (x <= p ? x : x >= oldEnd ? x + delta : p);
  return marks.map((m) => ({ ...m, start: move(m.start), end: m.start < p && m.end > oldEnd ? m.end + delta : move(m.end) })).filter((m) => m.end > m.start);
}

/** A mark's property toggled over [start, end) (on where it was off anywhere in it, else off). */
export function toggleMark(marks, start, end, key) {
  const prop = KEYS[key];
  if (!prop || end <= start) return marks;
  const on = (i) => marks.some((m) => m.prop === prop && i >= m.start && i < m.end);
  let all = true;
  for (let i = start; i < end; i++) if (!on(i)) { all = false; break; }
  const others = marks.filter((m) => m.prop !== prop);
  // (The property's marks cut out of [start, end), then the range added back when it is turned on.)
  const mine = marks.filter((m) => m.prop === prop).flatMap((m) => [
    ...(m.start < start ? [{ ...m, end: Math.min(m.end, start) }] : []),
    ...(m.end > end ? [{ ...m, start: Math.max(m.start, end) }] : []),
  ]);
  return [...others, ...mine, ...(all ? [] : [{ start, end, prop }])];
}

/** The text as runs [{ text, font }] by the marks, or null when nothing is marked. */
export function runsOf(text, marks) {
  if (!marks.length) return null;
  const cuts = [...new Set([0, text.length, ...marks.flatMap((m) => [m.start, m.end])])].filter((x) => x >= 0 && x <= text.length).sort((a, b) => a - b);
  const runs = [];
  for (let k = 0; k + 1 < cuts.length; k++) {
    const [a, b] = [cuts[k], cuts[k + 1]];
    const props = Object.fromEntries(marks.filter((m) => m.start <= a && m.end >= b).map((m) => [m.prop, true]));
    const font = Object.keys(props).length ? props : null;
    const last = runs[runs.length - 1];
    if (last && JSON.stringify(last.font) === JSON.stringify(font)) last.text += text.slice(a, b);
    else runs.push({ text: text.slice(a, b), font });
  }
  return runs.length > 1 || runs[0]?.font ? runs : null;
}
