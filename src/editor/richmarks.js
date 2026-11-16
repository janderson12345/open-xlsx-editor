// Formatting part of a cell's text while it is edited, as Excel lets you: select some of the text in the cell editor
// and press Ctrl+B, Ctrl+I, Ctrl+U or Ctrl+5 (or the Font group's Bold, Italic, Underline, Strikethrough), and that part
// takes it when the edit is entered, as rich text (strings.js richRunsXml). The text's formatting is kept as one font
// per character (null: the cell's own), so a run's size, colour, font and super- or subscript follow the text as it is
// edited as its bold and italic do. The editor draws it as it is typed (richview.js).

const KEYS = { b: 'b', i: 'i', u: 'u', s: 'strike', strike: 'strike' };
// (Ctrl+U gives a single underline, as Excel's.)
const ON = { b: true, i: true, u: 'single', strike: true };

/** Whether a font sets anything (a run of the cell's own font has none, or only properties turned off). */
const sets = (f) => !!f && Object.values(f).some((v) => v != null && v !== false);
/** A font's own properties, for telling runs apart. */
const keyOf = (f) => (sets(f) ? JSON.stringify(Object.entries(f).filter(([, v]) => v != null && v !== false).sort(([a], [b]) => (a < b ? -1 : 1))) : '');

/** One font per character of a cell's runs [{ text, font }] (null where a run has none). */
export function fontsFromRuns(runs) {
  return (runs || []).flatMap((run) => Array.from({ length: run.text.length }, () => (sets(run.font) ? run.font : null)));
}

/** Whether any character of `fonts` is formatted. */
export const anyFormatted = (fonts) => fonts.some(sets);

/**
 * The one edit that turned `before` into `after`: { a, b, n }, characters [a, b) of `before` replaced by the `n` at
 * `a` of `after`. `sel`, the selection before the edit ([start, end], null when not known), and `caret`, the caret after
 * it, place it where the text alone can't (typing "b" after "ab" adds the second b, not the third); otherwise the
 * common start and end of the two texts.
 */
export function editAt(before, after, sel, caret) {
  if (sel && caret != null) {
    const a = Math.min(sel[0], caret);
    const b = before.length - after.length + caret;
    if (a <= b && b <= before.length && a <= caret && before.slice(0, a) === after.slice(0, a) && before.slice(b) === after.slice(caret)) return { a, b, n: caret - a };
  }
  let p = 0;
  const most = Math.min(before.length, after.length, caret ?? Infinity);
  while (p < most && before[p] === after[p]) p++;
  let s = 0;
  while (s < before.length - p && s < after.length - p && before[before.length - 1 - s] === after[after.length - 1 - s]) s++;
  return { a: p, b: before.length - s, n: after.length - s - p };
}

/**
 * `fonts` after edit { a, b, n }. The new characters take a font as Excel's own insert gives them
 * (Characters(start, length).Insert over COM, test/richtype-excel.json): text added at a point takes the font of the
 * character before it (after a bold run bold, before one plain); at the very start, the first character's; text put
 * in place of a selection takes the font of the selection's first character. (At the end of the text the object model
 * gave the cell's own font; typing at the end is taken to go on in the font before the caret like everywhere else,
 * which isn't measured: keystrokes can't be sent to Excel over COM.)
 */
export function applyEdit(fonts, { a, b, n }) {
  const font = (b > a ? fonts[a] : a > 0 ? fonts[a - 1] : fonts[0]) ?? null;
  return [...fonts.slice(0, a), ...Array.from({ length: n }, () => font), ...fonts.slice(b)];
}

/** A property toggled over [start, end): on where it was off anywhere in it, else off. */
export function toggleFont(fonts, start, end, key) {
  const prop = KEYS[key];
  if (!prop || end <= start) return fonts;
  let all = true;
  for (let i = start; i < end; i++) if (!fonts[i]?.[prop]) { all = false; break; }
  // (One new font for each font the range had, so a run stays one object and is drawn as one.)
  const made = new Map();
  return fonts.map((f, i) => {
    if (i < start || i >= end) return f;
    if (!made.has(f)) {
      const next = { ...(f || {}), [prop]: all ? (prop === 'u' ? null : false) : ON[prop] };
      made.set(f, sets(next) ? next : null);
    }
    return made.get(f);
  });
}

/**
 * A font's properties set over [start, end): `patch` { name, size, color, b, i, u, strike, vertAlign }, what the Font
 * group's font, size and colour (or Format Cells' Font tab) give part of the text being edited. A new font name drops
 * the run's theme scheme and family, which belonged to the old one (strings.js richRunsXml then writes the new name
 * bare, as it writes any font of a run's own).
 */
export function setFont(fonts, start, end, patch) {
  if (end <= start || !patch) return fonts;
  const made = new Map();
  return fonts.map((f, i) => {
    if (i < start || i >= end) return f;
    if (!made.has(f)) {
      const next = { ...(f || {}), ...patch };
      if (patch.name) { next.scheme = null; next.family = null; }
      made.set(f, sets(next) ? next : null);
    }
    return made.get(f);
  });
}

/** The size of the first character of [start, end) (`base`, the cell's own, where it has none). */
export const sizeAt = (fonts, start, base) => fonts[start]?.size ?? base;

/** The text as runs [{ text, font }] by its fonts, or null when nothing is formatted. */
export function runsOf(text, fonts) {
  if (!anyFormatted(fonts)) return null;
  const runs = [];
  let key = null;
  for (let i = 0; i < text.length; i++) {
    const f = sets(fonts[i]) ? fonts[i] : null;
    const k = keyOf(f);
    if (runs.length && k === key) runs[runs.length - 1].text += text[i];
    else runs.push({ text: text[i], font: f });
    key = k;
  }
  return runs.length > 1 || runs[0]?.font ? runs : null;
}
