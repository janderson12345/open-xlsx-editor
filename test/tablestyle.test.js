// Table styles drawn as Excel shows them: every cell of a table in each of Excel's 60 built-in styles, with every Table
// Style Option on (test/tablelook-all-excel.json) and with only the header and total rows
// (test/tablelook-plain-excel.json), and the option and own-format cases of test/tablelook-excel.json, against Excel's
// DisplayFormat (tools/excel-compare/excel-tablelook.ps1): fill, font colour, bold, italic and the four edges.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import './setup.js';
import { tableLook, tableStyleElements, withOwnFormat, BUILTIN_TABLE_STYLES } from '../src/ooxml/tablestyle.js';
import { DEFAULT_THEME, INDEXED, applyTint } from '../src/ooxml/theme.js';
import { parseRange, colName } from '../src/ooxml/refs.js';

const load = (name) => JSON.parse(readFileSync(new URL(name, import.meta.url), 'utf8'));
/** COM's XlLineStyle/XlBorderWeight for the border styles the built-in styles use. */
const COM_LINE = { '1/2': 'thin', '1/-4138': 'medium', '1/4': 'thick', '-4119/4': 'double', '1/1': 'hair' };

/** Excel's cell string as { fill, color, bold, italic, edges }. */
function excelCell(s) {
  const out = { fill: null, color: null, bold: false, italic: false, top: null, bottom: null, left: null, right: null };
  for (const p of s.split(' ')) {
    const [k, v] = p.split(':');
    if (k === 'fill') out.fill = `#${v}`;
    else if (k === 'font') out.color = `#${v}`;
    else if (k === 'b') out.bold = true;
    else if (k === 'i') out.italic = true;
    else {
      const [a, w, col] = v.split('/');
      out[{ L: 'left', T: 'top', B: 'bottom', R: 'right' }[k]] = `${COM_LINE[`${a}/${w}`] || `${a}/${w}`} #${col}`;
    }
  }
  return out;
}

/** Ours in the same shape: the table's look under the cell's own (black Normal font, as a new workbook's). */
function ourCell(look) {
  const e = (side) => (look?.border?.[side] ? `${look.border[side].style} ${look.border[side].color.toUpperCase()}` : null);
  return { fill: look?.fill?.toUpperCase() || null, color: (look?.color || '#000000').toUpperCase(), bold: !!look?.bold, italic: !!look?.italic, top: e('top'), bottom: e('bottom'), left: e('left'), right: e('right') };
}

/** The table of a case: Excel's defaults (header row, banded rows) unless the case says otherwise. */
function tableOf(c) {
  const a = parseRange(c.ref);
  const totals = c.totals ? 1 : 0;
  return { ...a, headers: c.header === false ? 0 : 1, totals, style: { name: c.style, rows: c.rowStripes !== false, cols: !!c.colStripes, first: !!c.first, last: !!c.last } };
}

const sameCell = (a, b) => JSON.stringify(a) === JSON.stringify(b);

/** Compare every cell inside the case's table; returns the differences. */
function compare(c, own = () => null) {
  const t = tableOf(c);
  const look = tableLook(t, tableStyleElements(null, c.style), DEFAULT_THEME, INDEXED);
  const diffs = [];
  for (let r = t.r0; r <= t.r1; r++) {
    for (let col = t.c0; col <= t.c1; col++) {
      const ref = `${colName(col)}${r + 1}`;
      const ours = ourCell(own(ref, look.get(`${r},${col}`)) ?? look.get(`${r},${col}`));
      const theirs = excelCell(c.cells[ref]);
      if (!sameCell(ours, theirs)) diffs.push(`${c.name} ${ref}: ours ${JSON.stringify(ours)} Excel ${JSON.stringify(theirs)}`);
    }
  }
  return diffs;
}

test('the 60 built-in styles are measured, in the gallery order', () => {
  assert.equal(BUILTIN_TABLE_STYLES.length, 60);
  assert.deepEqual(BUILTIN_TABLE_STYLES.slice(0, 2), ['TableStyleLight1', 'TableStyleLight2']);
  assert.equal(BUILTIN_TABLE_STYLES[59], 'TableStyleDark11');
});

for (const [file, opts] of [['tablelook-all-excel.json', { totals: true, colStripes: true, first: true, last: true }], ['tablelook-plain-excel.json', { totals: true, rowStripes: false }]]) {
  test(`every built-in style as Excel shows it: ${file}`, () => {
    const cases = load(file);
    assert.equal(cases.length, 60);
    const diffs = cases.flatMap((c) => compare({ ...c, ...opts }));
    assert.deepEqual(diffs.slice(0, 10), []);
  });
}

test('Table Style Options and no style as Excel shows them', () => {
  const diffs = load('tablelook-excel.json').filter((c) => !c.name.endsWith('-own')).flatMap((c) => compare({ ...c, ...OPTIONS[c.name] }));
  assert.deepEqual(diffs.slice(0, 10), []);
});

/** The cases' options, as tools/excel-compare/tablelook-cases.json gives them. */
const OPTIONS = Object.fromEntries(JSON.parse(readFileSync(new URL('../tools/excel-compare/tablelook-cases.json', import.meta.url), 'utf8')).map((c) => [c.name, c]));

test("a cell's own format over the table's, as Excel shows it", () => {
  // (The own formats of the cases: fill, font colour as rgb, bold off (the Normal style's), a medium bottom border.)
  const normal = { color: { theme: 1, tint: 0 }, b: false, i: false };
  const fmt = {
    'medium2-own': { B2: { fill: 'FF0000' }, C3: { font: '0000FF' }, D2: { bold: false }, E4: { bottom: 'medium' }, F3: { font: '000000' }, C2: { font: '000000' }, D4: { fill: 'FFFFFF' } },
    'light9-own': { C2: { font: '000000' }, D3: { bottom: 'medium' }, E2: { bold: true } },
  };
  const diffs = [];
  for (const c of load('tablelook-excel.json').filter((x) => fmt[x.name])) {
    const f = fmt[c.name];
    // (The cell below a bordered one sees the same line as its top.)
    const above = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return `${m[1]}${Number(m[2]) - 1}`; };
    diffs.push(...compare({ ...c, ...OPTIONS[c.name] }, (ref, look) => {
      const o = f[ref];
      const up = f[above(ref)];
      if (!o && !up?.bottom) return null;
      const own = {
        fill: o?.fill ? { pattern: 'solid', fg: { rgb: o.fill } } : { pattern: 'none' },
        font: { ...normal, ...(o?.font ? { color: { rgb: o.font } } : {}), ...(o?.bold !== undefined ? { b: o.bold } : {}) },
        border: { ...(o?.bottom ? { bottom: { style: o.bottom, color: { auto: true } } } : {}), ...(up?.bottom ? { top: { style: up.bottom, color: { auto: true } } } : {}) },
      };
      // (What shows: the table's look left by its own format, and the cell's own fill drawn under it.)
      const left = withOwnFormat(look, own, normal, DEFAULT_THEME, INDEXED);
      return o?.fill ? { ...left, fill: `#${o.fill}` } : left;
    }));
  }
  assert.deepEqual(diffs, []);
});

test('tints as Excel gives them: 22 colours lightened and darkened by 10 amounts (test/tints-excel.json)', () => {
  const bad = [];
  for (const [base, row] of Object.entries(load('tints-excel.json'))) {
    for (const [t, want] of Object.entries(row)) if (applyTint(base, Number(t)) !== want) bad.push(`${base} ${t}`);
  }
  assert.deepEqual(bad, []);
});
