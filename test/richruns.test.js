import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { build } from './steps.js';
import { resolveColor } from '../src/ooxml/theme.js';
import { setFont, runsOf } from '../src/editor/richmarks.js';
import { snapshot, pasteCopy, pasteCut, pasteSpecial } from '../src/ooxml/clipops.js';
import { pasteHtml } from '../src/ooxml/cliphtml.js';

// Rich text against Excel's (tools/excel-compare/excel-richruns.ps1 on richruns-cases.json, giving
// test/richruns-excel.json): Format Cells' Font tab on part of a cell's text, Copy and Paste of a cell of rich text, and
// another program's HTML with mixed formatting in one cell. Each cell compared character by character: font, size,
// bold, italic, underline, strikethrough, colour, super- and subscript.
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/richruns-cases.json', 'utf8'));
const excel = JSON.parse(fs.readFileSync('test/richruns-excel.json', 'utf8').replace(new RegExp('^' + String.fromCharCode(0xfeff)), ''));
const answer = (group, name) => excel[group].find((x) => x.name === name);

const UNDERLINE = { single: 2, double: -4119, singleAccounting: 4, doubleAccounting: 5 };
const COM_UNDERLINE = Object.fromEntries(Object.entries(UNDERLINE).map(([k, v]) => [v, k]));

/** Excel's runs, one key per character. */
const excelChars = (runs) => (runs || []).flatMap((r) => Array.from(r.text, () => JSON.stringify(r.font)));
/** Our cell, one key per character as excel-richruns.ps1 reads Excel's (the fixture's Calibri read as Excel's new
 *  workbook's Aptos Narrow). */
function ourChars(wb, si, r, c) {
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  if (!cell || cell.v == null) return [];
  const font = wb.styles.fonts[wb.styles.cellXfs[cell.s || 0].fontId];
  const hex = (color) => (color && !color.auto ? (resolveColor(color, wb.theme)?.replace('#', '') || '000000') : '000000').toUpperCase().slice(-6);
  const key = (f) => JSON.stringify({ name: f.name === 'Calibri' ? 'Aptos Narrow' : f.name, size: f.size, bold: !!f.b, italic: !!f.i, underline: UNDERLINE[f.u === true ? 'single' : f.u] ?? -4142,
    strike: !!f.strike, color: hex(typeof f.color === 'string' ? { rgb: f.color } : f.color), sup: f.vertAlign === 'superscript', sub: f.vertAlign === 'subscript' });
  const text = String(cell.v);
  if (!cell.runs) return Array.from(text, () => key(font));
  // (A run with a font of its own: its bold, italic, underline, strikethrough and raising as it says, off where it says
  // nothing, as richRunsXml writes it and the grid draws it; its font, size and colour the cell's where not given.)
  const whole = (f) => ({ ...f, name: f.name ?? font.name, size: f.size ?? font.size, color: f.color ?? font.color });
  return cell.runs.flatMap((run) => Array.from(run.text, () => key(run.font ? whole(run.font) : font)));
}

/** The Font tab's choices for COM's Font properties (what the box gives richmarks.js setFont). */
function patchOf(font) {
  const out = {};
  for (const [k, v] of Object.entries(font)) {
    if (k === 'FontStyle') { out.b = /Bold/.test(v); out.i = /Italic/.test(v); }
    else if (k === 'Bold') out.b = v;
    else if (k === 'Italic') out.i = v;
    else if (k === 'Superscript') out.vertAlign = v ? 'superscript' : null;
    else if (k === 'Subscript') out.vertAlign = v ? 'subscript' : null;
    else if (k === 'Strikethrough') out.strike = v;
    else if (k === 'Underline') out.u = COM_UNDERLINE[v] || null;
    else if (k === 'Size') out.size = v;
    else if (k === 'Name') out.name = v;
    else if (k === 'color') out.color = { rgb: `FF${v}` };
  }
  return out;
}
/** A case's source cell written at (r, c) as the editor writes it: its text, each part's font set as Format Cells'
 *  Font tab sets it on the characters selected. */
function fill(wb, si, r, c, src) {
  if (src.number != null) { wb.apply(si, [{ r, c, value: src.number }]); return; }
  let fonts = Array(src.text.length).fill(null);
  for (const p of src.parts || []) fonts = setFont(fonts, p.start - 1, p.start - 1 + p.len, patchOf(p.font));
  const runs = runsOf(src.text, fonts);
  wb.apply(si, [{ r, c, value: src.text, ...(runs ? { runs } : {}) }]);
}
function applySteps(wb, steps) {
  for (const x of steps) {
    if (x.contents.length) wb.apply(x.si, x.contents);
    if (x.merges) wb.setMerges(x.si, x.merges);
    if (x.styles.length) wb.setStyles(x.si, x.styles);
  }
}

test("Format Cells' Font tab on part of the text: each character's font as Excel's (13 cases)", () => {
  const { wb, si } = build({ cells: {} });
  for (const [k, c] of cases.part.entries()) {
    fill(wb, si, k, 0, c.src);
    assert.deepEqual(ourChars(wb, si, k, 0), excelChars(answer('part', c.name).runs), c.name);
  }
});

test("Format Cells' Font tab on part of the text: the run written as Excel writes it (vertAlign, u val, strike)", () => {
  const { wb, si } = build({ cells: {} });
  const xml = (name) => {
    const c = cases.part.find((x) => x.name === name);
    fill(wb, si, 0, 0, c.src);
    const el = [...wb._part(si).doc.getElementsByTagName('c')].find((x) => x.getAttribute('r') === 'A1');
    return el.getElementsByTagName('is')[0].getElementsByTagName('r')[1].getElementsByTagName('rPr')[0];
  };
  const kids = (el) => [...el.childNodes].filter((n) => n.nodeType === 1).map((n) => `${n.localName}${n.getAttribute('val') ? `=${n.getAttribute('val')}` : ''}`);
  // (Excel's, from test/richruns-excel.json: the properties first, in its order, then sz, color, rFont.)
  assert.deepEqual(kids(xml('superscript')).slice(0, 2), ['vertAlign=superscript', 'sz=11']);
  assert.deepEqual(kids(xml('double accounting underline')).slice(0, 1), ['u=doubleAccounting']);
  assert.deepEqual(kids(xml('all at once')).slice(0, 7), ['i', 'strike', 'u', 'vertAlign=superscript', 'sz=9', 'color', 'rFont=Georgia']);
  assert.match(answer('part', 'all at once').xml, /<i \/><strike \/><u \/><vertAlign val="superscript" \/><sz val="9" \/><color rgb="FF0070C0" \/><rFont val="Georgia" \/>/);
});

test('Copy and Paste of a cell of rich text: its runs as Excel keeps them or lets them go (20 cases)', () => {
  const { wb, si } = build({ cells: {} });
  const other = build({ cells: {} });
  for (const [k, c] of cases.paste.entries()) {
    const r = 100 + k * 3;
    fill(wb, si, r, 0, c.src);
    if (c.dst) fill(wb, si, r, 2, c.dst);
    const clip = snapshot(wb, si, { r0: r, c0: 0, r1: r, c1: 0, ar: r, ac: 0 }, c.how === 'cut');
    const at = { r0: r, c0: 2, r1: r, c1: 2, ar: r, ac: 2 };
    const want = answer('paste', c.name);
    if (c.how === 'otherBook') {
      applySteps(other.wb, pasteCopy(other.wb, other.si, { r0: 0, c0: 2, r1: 0, c1: 2, ar: 0, ac: 2 }, clip));
      assert.deepEqual(ourChars(other.wb, other.si, 0, 2), excelChars(want.other.runs), c.name);
      continue;
    }
    const PASTE = { '-4104': 'all', '-4122': 'formats', '-4163': 'values', '-4123': 'formulas', 7: 'noBorders', 12: 'valuesFormats' };
    const steps = c.how === 'cut' ? pasteCut(wb, si, at, clip) : c.how === 'paste' ? pasteCopy(wb, si, at, clip)
      : pasteSpecial(wb, si, at, clip, { paste: PASTE[c.paste], transpose: !!c.transpose });
    applySteps(wb, steps);
    assert.deepEqual(ourChars(wb, si, r, 2), excelChars(want.runs), c.name);
  }
});

test("a later run of the cell's own font written with it, as Excel writes it, so Paste Formats leaves it as it was", () => {
  const { wb, si } = build({ cells: {} });
  fill(wb, si, 0, 0, { text: 'Hello brave world', parts: [{ start: 7, len: 5, font: { Bold: true } }] });
  const el = [...wb._part(si).doc.getElementsByTagName('c')].find((x) => x.getAttribute('r') === 'A1');
  const runs = [...el.getElementsByTagName('r')];
  // (Excel's: "Hello " bare, " world" with an rPr of the cell's font, test/richruns-excel.json.)
  assert.deepEqual(runs.map((r) => r.getElementsByTagName('rPr').length), [0, 1, 1]);
  assert.match(answer('paste', 'all').xml, /^<r><t xml:space="preserve">Hello <\/t><\/r><r><rPr><b \/>.*<\/r><r><rPr><sz val="11" \/>.*<t xml:space="preserve"> world<\/t><\/r>$/);
  // (A cell format given to it: the first run takes the new font, the others keep theirs.)
  const s = wb.styleWith(0, { font: { u: 'single' } });
  wb.setStyles(si, [{ r: 0, c: 0, s }]);
  const keys = ourChars(wb, si, 0, 0).map((k) => JSON.parse(k).underline);
  assert.deepEqual(keys, [...Array(6).fill(2), ...Array(11).fill(-4142)]);
});

test("another program's HTML with mixed formatting in one cell: runs as Excel makes them (8 cases)", () => {
  const { wb, si } = build({ cells: {} });
  for (const [k, c] of cases.html.entries()) {
    const r = 300 + k * 3;
    const steps = pasteHtml(wb, si, { r0: r, c0: 4, r1: r, c1: 4, ar: r, ac: 4 }, c.html);
    applySteps(wb, steps);
    assert.deepEqual(ourChars(wb, si, r, 4), excelChars(answer('html', c.name).runs), c.name);
  }
});
