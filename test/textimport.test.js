import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { TEXT_IMPORT_CASES, caseBytes } from './textimport-cases.js';
import { splitText, importCells, decodeText, detectSeparator, guessBreaks, splitDelimited, dateField, guessSettings } from '../src/ooxml/textimport.js';
import { blankWorkbook, csvWorkbook } from '../src/ooxml/csv.js';
import { formatCodeOf } from '../src/ooxml/styles.js';
import { localeOf, EN_US } from '../src/ooxml/locale.js';

const EXCEL = JSON.parse(fs.readFileSync('test/textimport-excel.json', 'utf8'));
const TAB = String.fromCharCode(9);
const refOf = (r, c) => `${String.fromCharCode(65 + c)}${r + 1}`;

/** A case imported here: { ref: cell } over a blank sheet. */
function ours(c, today) {
  const wb = blankWorkbook('t');
  const { rows } = splitText(caseBytes(c), { encoding: c.encoding, kind: c.kind || 'delimited', seps: c.seps, quote: c.quote, consecutive: !!c.consecutive, breaks: c.breaks });
  const { contents, styles } = importCells(wb, 0, 0, 0, rows, { columns: c.columns, startRow: c.startRow || 1, locale: EN_US, decimal: c.decimal, group: c.group, today });
  wb.apply(0, contents);
  wb.setStyles(0, styles);
  const out = {};
  for (const [r, line] of wb.sheet(0).cells) for (const [col, cell] of line) if (cell.v != null || cell.f != null) out[refOf(r, col)] = { cell, code: formatCodeOf(wb.styles, cell.s || 0) };
  return out;
}

test("text files imported as Excel's Text Import Wizard reads them: delimiters, qualifiers, fixed width, start row, encodings, column types (Excel's cells, measured)", () => {
  const today = new Date(`${EXCEL.today}T12:00:00`);
  let same = 0;
  let all = 0;
  for (const c of TEXT_IMPORT_CASES) {
    const excel = EXCEL.cases.find((x) => x.name === c.name);
    const got = ours(c, today);
    assert.deepEqual(Object.keys(got).sort(), excel.cells.map((x) => x.ref).sort(), `${c.name}: the cells filled`);
    for (const x of excel.cells) {
      all++;
      const { cell, code } = got[x.ref];
      const where = `${c.name} ${x.ref}`;
      if (x.formula) assert.equal(`=${cell.f}`, x.formula, where);
      if (x.type === 'number') assert.ok(cell.t === 'n' && Math.abs(cell.v - x.value) < 1e-9, `${where}: ${cell.v} for ${x.value}`);
      else if (x.type === 'boolean') assert.equal(cell.v, x.value, where);
      else if (x.type === 'error') assert.equal(cell.v, x.value, where);
      else assert.deepEqual([cell.t, cell.v], ['s', x.value], where);
      // (Excel shows a code without its quotes and backslashes: "$"#,##0 as $#,##0.)
      assert.equal(code.replace(/["\\]/g, ''), x.format, `${where}'s format`);
      same++;
    }
  }
  assert.equal(same, all);
  assert.equal(all, 108);
});

test('encodings told from the bytes: byte order marks, UTF-16 without one, UTF-8, Windows-1252', () => {
  const e = String.fromCharCode(0xe9);
  assert.deepEqual(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0x61])), { text: 'a', encoding: 'utf-8' });
  assert.deepEqual(decodeText(new Uint8Array([0xff, 0xfe, 0x61, 0])), { text: 'a', encoding: 'utf-16le' });
  assert.deepEqual(decodeText(new Uint8Array([0xfe, 0xff, 0, 0x61])), { text: 'a', encoding: 'utf-16be' });
  assert.deepEqual(decodeText(new Uint8Array([0x61, 0, 0x62, 0, 0x63, 0])), { text: 'abc', encoding: 'utf-16le' });
  assert.deepEqual(decodeText(new Uint8Array([0x63, 0xc3, 0xa9])), { text: `c${e}`, encoding: 'utf-8' });
  assert.deepEqual(decodeText(new Uint8Array([0x63, 0xe9, 0x80])), { text: `c${e}${String.fromCharCode(0x20ac)}`, encoding: 'windows-1252' });
  // (Chosen, not guessed: the bytes read that way, a mark of another kind kept as text.)
  assert.equal(decodeText(new Uint8Array([0x63, 0xc3, 0xa9]), 'windows-1252').text, `c${String.fromCharCode(0xc3, 0xa9)}`);
});

test('the delimiter a file was written with, fixed-width breaks guessed, a date column read in its own order', () => {
  assert.equal(detectSeparator('a;b;c\r\n1;2,5;3\r\n'), ';');
  assert.equal(detectSeparator(`a${TAB}b\r\n1${TAB}2\r\n`), TAB);
  assert.equal(detectSeparator('a|b|c\n1|2|3\n'), '|');
  assert.equal(detectSeparator('a,b\n"1,5",2\n'), ',');
  // (One field a line, either way: the locale's list separator.)
  assert.equal(detectSeparator('one\ntwo\n', { prefer: ';' }), ';');
  assert.deepEqual(guessBreaks('ID    NAME    AMOUNT\r\n00012 Widget    15.25\r\n7     Bolt     -3\r\n'), [6, 14]);
  assert.deepEqual(splitDelimited('a;;b\r\n', { seps: ';', consecutive: true }), [['a', 'b']]);
  assert.deepEqual(splitDelimited('a,b;c\n', { seps: ',;' }), [['a', 'b', 'c']]);
  const today = new Date(2026, 0, 15);
  assert.deepEqual(dateField('5.3.2024', 'DMY', { today }), { value: 45356, format: 'm/d/yyyy' });
  assert.deepEqual(dateField('3/5/2024', 'MDY', { today }), { value: 45356, format: 'm/d/yyyy' });
  assert.deepEqual(dateField('5-Mar-2024', 'DMY', { today }), { value: 45356, format: 'm/d/yyyy' });
  assert.equal(dateField('13/5/2024', 'MDY', { today }), null);
  assert.equal(dateField('31/2/2024', 'DMY', { today }), null);
  assert.deepEqual(dateField('5.3', 'DMY', { today }), { value: 46086, format: 'd-mmm' });
});

test('a CSV opened in a locale with a comma for the decimal point: semicolons, 1.234,5, day before month', () => {
  const bytes = new TextEncoder().encode('item;amount;when\r\nalpha;1.234,5;05.03.2024\r\nbeta;-2,25;13/12/2023\r\n');
  const wb = csvWorkbook(bytes, 'de.csv', new Date(2026, 0, 15), { locale: localeOf('de-DE') });
  const sh = wb.sheet(0);
  assert.equal(sh.cells.get(1).get(1).v, 1234.5);
  assert.equal(formatCodeOf(wb.styles, sh.cells.get(1).get(1).s), '#,##0.00');
  assert.equal(sh.cells.get(1).get(2).v, 45356);
  assert.equal(sh.cells.get(2).get(1).v, -2.25);
  assert.equal(sh.cells.get(2).get(2).v, 45273);
  const s = guessSettings(bytes, 'de.csv', localeOf('de-DE'));
  assert.deepEqual([s.seps, s.encoding, s.decimal, s.group], [';', 'utf-8', ',', '.']);
});
