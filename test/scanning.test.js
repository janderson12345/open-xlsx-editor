// The code scanning findings fixed (GitHub CodeQL): inputs that made a regular expression slow now run in linear time,
// and a rich text run's font attributes can't break out of their quotes. With the expressions they replaced, the AM/PM
// and trailing-zeros cases below took about 6 seconds each (Node 20); the date-time split and quoted sheet names were
// flagged for inputs like theirs. Each must now finish well inside a second.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseInput } from '../src/ooxml/input.js';
import { parseFormula } from '../src/calc/parse.js';
import { richRunsXml } from '../src/ooxml/strings.js';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseXml, serializeXml, NS } from '../src/ooxml/xml.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { readStyles } from '../src/ooxml/styles.js';
import { outlineOf } from '../src/ooxml/outline.js';
import { outlineRoom } from '../src/editor/geometry.js';
import { tableLook, tableLooks } from '../src/ooxml/tablestyle.js';
import { maxOf, minOf } from '../src/ooxml/listmath.js';

const quick = (label, fn) => {
  const t0 = performance.now();
  const result = fn();
  const ms = performance.now() - t0;
  assert.ok(ms < 500, `${label}: ${ms.toFixed(0)} ms`);
  return result;
};
const tryParse = (f) => { try { return parseFormula(f); } catch (e) { return e; } };

test('typed entries: long runs of spaces read in linear time, and read as before', () => {
  const spaces = ' '.repeat(100000);
  assert.deepEqual(quick('date-time split', () => parseInput(`! ${spaces}x`)).value, `! ${spaces}x`.trim());
  assert.equal(typeof quick('AM/PM', () => parseInput(`1${spaces}1pm`)).value, 'string');
  // (What the slices read: AM and PM after a space, the hour or time before it; a date and a time.)
  assert.equal(parseInput('10 am').value, 10 / 24);
  assert.equal(parseInput('10:30   PM').value, (22 * 60 + 30) / 1440);
  assert.equal(typeof parseInput('10:30pm').value, 'string');
  assert.equal(parseInput('1/2/2024 10:30').value, parseInput('1/2/2024').value + (10 * 60 + 30) / 1440);
  assert.equal(parseInput('1/2/2024   10:30 PM').value, parseInput('1/2/2024').value + (22 * 60 + 30) / 1440);
});

test('formulas: a long run of quotes and a long number tokenize in linear time; quoted sheet names as before', () => {
  quick('quotes', () => tryParse(`'${"''&".repeat(30000)}`));
  quick('quotes, closed', () => tryParse(`'${"''&".repeat(30000)}'!A1`));
  quick('zeros', () => tryParse(`1${'0'.repeat(100000)}1`));
  const ref = (f) => JSON.stringify(parseFormula(f));
  assert.match(ref("'Q1 ''24'!A1"), /Q1 '24/);
  assert.match(ref("'a''b'!B2+1"), /a'b/);
  assert.ok(tryParse("'unclosed!A1") instanceof Error);
  assert.ok(tryParse("''!A1") instanceof Error);
});

test('rich text: a run\'s font attributes are escaped, so a value with quotes stays inside its attribute', () => {
  const xml = richRunsXml([{ text: 'x', font: { u: 'double"/><evil/>', vertAlign: '"', color: 'FF0000"/>', name: 'A"B' } }],
    { name: 'Calibri', size: 11 });
  assert.ok(!xml.includes('<evil'), xml);
  assert.ok(!/val="[^"]*"[^ />]/.test(xml), xml);
  assert.match(xml, /<u val="double&quot;\/&gt;&lt;evil\/&gt;"\/>/);
});

// The security review of 0.7.0 (crafted workbooks that hung the page or put XML of their own into a saved one).
const sheetXml = (inner, cols = '') => parseXml(`<worksheet xmlns="${NS.x}">${cols}<sheetData>${inner}</sheetData></worksheet>`);

test('a table style\'s stripe size of 0, negative or huge is read as 1 to 9, so the stripes advance', () => {
  const styles = readStyles(parseXml(`<styleSheet xmlns="${NS.x}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs><tableStyles count="1"><tableStyle name="X"><tableStyleElement type="firstRowStripe" size="-1" dxfId="0"/><tableStyleElement type="secondRowStripe" size="99" dxfId="0"/></tableStyle><tableStyle name="__proto__"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles></styleSheet>`));
  assert.deepEqual([styles.tableStyles.X.firstRowStripe.size, styles.tableStyles.X.secondRowStripe.size], [1, 9]);
  assert.equal(Object.getPrototypeOf(styles.tableStyles), null);
  const t = { r0: 0, c0: 0, r1: 50, c1: 3, headers: 1, totals: 0, style: { rows: true } };
  const els = { firstRowStripe: { dxf: { fill: { rgb: 'FFFF0000' } }, size: -1 }, secondRowStripe: { dxf: { fill: { rgb: 'FF00FF00' } }, size: 0 } };
  assert.ok(quick('stripes of size -1 and 0', () => tableLook(t, els, null, null)).size > 0);
});

test('a table over the whole sheet: its looks worked out a block at a time, as drawn', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const t = { name: 'Huge', si: 0, r0: 0, c0: 0, r1: 1048575, c1: 16383, headers: 1, totals: 0, style: { name: 'TableStyleMedium2', rows: true } };
  const looks = quick('a whole-sheet table', () => tableLooks(wb, [t], 0));
  assert.ok(quick('one block of it', () => looks.get('500000,100')));
});

test('outline levels past 7 and column spans past the sheet are read within Excel\'s limits', () => {
  const sheet = readSheet(sheetXml('<row r="1" outlineLevel="100000000"/>', '<cols><col min="1" max="2000000000" outlineLevel="1"/></cols>'));
  assert.equal(sheet.rows.get(0).level, 7);
  assert.deepEqual([sheet.cols[0].min, sheet.cols[0].max], [0, 16383]);
  const room = quick('the outline of a two-billion-column span', () => outlineRoom(sheet));
  assert.ok(room.w < 200 && room.h < 200, JSON.stringify(room));
  quick('its outline', () => outlineOf(sheet));
});

test('large lists: the largest and smallest by a loop (Math.max(...list) refuses past about 125,000 items)', () => {
  const many = Array.from({ length: 300000 }, (_, k) => k % 1000);
  assert.deepEqual([maxOf(many), minOf(many), maxOf([]), minOf([]), maxOf([1, NaN, 2])], [999, 0, -Infinity, Infinity, NaN]);
  const rows = Array.from({ length: 200000 }, (_, k) => `<row r="${k + 1}" outlineLevel="1"/>`).join('');
  const big = readSheet(sheetXml(rows));
  // (It threw "Maximum call stack size exceeded" before; timed alone, the sheet's reading aside.)
  assert.equal(quick('an outline over 200,000 rows', () => outlineOf(big)).rows.level, 1);
});

test('a theme heading font with quotes and tags stays one attribute in the styles a cell style adds', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.theme = { ...wb.theme, fonts: { ...wb.theme.fonts, major: 'X"/><injected evil="1"/><name val="Y' } };
  wb.setStyles(0, [{ r: 1, c: 1, s: wb.styleWith(0, { cellStyle: 'Title' }) }]);
  const xml = serializeXml(wb.stylesDoc);
  assert.ok(!xml.includes('<injected'));
  assert.match(xml, /<name val="X&quot;\/&gt;&lt;injected evil=&quot;1&quot;\/&gt;&lt;name val=&quot;Y"\/>/);
});

test('ACCRINT with a date past 9999-12-31 is #NUM!, as Excel gives, at once', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  quick('ACCRINT over a hundred million days', () => wb.setFormula(0, { r: 300, c: 0 }, 'ACCRINT(1,99999999,100000000,0.1,1000,4)'));
  assert.equal(wb.sheet(0).cells.get(300).get(0).v, '#NUM!');
});
