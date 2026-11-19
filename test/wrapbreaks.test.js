import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { breakLines, excelLineCount, excelWrapPt, rowOf, MAX_ROW_PT } from '../src/editor/rowheights.js';
import { excelWidths } from '../src/ooxml/excelwidths.js';
import { fontRowPx, fontLinePx } from '../src/editor/fonts.js';

// Excel 365 over COM (tools/excel-compare/excel-wrapwidths.ps1 on wrapwidths-cases.mjs's made-up cases): rows AutoFit
// at every column width from 15 to 400px (wrapbreaks-excel.json: a pair [px, pt] where the height changes), and the
// narrowest column that keeps 2 to 100 of one character on a line (wrapwidths-excel.json).
const SCAN = JSON.parse(fs.readFileSync('test/wrapbreaks-excel.json', 'utf8'));
const NARROWEST = JSON.parse(fs.readFileSync('test/wrapwidths-excel.json', 'utf8')).cases;
const fontOf = (c) => ({ family: c.font, sizePx: (c.size * 4) / 3, bold: c.bold, italic: !!c.italic, css: '' });
/** A cell's part in a row's height (rowOf): `lines` lines of a font, at the cell's bottom. */
const part = (font, size, bold, italic, lines) => { const l = fontLinePx(font, (size * 4) / 3, bold, false, italic); return { pt: lines * l.row * 0.75, desc: l.desc, bottom: true }; };
const NORMAL = part('Calibri', 11, false, false, 1);
/** The height in points of a row holding the case alone, in a workbook whose Normal style is Calibri 11. */
const rowPt = (c, colPx) => Math.min(MAX_ROW_PT, rowOf([NORMAL, part(c.font, c.size, c.bold, !!c.italic, excelLineCount(c.text, fontOf(c), colPx))]));

test("wrapped rows are Excel's height at every column width from 15 to 400px", () => {
  let rows = 0;
  const off = [];
  for (const c of SCAN.cases) {
    let k = 0;
    for (let px = SCAN.from; px <= SCAN.to; px++) {
      while (k + 1 < c.heights.length && c.heights[k + 1][0] <= px) k++;
      rows++;
      const ours = rowPt(c, px);
      const excel = c.heights[k][1];
      if (ours !== excel) off.push(`${c.font} ${c.size}${c.bold ? ' bold' : ''} ${JSON.stringify(c.text)} at ${px}px: ${ours} (Excel ${excel})`);
    }
  }
  assert.equal(SCAN.cases.length, 164);
  assert.equal(rows, 164 * 386);
  // (Known: a-b-c-d-e-f-g-h in a column of 15 or 16px, narrower than a letter and its margin, where Excel has a line
  // fewer, regular and bold.)
  assert.deepEqual(off.filter((s) => !s.includes('"a-b-c-d-e-f-g-h" at 1')), []);
  assert.ok(off.length <= 4, off.join('\n'));
});

test("the narrowest column that keeps a run of one character on a line is Excel's", () => {
  let checked = 0;
  for (const c of NARROWEST) {
    // (13: the search's floor, a case that would fit in less; and a case wider than the widest column tried, 1790px.)
    if (c.px <= 13 || excelLineCount(c.text, fontOf(c), 1790) > 1) continue;
    checked++;
    assert.equal(excelLineCount(c.text, fontOf(c), c.px), 1, `${c.font} ${c.size} ${c.text[0]} x ${c.text.length} in ${c.px}px`);
    assert.equal(excelLineCount(c.text, fontOf(c), c.px - 1), 2, `${c.font} ${c.size} ${c.text[0]} x ${c.text.length} in ${c.px - 1}px`);
  }
  assert.ok(checked > 480, String(checked));
});

test("Excel's margin: 7px up to 84px of text in Calibri 11, a pixel more every 28px from there", () => {
  const calibri = { family: 'Calibri', sizePx: 44 / 3, bold: false, css: '' };
  // (Digits are 7px: 10 fit in 77px, 12 in 91, 20 in 149, 100 in 729.)
  for (const [n, px] of [[2, 21], [10, 77], [12, 91], [15, 113], [20, 149], [40, 294], [100, 729]]) {
    assert.equal(excelLineCount('0'.repeat(n), calibri, px), 1, `${n} digits in ${px}px`);
    assert.equal(excelLineCount('0'.repeat(n), calibri, px - 1), 2, `${n} digits in ${px - 1}px`);
  }
  // (The margin begins at 2 x ceil(digit / 4) + 3: 9px in Calibri 14, whose digits are 10px.)
  const c14 = { family: 'Calibri', sizePx: (14 * 4) / 3, bold: false, css: '' };
  assert.equal(excelWidths('Calibri', (14 * 4) / 3, false).digit, 10);
  assert.equal(excelLineCount('00', c14, 29), 1);
  assert.equal(excelLineCount('00', c14, 28), 2);
});

test('under 8pt every font is measured as Small Fonts, its bold a pixel wider a character and a line', () => {
  assert.equal(excelWidths('Arial', (7 * 4) / 3, false).small, true);
  assert.equal(excelWidths('No Such Font', 8, false).digit, 5);
  assert.equal(excelWidths('Calibri', (8 * 4) / 3, false).small, false);
  // (Arial 7: 20 digits of 5px fit in 108px; bold, 6px each and one more, in 130.)
  const arial = (bold) => ({ family: 'Arial', sizePx: (7 * 4) / 3, bold, css: '' });
  assert.equal(excelLineCount('0'.repeat(20), arial(false), 108), 1);
  assert.equal(excelLineCount('0'.repeat(20), arial(false), 107), 2);
  assert.equal(excelLineCount('0'.repeat(20), arial(true), 130), 1);
  assert.equal(excelLineCount('0'.repeat(20), arial(true), 129), 2);
});

test('a line breaks after spaces and after a hyphen, which do not count at its end', () => {
  // (Each character 10 wide, 45 to a line.)
  const lines = (text, width = 45) => breakLines(text, (a, b) => (b - a) * 10 <= width);
  assert.deepEqual(lines('aaaa bbbb'), ['aaaa ', 'bbbb']);
  assert.deepEqual(lines('aaaa-bbbb'), ['aaaa-', 'bbbb']);
  assert.deepEqual(lines('aaa - bbbb'), ['aaa - ', 'bbbb']);
  assert.deepEqual(lines('aaaa - bbbb'), ['aaaa ', '- ', 'bbbb']);
  assert.deepEqual(lines('aaaa    '), ['aaaa    ']);
  assert.deepEqual(lines(' aaa'), [' aaa']);
  assert.deepEqual(lines(' aaaa'), [' ', 'aaaa']);
  assert.deepEqual(lines('  aaaa'), ['  ', 'aaaa']);
  assert.deepEqual(lines('aa/bb,cc.dd'), ['aa/b', 'b,cc', '.dd']);
  assert.deepEqual(lines('aaaaaaaaa bb'), ['aaaa', 'aaaa', 'a bb']);
  assert.deepEqual(lines('------'), ['-----', '-']);
  assert.deepEqual(lines('one\r\ntwo\n\nfour'), ['one', 'two', '', 'four']);
  assert.deepEqual(lines(''), ['']);
  assert.deepEqual(lines('abc', 5), ['a', 'b', 'c']);
});

test('a last line that is empty or spaces is not counted after a line of one space', () => {
  const count = (text) => breakLines(text, () => true).length;
  // (Excel's own, over COM.)
  for (const [text, n] of [['abc', 1], ['abc\n', 2], ['abc\n ', 2], ['abc\n\n', 3], ['\nabc', 2], ['abc\n \ndef', 3], ['abc \n', 2], ['abc\n \n', 2], ['abc\n \n ', 2], ['abc\n \n  ', 2], ['abc\n  \n', 3], [' \n', 1], [' \n ', 1], [' \n \n ', 2], ['abc\n \n \n', 3], ['abc\n\n \n', 3], ['abc\n\n \n ', 3], ['abc\n \n\n', 4], ['abc\n \nd', 3], ['abc \n \n', 2], ['\n \n', 2], [' \n\n', 3], ['abc\n \n \n \n ', 4], ['a\n \n \n\n', 5], ['\n', 2], [' ', 1]]) {
    assert.equal(count(text), n, JSON.stringify(text));
  }
});

test('an indent takes its pixels off the column, and a rich text is measured run by run', () => {
  const calibri = { family: 'Calibri', sizePx: 44 / 3, bold: false, css: '' };
  assert.equal(excelLineCount('0'.repeat(10), calibri, 85, 9), 2);
  assert.equal(excelLineCount('0'.repeat(10), calibri, 86, 9), 1);
  // (Ten W: 13px each regular, 14 bold. Half of them bold: 135px and 7 of margin... the margin a pixel more past 140.)
  const runs = [{ text: 'WWWWW', font: null }, { text: 'WWWWW', font: { name: 'Calibri', size: 11, b: true } }];
  assert.equal(excelLineCount('WWWWWWWWWW', calibri, 139), 1);
  assert.equal(excelLineCount('WWWWWWWWWW', calibri, 139, 0, runs), 2);
  assert.equal(excelLineCount('WWWWWWWWWW', calibri, 144, 0, runs), 1);
  // (Runs that don't make up the text are left aside.)
  assert.equal(excelLineCount('WWWWWWWWWW', calibri, 139, 0, [{ text: 'other', font: null }]), 1);
});

test("a font Windows hasn't is measured as Arial; one the table hasn't counts line breaks only without a canvas", () => {
  assert.equal(excelWidths('No Such Font', 15, false), null);
  const lines = (family, text, px) => excelLineCount(text, { family, sizePx: 40 / 3, bold: false, css: '' }, px);
  // (Arial 10: digits of 7px, ten of them in 77px.)
  assert.equal(lines('No Such Font', '0000000000', 77), 1);
  assert.equal(lines('No Such Font', '0000000000', 76), 2);
  assert.equal(lines('Helvetica', '0000000000', 76), 2);
  assert.equal(lines('Segoe Print', 'a b c\nd', 20), 2);
});

test("a rich text's lines are each as tall as the tallest run on them", () => {
  // (Excel's own, over COM: Rows.AutoFit of one cell at each column width, the second part of the text in another
  // size by Range.Characters.)
  const pt = (font, size, bold, a, b, size2, bold2, colPx) => {
    const cell = { family: font, sizePx: (size * 4) / 3, bold, css: '' };
    const runs = [{ text: a, font: null }, { text: b, font: { name: font, size: size2, b: bold2 } }];
    return excelWrapPt(a + b, cell, fontRowPx(font, (size * 4) / 3, bold) * 0.75, colPx, 0, runs, (run) => fontRowPx(run.font.name, (run.font.size * 4) / 3, run.font.b) * 0.75);
  };
  // (Book Antiqua 12 bold, 16.5pt a line, then 10pt regular, 13.5pt.)
  for (const [px, excel] of [[45, 90], [60, 76.5], [80, 60], [100, 43.5], [120, 30], [150, 30], [200, 30], [250, 16.5], [390, 16.5]]) {
    assert.equal(pt('Book Antiqua', 12, true, 'Net Amount ', '(USD, thousands)', 10, false, px), excel, `Book Antiqua at ${px}px`);
  }
  // (Calibri 10, 12.75pt a line, then 16pt bold, 21pt: only the lines the large run is on grow.)
  for (const [px, excel] of [[45, 240], [60, 185.25], [80, 109.5], [100, 96.75], [120, 96.75], [150, 54.75], [200, 42], [250, 42], [300, 42], [390, 21]]) {
    assert.equal(pt('Calibri', 10, false, 'Small first part of it ', 'THEN A LARGE SECOND PART', 16, true, px), excel, `Calibri at ${px}px`);
  }
  // (No runs, or runs that aren't the text: the cell's own line for each.)
  const calibri = { family: 'Calibri', sizePx: 44 / 3, bold: false, css: '' };
  assert.equal(excelWrapPt('0000 0000', calibri, 15, 40), 30);
  assert.equal(excelWrapPt('0000 0000', calibri, 15, 40, 0, [{ text: 'x', font: null }], () => 99), 30);
});

test("an italic's lines break where Excel's do: its own widths and what its line takes beyond them", () => {
  const cases = JSON.parse(fs.readFileSync('test/wrapitalic-excel.json', 'utf8')).cases;
  let checked = 0;
  for (const c of cases) {
    if (c.px <= 13 || excelLineCount(c.text, fontOf(c), 1790) > 1) continue;
    checked++;
    assert.equal(excelLineCount(c.text, fontOf(c), c.px), 1, `${c.font} ${c.size}${c.bold ? ' bold' : ''} italic ${c.text[0]} x ${c.text.length} in ${c.px}px`);
    assert.equal(excelLineCount(c.text, fontOf(c), c.px - 1), 2, `${c.font} ${c.size}${c.bold ? ' bold' : ''} italic ${c.text[0]} x ${c.text.length} in ${c.px - 1}px`);
  }
  assert.ok(checked > 400, String(checked));
  // (Five italic zeros in each font, regular and bold, at each half point from 8 to 36pt: the allowance itself.)
  const lines = JSON.parse(fs.readFileSync('test/italiclines-excel.json', 'utf8'));
  let n = 0;
  for (const [font, v] of Object.entries(lines.fonts)) {
    for (const style of ['italic', 'boldItalic']) {
      v[style].forEach((px, k) => {
        const f = { family: font, sizePx: ((lines.from + k / 2) * 4) / 3, bold: style === 'boldItalic', italic: true, css: '' };
        n++;
        assert.equal(excelLineCount(lines.text, f, px), 1, `${font} ${style} ${lines.from + k / 2}pt in ${px}px`);
        assert.equal(excelLineCount(lines.text, f, px - 1), 2, `${font} ${style} ${lines.from + k / 2}pt in ${px - 1}px`);
      });
    }
  }
  assert.equal(n, 25 * 2 * 57);
});

test("a row is its lines and what the row's deepest descent has over the font's, as Excel's under four Normal fonts", () => {
  const data = JSON.parse(fs.readFileSync('test/wraprows-excel.json', 'utf8'));
  let n = 0;
  for (const set of data.sets) {
    const [, name, size] = set.normal.match(/^(.*) ([\d.]+)$/);
    const normal = part(name, Number(size), false, false, 1);
    assert.equal(rowOf([normal]), set.standard, `the standard row of ${set.normal}`);
    for (const [font, pt, lines, excel] of set.rows) {
      n++;
      assert.equal(rowOf([normal, part(font, pt, false, false, lines)]), excel, `${lines} lines of ${font} ${pt} under ${set.normal}`);
    }
  }
  assert.equal(n, 4 * 264);
});

test('only text at the bottom of its cell takes the deepest descent; a font the table has not takes its own', () => {
  const ten = part('Calibri', 10, false, false, 5);
  // (Five lines of Calibri 10, 17px each, descent 3, under Calibri 11's 4: 86px; aligned to the top 85; beside an
  // empty Calibri 20 cell, descent 7, 89; beside that cell aligned to the top, 86.)
  assert.equal(rowOf([NORMAL, ten]), 86 * 0.75);
  assert.equal(rowOf([NORMAL, { ...ten, bottom: false }]), 85 * 0.75);
  assert.equal(rowOf([NORMAL, ten, part('Calibri', 20, false, false, 1)]), 89 * 0.75);
  assert.equal(rowOf([NORMAL, ten, { ...part('Calibri', 20, false, false, 1), bottom: false }]), 86 * 0.75);
  // (An italic can be taller than its upright: Arial 14, 24px upright and 25 italic.)
  assert.deepEqual(fontLinePx('Arial', (14 * 4) / 3, false, false, true), { row: 25, desc: 5 });
  assert.equal(fontLinePx('Arial', (14 * 4) / 3).row, 24);
  // (A font Windows hasn't is Arial, a roman one Times New Roman.)
  assert.deepEqual(fontLinePx('No Such Font', 16), fontLinePx('Arial', 16));
  assert.deepEqual(fontLinePx('No Such Font', 16, false, true), fontLinePx('Times New Roman', 16));
});
