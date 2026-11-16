import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { localeOf, EN_US, setLocale, currentLocale, localCode } from '../src/ooxml/locale.js';
import { cellDisplay } from '../src/ooxml/numfmt.js';
import { parseInput, editText } from '../src/ooxml/input.js';
import { toLocalFormula, fromLocalFormula } from '../src/calc/localform.js';
import { entryChanges } from '../src/ooxml/entry.js';
import { blankWorkbook } from '../src/ooxml/csv.js';
import { formatCodeOf } from '../src/ooxml/styles.js';

const SHOWN = JSON.parse(fs.readFileSync('test/separators-excel.json', 'utf8')).shown;
const today = new Date(2026, 0, 15);
const de = localeOf('de-DE');
const gb = localeOf('en-GB');
const fr = localeOf('fr-FR');

test("numbers shown with other separators as Excel shows them (Excel's Range.Text with its own separators changed, measured)", () => {
  let n = 0;
  for (const x of SHOWN) {
    if (/^#+$/.test(x.text)) continue;
    const loc = localeOf({ tag: 'en-US', decimal: x.decimal, group: x.group });
    const got = cellDisplay({ t: 'n', v: x.value }, 60, x.code, false, loc).text;
    assert.equal(got, x.text, `${x.setting} ${x.code} ${x.value}`);
    n++;
  }
  assert.equal(n, 300);
});

test("locales: Windows' settings for a tag, Intl's for others, a setting changed over a tag's", () => {
  assert.equal(localeOf('en-US'), localeOf(EN_US));
  assert.deepEqual([de.decimal, de.group, de.dateOrder, de.shortDate, de.listSep, de.arrayCol], [',', '.', 'DMY', 'dd.mm.yyyy', ';', '.']);
  assert.deepEqual([gb.decimal, gb.dateOrder, gb.shortDate, gb.listSep], ['.', 'DMY', 'dd/mm/yyyy', ',']);
  assert.equal(localeOf('ja-JP').dateOrder, 'YMD');
  assert.deepEqual([localeOf('it-CH').decimal, localeOf('es-AR').decimal, localeOf('es-AR').listSep], ['.', ',', ';']);
  const us = localeOf({ tag: 'en-US', dateOrder: 'DMY' });
  assert.deepEqual([us.dateOrder, us.shortDate, us.decimal], ['DMY', 'dd/mm/yyyy', '.']);
  assert.equal(localeOf({ tag: 'en-US', decimal: ',' }).listSep, ';');
  // (The page's locale: en-US until set.)
  assert.equal(currentLocale(), EN_US);
  setLocale('de-DE');
  assert.equal(currentLocale().tag, 'de-DE');
  setLocale(EN_US);
});

test("built-in format 14 is the locale's short date; other date codes are kept as they are", () => {
  assert.equal(localCode('m/d/yyyy', de), 'dd.mm.yyyy');
  assert.equal(localCode('m/d/yyyy h:mm', gb), 'dd/mm/yyyy h:mm');
  assert.equal(localCode('d-mmm-yy', de), 'd-mmm-yy');
  assert.equal(cellDisplay({ t: 'n', v: 45356 }, Infinity, 'm/d/yyyy', false, de).text, '05.03.2024');
  assert.equal(cellDisplay({ t: 'n', v: 45356 }, Infinity, 'm/d/yyyy', false, gb).text, '05/03/2024');
  assert.equal(cellDisplay({ t: 'n', v: 45356.5 }, Infinity, 'm/d/yyyy h:mm', false, fr).text, '05/03/2024 12:00');
  assert.equal(cellDisplay({ t: 'n', v: 45356 }, Infinity, 'm/d/yyyy', false, EN_US).text, '3/5/2024');
  assert.equal(cellDisplay({ t: 'n', v: 1234.5 }, Infinity, '#,##0.00', false, fr).text, '1 234,50');
});

test('typing in another locale: its separators, day before month, a dot between date parts, ; in formulas', () => {
  const read = (t, loc) => parseInput(t, { today, locale: loc });
  assert.deepEqual(read('1.234,5', de), { value: 1234.5, format: '#,##0.00', kind: 'number' });
  assert.deepEqual(read('1,5', de), { value: 1.5, format: undefined, kind: 'number' });
  assert.deepEqual(read('-2,25', de), { value: -2.25, format: undefined, kind: 'number' });
  assert.deepEqual(read('50,5%', de), { value: 0.505, format: '0.00%', kind: 'number' });
  assert.deepEqual(read('1,2E3', de), { value: 1200, format: '0.00E+00', kind: 'number' });
  assert.deepEqual(read('1 234,5', fr), { value: 1234.5, format: '#,##0.00', kind: 'number' });
  assert.deepEqual(read(`1${String.fromCharCode(0xa0)}234`, fr), { value: 1234, format: '#,##0', kind: 'number' });
  // (A dot isn't a decimal point in Germany: 1.5 is the 1st of May, as there.)
  assert.deepEqual(read('1.5', de), { value: 46143, format: 'd-mmm', kind: 'date' });
  assert.deepEqual(read('5.3.2024', de), { value: 45356, format: 'm/d/yyyy', kind: 'date' });
  assert.deepEqual(read('13/5/2024', gb), { value: 45425, format: 'm/d/yyyy', kind: 'date' });
  assert.deepEqual(read('13/5', gb), { value: 46155, format: 'd-mmm', kind: 'date' });
  assert.deepEqual(read('5/2024', gb), { value: 45413, format: 'mmm-yy', kind: 'date' });
  assert.deepEqual(read('2024-03-05', gb), { value: 45356, format: 'm/d/yyyy', kind: 'date' });
  assert.deepEqual(read('13/5/2024', EN_US), { value: '13/5/2024' });
  assert.deepEqual(read('24/3/5', localeOf('ja-JP')), { value: 45356, format: 'm/d/yyyy', kind: 'date' });
  assert.equal(read('05.03.2024 13:45', de).kind, 'datetime');
  assert.deepEqual(read('=ROUND(1,5;0)', de), { formula: 'ROUND(1.5,0)' });
  assert.deepEqual(read('=SUM({1.2;3.4})', de), { formula: 'SUM({1,2;3,4})' });
  assert.deepEqual(read('=IF(A1>0;"a;b";B1)', de), { formula: 'IF(A1>0,"a;b",B1)' });
  // (en-US as it was: 1,234 grouped, 1.5 a number.)
  assert.deepEqual(read('1,234', EN_US), { value: 1234, format: '#,##0', kind: 'number' });
  assert.deepEqual(read('1.5', EN_US), { value: 1.5, format: undefined, kind: 'number' });
});

test("formulas shown and typed in a comma-decimal locale: ; between arguments, , in numbers, quotes and names kept", () => {
  const f = `IF(A1>1.5,"a,b",{1,2;3,4})+'my,sheet'!B1+LOG10(2.5)+Table1[[#This Row],[Col,1]]`;
  const shown = `IF(A1>1,5;"a,b";{1.2;3.4})+'my,sheet'!B1+LOG10(2,5)+Table1[[#This Row],[Col,1]]`;
  assert.equal(toLocalFormula(f, de), shown);
  assert.equal(fromLocalFormula(shown, de), f);
  assert.equal(toLocalFormula(f, gb), f);
  assert.equal(toLocalFormula('SUM(1E+3,.5)', de), 'SUM(1E+3;,5)');
  assert.equal(fromLocalFormula('SUM(1E+3;,5)', de), 'SUM(1E+3,.5)');
});

test("the formula bar's text in a locale: its short date with four-digit years, a 24-hour time, its decimal, ; in formulas; typed back, the same", () => {
  assert.equal(editText({ t: 'n', v: 45356.75 }, 'm/d/yyyy', false, null, de), '05.03.2024 18:00:00');
  assert.equal(editText({ t: 'n', v: 0.75 }, 'h:mm', false, null, de), '18:00:00');
  assert.equal(editText({ t: 'n', v: 45356 }, 'm/d/yyyy', false, null, localeOf('fr-BE')), '05-03-2024');
  assert.equal(editText({ t: 'n', v: 1.5 }, 'General', false, null, de), '1,5');
  assert.equal(editText({ t: 'n', v: 0.125 }, '0.0%', false, null, de), '12,5%');
  assert.equal(editText({ t: 'n', v: 3.5, f: 'SUM(1.5,2)' }, 'General', false, null, de), '=SUM(1,5;2)');
  for (const [v, code] of [[45356.75, 'm/d/yyyy'], [1.5, 'General'], [0.125, '0.0%'], [45356, 'm/d/yyyy']]) {
    for (const loc of [de, gb, fr, localeOf('ja-JP')]) {
      const back = parseInput(editText({ t: 'n', v }, code, false, null, loc), { today, locale: loc });
      assert.ok(Math.abs(back.value - v) < 1e-9, `${loc.tag} ${code} ${v}: ${JSON.stringify(back)}`);
    }
  }
});

test('an entry typed in the page locale: entryChanges reads it there, the file keeps en-US', () => {
  const wb = blankWorkbook('t');
  setLocale('de-DE');
  try {
    const ch = entryChanges(wb, 0, 0, 0, '1.234,5', today);
    assert.equal(ch.content.value, 1234.5);
    assert.equal(formatCodeOf(wb.styles, ch.style.s), '#,##0.00');
    assert.equal(entryChanges(wb, 0, 0, 1, '=RUNDEN(1,5;0)', today).content.formula, 'RUNDEN(1.5,0)');
  } finally {
    setLocale(EN_US);
  }
});
