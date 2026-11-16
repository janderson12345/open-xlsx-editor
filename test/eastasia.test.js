import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { localeOf, EN_US } from '../src/ooxml/locale.js';
import { cellDisplay } from '../src/ooxml/numfmt.js';
import { builtinFormat } from '../src/ooxml/format.js';
import { EAST_ASIAN_FORMATS, EAST_ASIAN_IDS } from '../src/ooxml/eastasia.js';

// Excel 365 over COM (tools/excel-compare/excel-eastasia.ps1): the ids alone, and every locale's code, Range.Text.
const EXCEL = JSON.parse(fs.readFileSync('test/eastasia-excel.json', 'utf8'));
const shown = (code, value, loc = EN_US) => cellDisplay({ t: 'n', v: value }, Infinity, code, false, loc).text;

test("every East Asian code as Excel draws it (eras, gannen, ROC's e, 上午/下午; Excel's Range.Text, measured)", () => {
  for (const x of EXCEL.codes) assert.equal(shown(x.code, x.value), x.text, `${x.code} ${x.value}`);
  assert.equal(EXCEL.codes.length, 840);
});

test('the ids alone in en-US are what Excel reads them as: the short date, or h:mm:ss (measured)', () => {
  for (const x of EXCEL.ids) {
    assert.equal(builtinFormat(x.id, EN_US), x.code, `id ${x.id}`);
    assert.equal(shown(builtinFormat(x.id, EN_US), x.value), x.text, `id ${x.id} ${x.value}`);
  }
  assert.equal(EXCEL.ids.length, 19 * 14);
  // Outside East Asia they follow the short date as format 14 does.
  assert.equal(shown(builtinFormat(27, localeOf('de-DE')), 45356, localeOf('de-DE')), '05.03.2024');
});

test('ja-JP, zh-CN, zh-TW and ko-KR read each id as their own code, drawn as Excel draws it', () => {
  const tags = { 'ja-JP': 'ja', 'zh-CN': 'zh-CN', 'zh-TW': 'zh-TW', 'ko-KR': 'ko', 'zh-HK': 'zh-TW', 'zh-SG': 'zh-CN' };
  for (const [tag, key] of Object.entries(tags)) {
    const loc = localeOf(tag);
    for (const id of EAST_ASIAN_IDS) {
      const code = builtinFormat(id, loc);
      assert.equal(code, EAST_ASIAN_FORMATS[key][id], `${tag} ${id}`);
      for (const x of EXCEL.codes.filter((c) => c.code === code)) assert.equal(shown(code, x.value, loc), x.text, `${tag} ${id} ${x.value}`);
    }
  }
  const ja = localeOf('ja-JP');
  assert.equal(shown(builtinFormat(28, ja), 43586, ja), '令和1年5月1日');
  assert.equal(shown(builtinFormat(27, ja), 32515, ja), 'S64.1.7');
  assert.equal(shown(builtinFormat(34, localeOf('zh-CN')), 0.5625, localeOf('zh-CN')), '下午1时30分');
  // Other ids are as they were.
  assert.equal(builtinFormat(14, ja), 'm/d/yyyy');
  assert.equal(builtinFormat(49, ja), '@');
});

test('e and g outside a Japanese code: e is the year, g nothing; E+ is still scientific', () => {
  assert.equal(shown('e/m/d', 45356), '2024/3/5');
  assert.equal(shown('gge', 45356), '2024');
  assert.equal(shown('0.0E+0', 12345), '1.2E+4');
  assert.equal(shown('[$-ja-JP]ggge', 45356), '令和6');
});
