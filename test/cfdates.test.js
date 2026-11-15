import './setup.js';
import { test, mock } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { build, area } from './steps.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { conditionalLooks } from '../src/ooxml/condformat.js';
import { CF_PERIODS } from '../src/ooxml/cfedit.js';

// Highlight Cells Rules > A Date Occurring, against Excel (test/cf-dates-excel.json, excel-cf-dates.ps1): columns A to J
// each the dates from 45 days before Excel's today to 45 after, one period a column; each rule written as Excel wrote
// it, and the cells it highlights those Excel highlighted, with the clock at Excel's today.
const EXCEL = JSON.parse(fs.readFileSync('test/cf-dates-excel.json', 'utf8').trim());
const PERIODS = Object.keys(CF_PERIODS);
const COL = 'ABCDEFGHIJ';
const rule = (xml) => /<cfRule\b[\s\S]*?<\/cfRule>/.exec(xml)[0].replace(/ stopIfTrue="1"/, '').replace(/ priority="\d+"/, '').replace(/ dxfId="\d+"/, ' dxfId');

test('A Date Occurring: the ten periods written and drawn as Excel writes and draws them', () => {
  assert.deepEqual(PERIODS, Object.keys(EXCEL.periods), "the menu's order");
  // (Noon of Excel's today, in local time: TODAY() reads the local date.)
  const today = new Date(1899, 11, 30 + EXCEL.today, 12);
  mock.timers.enable({ apis: ['Date'], now: today.getTime() });
  try {
    const cells = {};
    for (let i = 0; i < 91; i++) for (const c of COL) cells[`${c}${i + 1}`] = String(EXCEL.today - 45 + i);
    const { wb, si } = build({ cells });
    PERIODS.forEach((period, j) => wb.setConditional(si, [area(`${COL[j]}1:${COL[j]}91`)], { kind: 'timePeriod', period, preset: 'lightRed' }));
    const root = wb._part(si).doc.documentElement;
    const written = [...root.childNodes].filter((n) => n.localName === 'conditionalFormatting').map((n) => serializeXml(n).replace(/ xmlns(:\w+)?="[^"]*"/g, ''));
    const bySqref = (list) => Object.fromEntries(list.map((x) => [/sqref="([^"]*)"/.exec(x)[1], rule(x)]));
    assert.deepEqual(bySqref(written), bySqref(EXCEL.conditionalFormatting), 'the rules');
    const looks = conditionalLooks(wb, si);
    PERIODS.forEach((period, j) => {
      const hits = [];
      for (let i = 0; i < 91; i++) if (looks.get(`${i},${j}`)) hits.push(i - 45);
      assert.deepEqual(hits, EXCEL.periods[period], period);
    });
  } finally {
    mock.timers.reset();
  }
});
