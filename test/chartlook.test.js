import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chartParts } from '../src/ooxml/chartedit.js';
import { readCharts, valueAxis } from '../src/ooxml/charts.js';
import { excel, withChart } from './chartformat-cases.js';

// Excel's own formatted charts (test/chartformat-excel.json, excel-chart-format.ps1) read here: each chart part Excel
// saved put in place of the chart inserted here, then read as charts.js reads a file.
function excelChart(name) {
  const wb = withChart(name);
  wb.pkg.setText(chartParts(wb, 0)[0].path, excel.charts[name]);
  return readCharts(wb, 0)[0];
}
const font = (name, size, bold, rgb) => ({ name, size, bold, color: { rgb } });

test('Excel\'s axis titles, fills, lines and fonts read from its parts', () => {
  assert.deepEqual([excelChart('catTitle').catTitle.text, excelChart('valTitle').axis.title.text, excelChart('secondaryTitle').axis2.title.text], ['Region', 'Amount', 'Cost']);
  assert.deepEqual([excelChart('catTitleAuto').catTitle.text, excelChart('catTitleAuto').catTitle.own], ['Axis Title', false]);
  assert.deepEqual([excelChart('titlesOff').catTitle, excelChart('titlesOff').axis.title], [null, null]);
  assert.deepEqual(excelChart('seriesFill').series[0].look.fill, { rgb: 'C00000' });
  assert.deepEqual(excelChart('seriesLine').series[0].look.line, { rgb: '7030A0' });
  assert.deepEqual(excelChart('seriesNoLine').series[0].look.line, { none: true });
  assert.deepEqual(excelChart('pointFill').series[0].look.points.get(1).fill, { rgb: '00B050' });
  assert.deepEqual(excelChart('pieSlice').series[0].look.points.get(1).fill, { rgb: '00B050' });
  assert.deepEqual(excelChart('lineColor').series[0].look.line, { rgb: 'C00000' });
  assert.deepEqual([excelChart('chartFill').look.area.fill, excelChart('chartLine').look.area.line, excelChart('chartNoLine').look.area.line], [{ rgb: 'FFF2CC' }, { rgb: '1F4E79' }, { none: true }]);
  assert.deepEqual([excelChart('plotFill').look.plot.fill, excelChart('plotLine').look.plot.line], [{ rgb: 'DDEBF7' }, { rgb: '808080' }]);
  // (The default style's own: the area white, lined in 15% grey; the title 14pt in 65% grey in the theme's body font.)
  const none = excelChart('none');
  assert.deepEqual([none.look.area.fill, none.look.area.line], [{ scheme: 'bg1' }, { scheme: 'tx1', lumMod: 0.15, lumOff: 0.85 }]);
  assert.deepEqual(none.look.fonts.title, { name: '+mn-lt', size: 14, bold: false, color: { scheme: 'tx1', lumMod: 0.65, lumOff: 0.35 } });
  // (Fonts by Excel's Font object and by TextFrame2, which writes the title's text and the run's own properties.)
  assert.deepEqual(excelChart('titleFont1').look.fonts.title, font('Arial', 18, true, 'C00000'));
  assert.deepEqual(excelChart('titleTextFont').look.fonts.title, font('Arial', 18, true, 'C00000'));
  assert.deepEqual(excelChart('titleTextFontText').look.fonts.title, font('Arial', 18, true, 'C00000'));
  assert.deepEqual([excelChart('catFont').catFont, excelChart('valFont').axis.font], [font('Arial', 11, true, '1F4E79'), font('Arial', 11, true, '1F4E79')]);
  assert.deepEqual([excelChart('legendFont1').look.fonts.legend, excelChart('labelFont1').look.fonts.labels], [font('Arial', 11, true, '1F4E79'), font('Arial', 11, true, '1F4E79')]);
  assert.deepEqual([excelChart('catTitleFont1').catTitle.font, excelChart('catTitleFont').catTitle.font], [font('Arial', 12, true, 'C00000'), font('Arial', 12, true, 'C00000')]);
});

test('combo charts and the secondary axis read as Excel reads them (Series.ChartType, AxisGroup and the axes\' scales)', () => {
  // (Excel's chart types: 51 clustered columns, 4 line, 65 line with markers, 1 area, 5 pie.)
  const TYPES = { 51: 'bar', 4: 'line', 65: 'line', 1: 'area', 5: 'pie' };
  const wrong = [];
  for (const [name, list] of Object.entries(excel.types)) {
    const c = excelChart(name);
    const got = c.series.map((s) => `${s.type}/${s.secondary ? 2 : 1}${s.marker && s.marker !== 'none' ? '+' : ''}`).join(' ');
    const want = list.map((s) => `${TYPES[s.type]}/${s.group}${s.type === 65 ? '+' : ''}`).join(' ');
    if (got !== want) wrong.push(`${name}: Excel ${want}, here ${got}`);
    const scale = excel.secondary[name];
    if (scale) {
      const [two, one] = [valueAxis(c, true), valueAxis(c)];
      if (two.min !== scale.min || two.max !== scale.max || two.major !== scale.major) wrong.push(`${name}: Excel's secondary axis ${scale.min} to ${scale.max} by ${scale.major}, here ${two.min} to ${two.max} by ${two.major}`);
      if ([one.min, one.max, one.major].join() !== scale.primary.join()) wrong.push(`${name}: Excel's primary axis ${scale.primary}, here ${[one.min, one.max, one.major]}`);
    }
  }
  assert.deepEqual(wrong, []);
  assert.equal(Object.keys(excel.types).length, 44);
});
