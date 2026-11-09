import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const run = (...args) => spawnSync(process.execPath, ['tools/check-ooxml.mjs', ...args], { encoding: 'utf8' });
const dir = 'test/out/check';

/** basic.xlsx with one part's text changed. */
function broken(name, part, change) {
  const files = unzipSync(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  if (change === null) delete files[part];
  else files[part] = strToU8(change(strFromU8(files[part])));
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(`${dir}/${name}.xlsx`, zipSync(files));
  return `${dir}/${name}.xlsx`;
}

test('the fixtures pass', () => {
  const r = run('test/fixtures');
  assert.equal(r.status, 0, r.stdout);
  assert.match(r.stdout, /all \d+ files ok/);
});

const cases = [
  ['not well-formed', 'xl/styles.xml', (s) => s.replace('</styleSheet>', ''), /not well-formed/],
  ['missing target', 'xl/styles.xml', null, /rId4 points to missing xl\/styles.xml|override for missing part/],
  ['undefined r:id', 'xl/workbook.xml', (s) => s.replace('r:id="rId2"', 'r:id="rId99"'), /rId99.*not in its relationships/],
  ['rows out of order', 'xl/worksheets/sheet1.xml', (s) => s.replace('<row r="40">', '<row r="3">'), /row 3 after row 15/],
  ['cell in the wrong row', 'xl/worksheets/sheet1.xml', (s) => s.replace('<c r="A40">', '<c r="A41">'), /A41 inside row 40/],
  ['shared string out of range', 'xl/worksheets/sheet2.xml', (s) => s.replace(/t="s"><v>\d+</, 't="s"><v>999<'), /shared string 999/],
  ['style out of range', 'xl/worksheets/sheet1.xml', (s) => s.replace('<c r="A1" t="s" s="1">', '<c r="A1" t="s" s="7">'), /style s="7", cellXfs has 2/],
  ['overlapping merges', 'xl/worksheets/sheet1.xml', (s) => s.replace('<mergeCell ref="A12:C13"/>', '<mergeCell ref="A12:C13"/><mergeCell ref="B13:D14"/>'), /overlap/],
  ['duplicate sheet name', 'xl/workbook.xml', (s) => s.replace('name="Hidden"', 'name="data"'), /two sheets named "data"/],
  ['bad sheet name', 'xl/workbook.xml', (s) => s.replace('name="Hidden"', 'name="a/b"'), /character Excel refuses/],
  ['no content type', '[Content_Types].xml', (s) => s.replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*>/, '').replace('<Default Extension="xml" ContentType="application/xml"/>', ''), /no content type/],
];
for (const [name, part, change, expect] of cases) {
  test(`caught: ${name}`, () => {
    const r = run(broken(name.replace(/\W+/g, '-'), part, change));
    assert.equal(r.status, 1, r.stdout);
    assert.match(r.stdout, expect);
  });
}
