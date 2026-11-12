import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';

// (Excel 365 over COM: FILTERXML's results, the first value and those spilled under it to the sixth, each as S: text,
// N: number or boolean, E: an error's code.)
const EXCEL = [
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i\")", "N:10 | S:x | N:2.5 | S:deep"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"/r/i\")", "N:10 | S:x | N:2.5"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"/r/i[1]\")", "N:10"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[2]\")", "S:x"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[@a]\")", "N:10 | S:x"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[@a='2']\")", "S:x"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i/@a\")", "N:1 | N:2"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//g/i\")", "S:deep"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//n\")", "N:True"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//e\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//d\")", "N:45296"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//zz\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"count(//i)\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"sum(//i[@a='1'])\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[last()]\")", "N:2.5 | S:deep"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[position()>2]\")", "N:2.5"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[.>5]\")", "N:10"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[contains(.,'x')]\")", "S:x"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"/r/*\")", "N:10 | S:x | N:2.5 | E:-2146826273 | N:True | E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i/text()\")", "E:-2146826273 | E:-2146826273 | E:-2146826273 | E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"string(//i)\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i | //n\")", "N:10 | S:x | N:2.5 | S:deep | N:True"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"name(/r/*[1])\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[not(@a)]\")", "N:2.5 | S:deep"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[@a!='1']\")", "S:x"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//I\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[1]/following-sibling::i\")", "S:x | N:2.5"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//g/..\")", "E:-2146826273"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"/r/i[2]/@a\")", "N:2"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"//i[starts-with(.,'d')]\")", "S:deep"],
  ["FILTERXML(\"<r><i a=\"\"1\"\">10</i><i a=\"\"2\"\">x</i><i>2.5</i><g><i>deep</i></g><n>TRUE</n><e/><d>2024-01-05</d></r>\",\"concat(//i[1],'-',//i[2])\")", "E:-2146826273"],
  ["FILTERXML(\"<a>x</a>\",\"/a\")", "S:x"],
  ["FILTERXML(\"<a>x\",\"/a\")", "E:-2146826273"],
  ["FILTERXML(\"<a>x</a>\",\"/b\")", "E:-2146826273"],
  ["FILTERXML(\"<a>x</a>\",\"//\")", "E:-2146826273"],
  ["FILTERXML(\"<a>x</a>\",\"\")", "E:-2146826273"],
  ["FILTERXML(\"\",\"/a\")", "E:-2146826273"],
  ["FILTERXML(\"<a> 12 </a>\",\"/a\")", "N:12"],
  ["FILTERXML(\"<a>1,234</a>\",\"/a\")", "N:1234"],
  ["FILTERXML(\"<a>1e3</a>\",\"/a\")", "N:1000"],
  ["FILTERXML(\"<a></a>\",\"/a\")", "E:-2146826273"],
  ["FILTERXML(\"<a>50%</a>\",\"/a\")", "N:0.5"],
  ["FILTERXML(\"<a>1/2/2024</a>\",\"/a\")", "N:45293"],
  ["FILTERXML(\"<a><b>1</b><b>2</b></a>\",\"//b\")", "N:1 | N:2"],
  ["ROWS(FILTERXML(\"<a><b>1</b><b>2</b></a>\",\"//b\"))", "N:2"],
  ["COLUMNS(FILTERXML(\"<a><b>1</b><b>2</b></a>\",\"//b\"))", "N:1"],
  ["FILTERXML(\"<a>&amp;&lt;</a>\",\"/a\")", "S:&<"],
  ["FILTERXML(\"<a><![CDATA[q]]></a>\",\"/a\")", "S:q"],
  ["FILTERXML(\"<a>x<b>y</b>z</a>\",\"/a\")", "S:xz"],
  ["FILTERXML(\"<x:a xmlns:x='u'>1</x:a>\",\"//*\")", "N:1"],
  ["FILTERXML(\"<a>TRUE</a>\",\"/a\")", "N:True"],
  ["FILTERXML(\"<a>#N/A</a>\",\"/a\")", "S:#N/A"],
  ["FILTERXML(12,\"/a\")", "E:-2146826273"],
  ["FILTERXML(\"<a b='1'/>\",\"/a/@b\")", "N:1"],
  ["FILTERXML(\"<a>x</a>\",\"/a/text()\")", "E:-2146826273"],
];
const CODES = { '#VALUE!': -2146826273, '#N/A': -2146826246 };
const show = (v) => (v && v.code ? `E:${CODES[v.code] ?? v.code}` : typeof v === 'string' ? `S:${v}` : v === true ? 'N:True' : v === false ? 'N:False' : `N:${v}`);

test('FILTERXML: 55 queries, paths, predicates, axes and the values they give, as Excel', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  for (const [f, want] of EXCEL) {
    const o = wb.calc.formulaAt(0, 60, 0, f, true);
    const rows = o?.rows ?? (Array.isArray(o) ? o : [[o]]);
    assert.equal(rows.slice(0, 6).map((r) => show(r[0])).join(' | '), want, f);
  }
});
