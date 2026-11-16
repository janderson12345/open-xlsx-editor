// Made-up text files and Text Import settings for excel-textimport.ps1 (Workbooks.OpenText, Excel's Text Import
// Wizard over COM) and test/textimport.test.js. Writes each file's bytes into `dir` (outside the repo) and the cases,
// with their paths, to `out` (outside the repo too: the paths are this computer's).
// usage: node tools/excel-compare/textimport-cases.mjs ..\probe\textimport ..\probe\textimport\cases.json
import fs from 'node:fs';
import path from 'node:path';
import { TEXT_IMPORT_CASES, caseBytes } from '../../test/textimport-cases.js';

const [dir, out] = process.argv.slice(2);
fs.mkdirSync(dir, { recursive: true });
const cases = TEXT_IMPORT_CASES.map((c) => {
  const file = path.resolve(dir, `${c.name}.txt`);
  fs.writeFileSync(file, caseBytes(c));
  return { ...c, file };
});
fs.writeFileSync(out, `${JSON.stringify(cases, null, 1)}\n`);
console.log(`${cases.length} cases`);
