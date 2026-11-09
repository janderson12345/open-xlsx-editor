// Finds bytes that shouldn't be in source text: control characters other than tab, LF and CR, and byte order marks.
// (Inline shell scripts on this machine turn `\\` into `\`, which once put 0x08 and 0x02 into regexes; LESSONS.md
// section 1.) usage: node tools/scan-bytes.mjs [dir...]   (default: src test tools demo)
import fs from 'node:fs';
import path from 'node:path';

const roots = process.argv.slice(2).length ? process.argv.slice(2) : ['src', 'test', 'tools', 'demo'];
const skip = new Set(['node_modules', 'dist', 'out', '.git']);
const text = /\.(js|mjs|cjs|ts|css|html|json|md|ps1|py|yml)$/;
let bad = 0;
const walk = (dir) => {
  if (!fs.existsSync(dir)) return;
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) { if (!skip.has(e.name)) walk(p); continue; }
    if (!text.test(e.name)) continue;
    const b = fs.readFileSync(p);
    for (let i = 0; i < b.length; i++) {
      const x = b[i];
      const bom = x === 0xef && b[i + 1] === 0xbb && b[i + 2] === 0xbf;
      if ((x < 32 && x !== 9 && x !== 10 && x !== 13) || bom) {
        const line = b.subarray(0, i).toString('utf8').split('\n').length;
        console.log(`${p}:${line}: ${bom ? 'byte order mark' : `control character 0x${x.toString(16).padStart(2, '0')}`}`);
        bad++;
      }
    }
  }
};
roots.forEach(walk);
if (bad) process.exit(1);
console.log('no stray control characters or byte order marks');
