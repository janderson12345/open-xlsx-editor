// Runs every test/*.test.js with Node's own runner. (npm runs scripts through cmd.exe on Windows, which doesn't expand
// `test/*.test.js`, and Node before 21 doesn't either; listing the files here works on every shell and Node 18+.)
import fs from 'node:fs';
import { spawnSync } from 'node:child_process';

const files = fs.readdirSync('test').filter((f) => f.endsWith('.test.js')).map((f) => `test/${f}`);
const r = spawnSync(process.execPath, ['--test', ...files], { stdio: 'inherit' });
process.exit(r.status ?? 1);
