// What each listed function and its arguments do, for the Function Arguments box: one sentence each, written for this
// project (describe-*.js), not Microsoft's text. Every function in calc/catalog.js has one (test/fnargs.test.js
// checks); an argument takes its function's own sentence ("SUM.number") if there is one, else its name's.
import { FINANCIAL, DATES } from './describe-money.js';
import { MATH, LOOKUP } from './describe-math.js';
import { STATISTICAL, COMPATIBILITY, DATABASE } from './describe-stats.js';
import { TEXT, LOGICAL, INFORMATION, ENGINEERING, WEB } from './describe-text.js';
import { BY_NAME, BY_FUNCTION } from './describe-args.js';

/** "KEY: sentence" lines into a map. */
const table = (...texts) => {
  const map = new Map();
  for (const line of texts.join('\n').split('\n')) {
    const at = line.indexOf(': ');
    if (at > 0) map.set(line.slice(0, at).trim(), line.slice(at + 2).trim());
  }
  return map;
};
const FUNCTIONS = table(FINANCIAL, DATES, MATH, LOOKUP, STATISTICAL, COMPATIBILITY, DATABASE, TEXT, LOGICAL, INFORMATION, ENGINEERING, WEB);
const ARGS = table(BY_NAME);
const OWN = table(BY_FUNCTION);

/** What the function does, or ''. */
export const describeFunction = (name) => FUNCTIONS.get(String(name).toUpperCase()) || '';

/** What argument `arg` (as the box names it: "number3", "criteria_range2") of function `name` is, or ''. */
export function describeArg(name, arg) {
  const base = String(arg).replace(/\d+$/, '');
  const fn = String(name).toUpperCase();
  return OWN.get(`${fn}.${base}`) || OWN.get(`${fn}.${arg}`) || ARGS.get(base) || ARGS.get(arg) || '';
}
