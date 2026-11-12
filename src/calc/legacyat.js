// The @s Excel 365 shows in a formula of the old kind (saved without dynamic-array metadata, or by an older Excel):
// where it cuts a range or an array to one value by implicit intersection. Measured with a hidden Excel, a formula set
// by Range.Formula read back by Formula2 (38 formulas): a range where one value is wanted, an operand or a value
// argument, =@A1:A3*2, =SUM(@A1:A3*2), =VLOOKUP(@A1:A3,A1:B3,2,FALSE), =COUNTIF(A1:A3,@A1:A3); a function giving
// several cells there, =@OFFSET(A1,0,0,2), =@CHOOSE(1,A1:A3), =@INDEX(A1:A3,0,1); none where the range is taken whole,
// SUMPRODUCT, MATCH, ROWS, SUMIF's. Found by working the formula out as the old kind at its cell and marking each part
// whose range or array the evaluator cut to one value. The formula bar shows it so; typed back, the @s keep its meaning.
import { parseFormula } from './parse.js';
import { result } from './evaluate.js';

/** `shown` (a formula as the formula bar shows it, without "=") with Excel 365's @s, worked out at cell (r, c) of sheet
 *  si by the engine `calc`; as it was when it can't be worked out here. */
export function legacyAt(calc, si, r, c, shown) {
  let tree;
  try { tree = parseFormula(shown); } catch { return shown; }
  const ctx = calc.context({ si, r, c });
  const marks = new Set();
  ctx.origin = new Map();
  ctx.shrunk = (v) => { const n = ctx.origin.get(v); if (n?.at != null) marks.add(n.at); };
  try { result(tree, ctx, false); } catch { return shown; }
  // (Where Excel's functions take their arguments otherwise than the evaluator does, as measured: ROW and COLUMN of
  // several cells give an array, @ROW(A1:A3); N takes a reference, N(A1:A3); TRANSPOSE's range is cut too,
  // @TRANSPOSE(@A1:A3).)
  const multi = (n) => n?.k === 'ref' && n.area && (n.area.kind === 'cols' || n.area.kind === 'rows' || n.area.r0 !== n.area.r1 || n.area.c0 !== n.area.c1);
  const walk = (n) => {
    if (!n || typeof n !== 'object') return;
    if (n.k === 'fn') {
      if ((n.name === 'ROW' || n.name === 'COLUMN') && multi(n.args[0])) marks.add(n.at);
      if (n.name === 'N' && multi(n.args[0])) marks.delete(n.args[0].at);
      if (n.name === 'TRANSPOSE' && multi(n.args[0])) marks.add(n.args[0].at);
    }
    for (const x of [n.a, n.b, ...(n.args || []), ...(n.items || [])]) walk(x);
  };
  walk(tree);
  let out = shown;
  for (const at of [...marks].sort((a, b) => b - a)) out = `${out.slice(0, at)}@${out.slice(at)}`;
  return out;
}
