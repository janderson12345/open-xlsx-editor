// 3D references (=SUM(Jan:Apr!A1)) and the intersection operator (=A1:C3 B2:D4), as Excel 365 works them out,
// measured with a hidden Excel on 86 formulas over four sheets. evaluate.js calls these.
import { Area, Ref3D, E, isError } from './values.js';

/** A 3D reference's areas, one a sheet from its first to its last in the workbook's order (Apr:Jan as Jan:Apr); one
 *  sheet, a plain reference. A sheet that isn't there: #REF!. */
export function threeD(node, ctx) {
  const [s0, s1] = node.sheets.map((s) => ctx.sheetIndex(s));
  if (s0 < 0 || s1 < 0) return E.ref;
  const a = node.area;
  if (s0 === s1) return new Area(s0, a.r0, a.c0, a.r1, a.c1);
  const areas = [];
  for (let si = Math.min(s0, s1); si <= Math.max(s0, s1); si++) areas.push(new Area(si, a.r0, a.c0, a.r1, a.c1));
  return new Ref3D(areas);
}

/** The intersection of two references (=A1:C3 B2:D4 is B2:C3): #NULL! when they don't meet, #VALUE! for references
 *  on different sheets, a 3D one or what isn't a reference. */
export function intersect(a, b) {
  if (isError(a)) return a;
  if (isError(b)) return b;
  if (!(a instanceof Area) || !(b instanceof Area) || a.si !== b.si) return E.value;
  const r0 = Math.max(a.r0, b.r0);
  const c0 = Math.max(a.c0, b.c0);
  const r1 = Math.min(a.r1, b.r1);
  const c1 = Math.min(a.c1, b.c1);
  return r0 > r1 || c0 > c1 ? E.null : new Area(a.si, r0, c0, r1, c1);
}

// The functions that take a 3D reference in Excel 365 (Microsoft's list and more, measured): each sheet's area read
// in turn.
const TAKES = new Set(['SUM', 'AVERAGE', 'AVERAGEA', 'COUNT', 'COUNTA', 'MAX', 'MAXA', 'MIN', 'MINA', 'PRODUCT',
  'STDEV', 'STDEVA', 'STDEVP', 'STDEVPA', 'STDEV.S', 'STDEV.P', 'VAR', 'VARA', 'VARP', 'VARPA', 'VAR.S', 'VAR.P',
  'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'SUMSQ', 'MEDIAN', 'LARGE', 'SMALL', 'PERCENTILE', 'QUARTILE', 'AVEDEV',
  'GEOMEAN', 'HARMEAN', 'DEVSQ', 'SKEW', 'KURT', 'CONCAT', 'TEXTJOIN', 'AND', 'OR', 'ISREF']);
// (Those that pass a value on, or ask what it is: SUM(CHOOSE(1,Jan:Apr!A1)) is a sum, IFERROR(Jan:Apr!A1,"e") "e".)
const PASSES = new Set(['IF', 'IFS', 'CHOOSE', 'SWITCH', 'IFERROR', 'IFNA', 'ISERROR', 'ISERR', 'ISNA', 'ISNUMBER',
  'ISTEXT', 'ISNONTEXT', 'ISLOGICAL', 'ISBLANK', 'ERROR.TYPE']);
// (Those that take a reference of one sheet: #VALUE!. Any other function: #REF!, as SUMPRODUCT and MODE are.)
const ONE_SHEET = new Set(['ROWS', 'COLUMNS', 'INDEX', 'N', 'T', 'COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF',
  'AVERAGEIFS', 'MAXIFS', 'MINIFS', 'COUNTBLANK', 'AGGREGATE', 'SUBTOTAL', 'OFFSET', 'AREAS', 'CELL', 'ROW', 'COLUMN']);

/** A function given a 3D reference that it doesn't take: Excel's error for it; null when it takes it. */
export function refused3D(name, args) {
  if (!args.some((x) => x instanceof Ref3D)) return null;
  const plain = name.replace(/^_XLFN\./, '');
  if (TAKES.has(plain) || PASSES.has(plain) || /^(PERCENTILE|QUARTILE)\./.test(plain)) return null;
  return ONE_SHEET.has(plain) ? E.value : E.ref;
}

/** A 3D reference as a cell's answer: #REF! written so (=Jan:Apr!A1), #VALUE! given by a function (=CHOOSE(1,...)). */
export const answer3D = (tree) => (tree.k === 'ref' || (tree.k === 'paren' && tree.a.k === 'ref') ? E.ref : E.value);
