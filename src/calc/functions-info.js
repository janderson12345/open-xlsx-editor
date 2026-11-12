// CELL and INFO as Excel 365 (English) gives them (measured over COM, docs/CALC.md). Both are worked out again at every
// recalculation, as Excel's are.
import { Area, Arr, E } from './values.js';
import { def, given, text } from './functions.js';
import { formatLetters } from './cellinfo.js';
import { colName } from '../ooxml/refs.js';

def('CELL', (a, ctx) => {
  const kind = text(a[0], ctx).toLowerCase();
  // (Without a reference, the cell the formula is in; Excel takes the active cell then.)
  const ref = given(a[1]) ? a[1] : new Area(ctx.here.si, ctx.here.r, ctx.here.c, ctx.here.r, ctx.here.c);
  if (!(ref instanceof Area)) throw E.value;
  const { si, r0: r, c0: c } = ref;
  const info = () => ctx.cellInfo(si, r, c);
  switch (kind) {
    case 'address': {
      const at = `$${colName(c)}$${r + 1}`;
      if (si === ctx.here.si) return at;
      const name = `[${ctx.bookName?.() || 'Book1'}]${ctx.sheetName(si)}`;
      return `${/^[A-Za-z0-9_.[\]]+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`}!${at}`;
    }
    case 'row': return r + 1;
    case 'col': return c + 1;
    case 'contents': { const v = ctx.cellValue(si, r, c); return v ?? 0; }
    case 'type': { const v = ctx.cellValue(si, r, c); return v == null ? 'b' : typeof v === 'string' ? 'l' : 'v'; }
    case 'prefix': return info().prefix;
    case 'protect': return info().locked ? 1 : 0;
    case 'width': { const i = info(); return new Arr([[i.width, i.defaultWidth]]); }
    case 'format': return formatLetters(info().code).format;
    case 'color': return formatLetters(info().code).color;
    case 'parentheses': return formatLetters(info().code).parentheses;
    // (Excel gives the saved file's full path; a page has none, as Excel has none before a workbook is saved.)
    case 'filename': return '';
    default: throw E.value;
  }
}, { volatile: true });

def('INFO', (a, ctx) => {
  switch (text(a[0], ctx).toLowerCase()) {
    case 'numfile': return ctx.sheetCount();
    case 'origin': return '$A:$A$1';
    case 'osversion': return 'Windows (64-bit) NT 10.00';
    case 'recalc': return ctx.calcMode?.() === 'manual' ? 'Manual' : 'Automatic';
    case 'release': return '16.0';
    case 'system': return 'pcdos';
    case 'memavail': case 'memused': case 'totmem': throw E.na;
    default: throw E.value;
  }
}, { volatile: true });
