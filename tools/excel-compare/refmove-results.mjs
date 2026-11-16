// Excel's results from excel-refmoves.ps1 or excel-slicers.ps1 cut down to what the tests compare, with each case's
// step: test/refmoves-excel.json and test/slicers-excel.json. (No ids, no names: the parts' revision ids and Excel's
// user name stay out.)
// usage: node tools/excel-compare/refmove-results.mjs refmoves <cases.json> <excel out.json> > test/refmoves-excel.json
//        node tools/excel-compare/refmove-results.mjs slicers <cases.json> <excel out.json> > test/slicers-excel.json
import fs from 'node:fs';

const [kind, casesPath, outPath] = process.argv.slice(2);
const cases = JSON.parse(fs.readFileSync(casesPath, 'utf8')).cases;
const results = JSON.parse(fs.readFileSync(outPath, 'utf8').replace(/^\uFEFF/, ''));
const unescape = (s) => s.replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const list = (x) => (x == null ? [] : Array.isArray(x) ? x : [x]);

/** What the refmoves tests compare: sparklines (cell <- data, in the file's order), rules ("sqref | formula1 |
 *  formula2", sorted: Excel's own order of them shifts), notes, threads, the notes' boxes, the ovals' cells. */
function refmove(c, k) {
  const x = c.xml;
  const formulas = (body, tag) => [...body.matchAll(new RegExp(`<${tag}(\\d)>(?:<xm:f>)?([^<]*)`, 'g'))].map((m) => unescape(m[2]));
  const rules = [...(x.dataValidations || '').matchAll(/<dataValidation\b[^>]*\ssqref="([^"]*)"[^>]*>([\s\S]*?)<\/dataValidation>/g)]
    .map((m) => [m[1], ...formulas(m[2], 'formula')].join(' | ')).sort();
  const x14 = [...(x.x14DataValidations || '').matchAll(/<x14:dataValidation\b[^>]*>([\s\S]*?)<\/x14:dataValidation>/g)]
    .map((m) => [/<xm:sqref>([^<]*)/.exec(m[1])[1], ...formulas(m[1], 'x14:formula')].join(' | ')).sort();
  return {
    name: c.name, step: cases[k].step, refused: c.refused,
    sparklines: [...(x.sparklines || '').matchAll(/<xm:f>([^<]*)<\/xm:f><xm:sqref>([^<]*)<\/xm:sqref>/g)].map((m) => `${m[2]} <- ${unescape(m[1])}`),
    rules, x14,
    notes: list(x.commentRefs).sort(), threads: list(x.threadRefs).sort(), boxes: list(x.vmlNotes).sort(),
    ovals: list(c.ovals).map((o) => o.split(' ')[0]),
  };
}

/** What the slicer tests compare: each object's anchor in the drawing, the slicer caches and what's left of the
 *  slicer parts, the table and the pivot. */
function slicer(c, k) {
  const p = c.parts;
  const drawing = p['xl/drawings/drawing1.xml'] || '';
  const anchors = {};
  for (const m of drawing.matchAll(/<xdr:twoCellAnchor editAs="(\w+)">([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
    const pt = (which) => /<xdr:col>(\d+)<\/xdr:col><xdr:colOff>(-?\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(-?\d+)<\/xdr:rowOff>/.exec(new RegExp(`<xdr:${which}>([\\s\\S]*?)</xdr:${which}>`).exec(m[2])[1]).slice(1).join(',');
    const name = /<(?:sle|tsle):(?:slicer|timeslicer)\b[^>]*name="([^"]*)"/.exec(m[2])[1];
    anchors[name] = `${m[1]} ${pt('from')} ${pt('to')}`;
  }
  const caches = Object.entries(p).filter(([path]) => /^xl\/(slicerCaches|timelineCaches)\//.test(path)).map(([, text]) => {
    const name = /\sname="([^"]*)"/.exec(text)[1];
    const tc = /<x15:tableSlicerCache tableId="(\d+)" column="(\d+)"/.exec(text);
    return tc ? `${name} table ${tc[1]} column ${tc[2]}` : name;
  }).sort();
  const slicers = Object.entries(p).filter(([path]) => /^xl\/(slicers|timelines)\//.test(path)).flatMap(([, text]) => [...text.matchAll(/<(?:slicer|timeline) name="([^"]*)"/g)].map((m) => m[1])).sort();
  const table = Object.entries(p).find(([path]) => /^xl\/tables\//.test(path))?.[1];
  const pivot = Object.entries(p).find(([path]) => /^xl\/pivotTables\//.test(path))?.[1];
  return {
    name: c.name, step: cases[k].step, refused: c.refused, anchors, caches, slicers,
    table: table ? `${/ ref="([^"]*)"/.exec(table)[1]} ${[...table.matchAll(/<tableColumn id="(\d+)"[^>]*name="([^"]*)"/g)].map((m) => `${m[1]}:${m[2]}`).join(',')}` : null,
    pivot: pivot ? /<location ref="([^"]*)"/.exec(pivot)[1] : null,
  };
}

const out = results.map((c, k) => (kind === 'slicers' ? slicer(c, k) : refmove(c, k)));
console.log(JSON.stringify(out, null, 1));
