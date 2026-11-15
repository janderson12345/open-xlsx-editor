// Headers and footers on a printed page, as Excel prints them (print.js draws the rest of the page): their sections'
// text, fields filled in, and their pictures (&G, ooxml/hfpictures.js). Measured against Excel's PDFs (docs/PRINT.md).
import { setVars, div } from './dom.js';
import { fontStack, fontRowPx } from './fonts.js';
import { printRowPx } from '../ooxml/units.js';
import { parseHeaderFooter, fillFields } from '../ooxml/headerfooter.js';

const DPI = 96;
/** A header or footer picture prints larger than its VML size, by these across and down (Excel's PDFs: 60 by 30pt as
 *  85.92 by 40.64px alone, 40.5 by 20.25pt as 58.08 by 27.36px), PIC_BESIDE px wider still for text before it on its
 *  line and as much for text after it (87.84px after "Title ", 89.76px between "Foot " and " end"), and on a line with
 *  any characters (its line break too) with its bottom this many em below the line's baseline (2.6 to 2.7px at 11pt).
 *  All with the sheet's scale when the header scales with it (half at 50%). */
const PIC_W = 1.075;
const PIC_H = 1.015;
const PIC_DROP = 0.183;
const PIC_BESIDE = 1.92;
/** In a section of several lines, a line's baseline this many em below its top (a header's first: 1.014em below the
 *  margin), and a footer's last 0.361em above the footer margin. */
const LINE_IN = 1.014;
const FOOT_LAST = 0.361;

/** A run that prints something other than a picture. */
const words = (r) => !r.picture && (r.field || r.text?.trim());

/** Runs (fields filled in) as lines: a text run with line breaks split at them. */
export function linesOf(runs) {
  const out = [[]];
  for (const r of runs) {
    if (r.picture || !r.text.includes('\n')) { out.at(-1).push(r); continue; }
    r.text.split('\n').forEach((t, i) => { if (i) out.push([]); if (t) out.at(-1).push({ ...r, text: t }); });
  }
  return out;
}

/** A section's runs (one line of it, for a section with a picture) into `el`: spans, and the picture `pic` where &G
 *  is, wider for words beside it. */
function appendRuns(el, runs, pic, px, scale) {
  runs.forEach((r, i) => {
    if (r.picture) {
      if (!pic) return;
      const img = document.createElement('img');
      img.className = 'oxe-hf-pic';
      img.alt = '';
      img.src = pic.src;
      // (`runs` one line of the section here: words before it and after it on that line.)
      const beside = (runs.slice(0, i).some(words) ? 1 : 0) + (runs.slice(i + 1).some(words) ? 1 : 0);
      setVars(img, { '--oxe-hf-pw': pic.w + beside * PIC_BESIDE * scale, '--oxe-hf-ph': pic.h });
      el.appendChild(img);
      return;
    }
    const span = document.createElement('span');
    span.textContent = r.text;
    const cls = [r.bold && 'oxe-hf-b', r.italic && 'oxe-hf-i', r.underline && 'oxe-hf-u', r.strike && 'oxe-hf-s'].filter(Boolean);
    if (cls.length) span.className = cls.join(' ');
    if (r.size) span.style.setProperty('--oxe-hf-run-size', `${px(r.size)}px`);
    if (r.font) span.style.setProperty('--oxe-hf-run-font', fontStack(r.font));
    if (r.color) span.style.setProperty('--oxe-hf-run-color', r.color);
    el.appendChild(span);
  });
}

/**
 * A page's header or footer (`which`), as Excel prints it (print-headers.xlsx): in the Normal style's font and size
 * (scaled with the sheet unless scaleWithDoc is off), the header's baseline 1.05em below its margin, the footer's 0.436em
 * above its margin; the left section from the left margin and the right one to the right margin, each 1.25px in, the
 * centre centred on the page. A section with a picture: drawPictureSection.
 */
export function drawHeaderFooter(pageEl, which, k, layout, hf) {
  const { page } = layout.sheet;
  const setup = page.headerFooter;
  if (!setup) return;
  const set = k === 0 && setup.differentFirst ? setup.first : setup.differentOddEven && k % 2 === 1 ? setup.even : setup.odd;
  const sections = parseHeaderFooter(set[which]);
  if (!sections.left.length && !sections.center.length && !sections.right.length) return;
  const scale = setup.scaleWithDoc ? layout.scale : 1;
  const basePt = hf.sizePt || 11;
  const px = (pt) => (pt * 96) / 72 * scale;
  const em = px(basePt);
  const m = page.margins;
  const y = which === 'header' ? m.header * DPI + 1.05 * em : layout.paper.h - m.footer * DPI - 0.436 * em;
  const ctx = (drawHeaderFooter.canvas ||= document.createElement('canvas')).getContext('2d');
  ctx.font = `${em}px ${fontStack(hf.family || 'Calibri')}`;
  const met = ctx.measureText('Hg');
  const asc = met.fontBoundingBoxAscent ?? em * 0.95;
  const desc = met.fontBoundingBoxDescent ?? em * 0.25;
  const inset = 1.25;
  const left = setup.alignWithMargins ? m.left * DPI + inset : inset;
  const right = layout.paper.w - (setup.alignWithMargins ? m.right * DPI + inset : inset);
  const tail = k === 0 && setup.differentFirst ? 'FIRST' : setup.differentOddEven && k % 2 === 1 ? 'EVEN' : '';
  for (const [name, parsed] of Object.entries(sections)) {
    if (!parsed.length) continue;
    const runs = fillFields(parsed, { page: k + 1, pages: layout.pages.length, sheet: hf.sheet, file: hf.file });
    const x = name === 'left' ? left - inset : name === 'right' ? right + inset : layout.paper.w / 2;
    const size = Math.max(basePt, ...runs.map((r) => r.size || 0));
    const sem = px(size);
    const pitch = printRowPx(fontRowPx(hf.family || 'Calibri', (size * 4) / 3) * 0.75, hf.family) * scale;
    // (The section's picture (&G), its VML shape named for the section: LH, CF, RHFIRST, ...)
    const found = hf.pictures?.[`${name[0].toUpperCase()}${which === 'header' ? 'H' : 'F'}${tail}`];
    if (found && runs.some((r) => r.picture)) {
      const pic = { ...found, w: found.w * PIC_W * (4 / 3) * scale, h: found.h * PIC_H * (4 / 3) * scale };
      drawPictureSection(pageEl, { name, which, runs, pic, x, em, sem, pitch, asc, desc, px, scale, family: hf.family, edge: which === 'header' ? m.header * DPI : layout.paper.h - m.footer * DPI });
      continue;
    }
    const el = div(`oxe-hf oxe-hf-${name}`);
    // (A section of several lines, as Excel prints one (its PDFs: two and three lines, 11 and 14pt, header and
    // footer): the lines a printed row of their size apart; a header's first baseline 1.014em below its margin, a
    // footer's last 0.361em above it, the lines before it above; no inset from the margin at the left or right.)
    const lines = runs.map((r) => r.text).join('').split('\n').length;
    if (lines > 1) {
      const first = which === 'header' ? m.header * DPI + LINE_IN * sem : layout.paper.h - m.footer * DPI - FOOT_LAST * sem - (lines - 1) * pitch;
      // (The section in its own size, so that its lines are the pitch apart, its ascent and descent with it.)
      const f = sem / em;
      const inBox = (pitch - (asc + desc) * f) / 2 + asc * f;
      setVars(el, { '--oxe-hf-top': first - inBox, '--oxe-hf-lh': pitch, '--oxe-hf-left': x });
    } else setVars(el, { '--oxe-hf-top': y - asc, '--oxe-hf-lh': asc + desc, '--oxe-hf-left': name === 'left' ? left : name === 'right' ? right : layout.paper.w / 2 });
    el.style.setProperty('--oxe-hf-font', fontStack(hf.family || 'Calibri'));
    el.style.setProperty('--oxe-hf-size', `${lines > 1 ? sem : em}px`);
    appendRuns(el, runs, null, px, scale);
    pageEl.appendChild(el);
  }
}

/**
 * A section with a picture, as Excel prints one (its PDFs: alone and beside text, one to three lines, header and
 * footer, at 50% and 75% with and without scaleWithDoc): its lines boxes one under the other from the header margin
 * (from the footer margin up), at the margin itself (no inset). A text line's baseline LINE_IN em below its box's top
 * and its box a printed row tall; a picture's line's baseline the picture's height below its top, the picture's bottom
 * on it, or PIC_DROP em below it when the line has characters (its line break too), the box then going on below the
 * baseline as a text line's does. Each line drawn as an element of its own: a picture's with line-height 0 and its top
 * the picture's.
 */
function drawPictureSection(pageEl, { name, which, runs, pic, x, em, sem, pitch, asc, desc, px, scale, family, edge }) {
  const f = sem / em;
  for (const p of pictureLines(linesOf(runs), { which, edge, h: pic.h, em, sem, pitch })) {
    if (!p.line.length) continue;
    const el = div(`oxe-hf oxe-hf-${name}`);
    if (p.picture) setVars(el, { '--oxe-hf-top': p.picTop, '--oxe-hf-lh': 0, '--oxe-hf-left': x, '--oxe-hf-drop': -p.drop });
    else setVars(el, { '--oxe-hf-top': p.base - asc * f, '--oxe-hf-lh': (asc + desc) * f, '--oxe-hf-left': x });
    el.style.setProperty('--oxe-hf-font', fontStack(family || 'Calibri'));
    el.style.setProperty('--oxe-hf-size', `${sem}px`);
    appendRuns(el, p.line, pic, px, scale);
    pageEl.appendChild(el);
  }
}

/** The lines of a section with a picture placed (drawPictureSection's rule): `lines` each a list of runs, from the
 *  header margin down or the footer margin up (`edge`, px), the picture `h` px tall, the text `sem` px (the Normal
 *  size `em`), lines a printed row (`pitch`) apart. [{ line, picture, chars, base (its baseline), picTop, drop }]. */
export function pictureLines(lines, { which, edge, h, em, sem, pitch }) {
  const below = pitch - LINE_IN * sem;
  const placed = lines.map((line, i) => ({ line, picture: line.some((r) => r.picture), chars: line.some(words) || i < lines.length - 1 }));
  if (which === 'header') {
    let top = edge;
    for (const p of placed) {
      p.base = top + (p.picture ? h : LINE_IN * sem);
      top = p.base + (p.picture && !p.chars ? 0 : below);
    }
  } else {
    let bottom = edge;
    for (const p of [...placed].reverse()) {
      p.base = bottom - (p.picture && !p.chars ? 0 : below);
      bottom = p.base - (p.picture ? h : LINE_IN * sem);
    }
  }
  for (const p of placed) {
    p.drop = p.chars ? PIC_DROP * em : 0;
    p.picTop = p.base - h + p.drop;
  }
  return placed;
}
