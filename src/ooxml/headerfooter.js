// Headers and footers (Page Setup > Header/Footer): a text whose codes Excel reads (ECMA-376 Part 1, 18.3.1.46):
// &L, &C and &R start the left, centre and right sections (text before any of them is the centre's); &P the page
// number (&P+1, &P-1 offset it), &N the number of pages, &D the date, &T the time, &F the file's name, &A the sheet's,
// &Z the file's path; &B, &I, &U, &E, &S, &X, &Y bold, italic, underline, double underline, strikethrough, superscript
// and subscript (each toggles); &"Font,Style" a font (- for the same font, "Bold", "Italic" in its style), &nn a size
// in points, &Kxxxxxx a colour; && an ampersand; &G a picture (not drawn).

/**
 * A header's or footer's text as its three sections: { left, center, right }, each a list of runs { text } or
 * { field: 'page' | 'pages' | 'date' | 'time' | 'file' | 'sheet' | 'path', offset }, with { bold, italic, underline,
 * strike, size, font, color } as the codes before it set them. Each section starts in the default format.
 */
export function parseHeaderFooter(text) {
  const out = { left: [], center: [], right: [] };
  if (!text) return out;
  const fresh = () => ({ bold: false, italic: false, underline: false, strike: false, size: null, font: null, color: null });
  let section = 'center';
  let fmt = fresh();
  let buf = '';
  const flush = () => { if (buf) { out[section].push({ text: buf, ...fmt }); buf = ''; } };
  const FIELDS = { P: 'page', N: 'pages', D: 'date', T: 'time', F: 'file', A: 'sheet', Z: 'path' };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch !== '&' || i === text.length - 1) { buf += ch; continue; }
    const code = text[++i];
    if (code === '&') { buf += '&'; continue; }
    flush();
    if (code === 'L' || code === 'C' || code === 'R') { section = { L: 'left', C: 'center', R: 'right' }[code]; fmt = fresh(); continue; }
    if (FIELDS[code]) {
      // (&P+1, &P-1: the page number offset.)
      const m = code === 'P' ? /^([+-]\d+)/.exec(text.slice(i + 1)) : null;
      if (m) i += m[1].length;
      out[section].push({ field: FIELDS[code], offset: m ? Number(m[1]) : 0, ...fmt });
      continue;
    }
    if (code === 'B') fmt = { ...fmt, bold: !fmt.bold };
    else if (code === 'I') fmt = { ...fmt, italic: !fmt.italic };
    else if (code === 'U' || code === 'E') fmt = { ...fmt, underline: !fmt.underline };
    else if (code === 'S') fmt = { ...fmt, strike: !fmt.strike };
    else if (code === '"') {
      const end = text.indexOf('"', i + 1);
      const [name, style = ''] = text.slice(i + 1, end < 0 ? text.length : end).split(',');
      fmt = { ...fmt, font: name && name !== '-' ? name : fmt.font, bold: /bold/i.test(style) || (fmt.bold && !style), italic: /italic|oblique/i.test(style) || (fmt.italic && !style) };
      i = end < 0 ? text.length : end;
    } else if (/\d/.test(code)) {
      let digits = code;
      while (/\d/.test(text[i + 1] || '') && digits.length < 3) digits += text[++i];
      fmt = { ...fmt, size: Number(digits) };
    } else if (code === 'K') {
      const hex = /^[0-9A-Fa-f]{6}/.exec(text.slice(i + 1));
      if (hex) { fmt = { ...fmt, color: `#${hex[0]}` }; i += 6; }
    }
    // (X, Y, G and anything else: no change drawn.)
  }
  flush();
  return out;
}

/** A section's text for page `page` of `pages`: its runs with their fields filled in. */
export function fillFields(runs, { page, pages, sheet = '', file = '', path = '', now = new Date() }) {
  const value = (r) => {
    switch (r.field) {
      case 'page': return String(page + r.offset);
      case 'pages': return String(pages);
      case 'date': return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
      case 'time': { const h = now.getHours(); return `${h % 12 || 12}:${String(now.getMinutes()).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`; }
      case 'file': return file;
      case 'sheet': return sheet;
      case 'path': return path;
      default: return '';
    }
  };
  return runs.map((r) => ({ ...r, text: r.field ? value(r) : r.text }));
}
