// drawings.xlsx: pictures, shapes, icons and a group on a sheet, written as Excel writes its drawing parts, for drawing
// them where Excel does (on screen and in print: Excel's PDF of it, docs/LAYOUT.md "Pictures, shapes and icons") and
// for test/drawobjects.test.js. All made in code: the pictures are PNGs of four coloured quarters (so a crop, a turn
// or a flip shows), the icon an SVG circle with a PNG beside it, as Excel's icons have them.
//
// Sheet "Drawings" (column G hidden, row 6 hidden, column C 115px wide):
//   B2:E8   "Quarters", a picture, editAs oneCell, its extent the size it has with row 6 shown
//   F2:J8   "Cropped", a picture over hidden column G, cropped 25% left, 10% top, 20% bottom
//   K3      "Turned", a picture (one-cell anchor) turned 30 degrees and flipped across
//   N2      "Icon 1", an SVG icon and its PNG (one-cell anchor)
//   B11:C15 "Rectangle 5", orange with a 2pt dark line
//   E11:F15 "Rounded 6", a rounded rectangle in the default shape style, white bold 14pt "Rounded" in the middle
//   H11:I15 "Oval 7", an ellipse in the default style with "Oval" in it
//   K11:L15 "Triangle 8", accent 2 at 75% luminance, no line
//   B18:D22 "Arrow 9", a right arrow turned 45 degrees, with a dashed line
//   F18:I22 "TextBox 10", a text box: two paragraphs, left and right, wrapped, a thin black line
//   K18:M22 "Connector 11", a straight connector with an arrowhead at its end, 1.5pt
//   B25:D29 "Elbow 12", a bent connector flipped upright, arrowhead at its start
//   F25:G29 "Cloud 13", a preset not drawn yet (its box)
//   I25:L30 "Group 14", a rectangle and an oval in a group whose children's space is half its own
//   at 900, 220px "Star 15", a five-pointed star on an absolute anchor
//   B32:C34 "Hidden 16", a shape hidden in the file
import zlib from 'node:zlib';
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, THEME, esc } from './fixture-common.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const EMU = 9525;

/** A PNG of w by h pixels, `px(x, y)` giving [r, g, b] for each. */
export function png(w, h, px) {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => {
    const out = Buffer.alloc(12 + data.length);
    out.writeUInt32BE(data.length, 0);
    out.write(type, 4, 'ascii');
    data.copy(out, 8);
    out.writeUInt32BE(crc(out.subarray(4, 8 + data.length)), 8 + data.length);
    return out;
  };
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 3 + 1)] = 0;
    for (let x = 0; x < w; x++) raw.set(px(x, y), y * (w * 3 + 1) + 1 + x * 3);
  }
  const head = Buffer.alloc(13);
  head.writeUInt32BE(w, 0); head.writeUInt32BE(h, 4); head[8] = 8; head[9] = 2; head[10] = 0; head[11] = 0; head[12] = 0;
  return new Uint8Array(Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', head), chunk('IDAT', zlib.deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]));
}

// (Four quarters: red, green, blue, yellow, as the picture is drawn upright and unflipped.)
const QUARTERS = png(120, 80, (x, y) => (x < 60 ? (y < 40 ? [220, 40, 40] : [40, 90, 220]) : y < 40 ? [40, 170, 60] : [240, 200, 20]));
const ICON_PNG = png(48, 48, (x, y) => ((x - 23.5) ** 2 + (y - 23.5) ** 2 < 22 * 22 ? [112, 48, 160] : [255, 255, 255]));
const ICON_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96" width="96" height="96"><circle cx="48" cy="48" r="44" fill="#7030A0"/><path d="M28 50 L42 64 L70 34" stroke="#FFFFFF" stroke-width="8" fill="none"/></svg>';

const COLS = { 2: 120 }; // column C (index 2) 120px, G (6) hidden
const HIDDEN_COL = 6;
const HIDDEN_ROW = 5;
const colPx = (c) => (c === HIDDEN_COL ? 0 : COLS[c] ?? 64);
const rowPx = (r) => (r === HIDDEN_ROW ? 0 : 20);
const xAt = (c) => { let x = 0; for (let i = 0; i < c; i++) x += colPx(i); return x; };
const yAt = (r) => { let y = 0; for (let i = 0; i < r; i++) y += rowPx(i); return y; };

const pt = (tag, c, r, dx = 0, dy = 0) => `<xdr:${tag}><xdr:col>${c}</xdr:col><xdr:colOff>${dx * EMU}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${dy * EMU}</xdr:rowOff></xdr:${tag}>`;
const xfrm = (x, y, w, h, extra = '') => `<a:xfrm${extra}><a:off x="${Math.round(x * EMU)}" y="${Math.round(y * EMU)}"/><a:ext cx="${Math.round(w * EMU)}" cy="${Math.round(h * EMU)}"/></a:xfrm>`;
const two = (c0, r0, c1, r1, body, editAs = '') => `<xdr:twoCellAnchor${editAs ? ` editAs="${editAs}"` : ''}>${pt('from', c0, r0)}${pt('to', c1, r1)}${body}<xdr:clientData/></xdr:twoCellAnchor>`;
const box = (c0, r0, c1, r1) => [xAt(c0), yAt(r0), xAt(c1) - xAt(c0), yAt(r1) - yAt(r0)];
// (Excel's default shape style: accent 1, its line darker.)
const STYLE = '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>';
const body = (paras, bodyPr = 'vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"') => `<xdr:txBody><a:bodyPr ${bodyPr}/><a:lstStyle/>${paras}</xdr:txBody>`;
const para = (text, { algn = null, sz = 1100, b = false, color = null } = {}) => `<a:p>${algn ? `<a:pPr algn="${algn}"/>` : ''}<a:r><a:rPr lang="en-US" sz="${sz}"${b ? ' b="1"' : ''}>${color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : ''}</a:rPr><a:t>${esc(text)}</a:t></a:r></a:p>`;
const sp = (id, name, prst, geomBox, { fill = '', line = '', style = '', text = '', extra = '', hidden = false, av = '' } = {}) =>
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"${hidden ? ' hidden="1"' : ''}/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${xfrm(...geomBox, extra)}<a:prstGeom prst="${prst}"><a:avLst>${av}</a:avLst></a:prstGeom>${fill}${line}</xdr:spPr>${style}${text}</xdr:sp>`;
const pic = (id, name, rid, geomBox, { crop = '', extra = '', svg = null, descr = '' } = {}) =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name}"${descr ? ` descr="${esc(descr)}"` : ''}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
  `<xdr:blipFill><a:blip xmlns:r="${R}" r:embed="${rid}">${svg ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${svg}"/></a:ext></a:extLst>` : ''}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  `<xdr:spPr>${xfrm(...geomBox, extra)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
const solid = (inner) => `<a:solidFill>${inner}</a:solidFill>`;

export function drawings() {
  const anchors = [];
  // Pictures.
  // (Quarters: from B2 to D8 with row 6 hidden; its extent is 2 columns and 6 rows of 20px as if row 6 were shown.)
  const q = [xAt(1), yAt(1), xAt(4) - xAt(1), 6 * 20 + 20];
  anchors.push(two(1, 1, 4, 8, pic(2, 'Quarters', 'rId1', q, { descr: 'Four coloured quarters' }), 'oneCell'));
  anchors.push(two(5, 1, 9, 8, pic(3, 'Cropped', 'rId1', box(5, 1, 9, 8), { crop: '<a:srcRect l="25000" t="10000" b="20000"/>' })));
  anchors.push(`<xdr:oneCellAnchor>${pt('from', 10, 2, 8, 4)}<xdr:ext cx="${120 * EMU}" cy="${80 * EMU}"/>${pic(4, 'Turned', 'rId1', [xAt(10) + 8, yAt(2) + 4, 120, 80], { extra: ' rot="1800000" flipH="1"' })}<xdr:clientData/></xdr:oneCellAnchor>`);
  anchors.push(`<xdr:oneCellAnchor>${pt('from', 13, 1)}<xdr:ext cx="${48 * EMU}" cy="${48 * EMU}"/>${pic(5, 'Icon 1', 'rId2', [xAt(13), yAt(1), 48, 48], { svg: 'rId3', descr: 'A tick in a circle' })}<xdr:clientData/></xdr:oneCellAnchor>`);
  // Shapes.
  anchors.push(two(1, 10, 3, 15, sp(6, 'Rectangle 5', 'rect', box(1, 10, 3, 15), { fill: solid('<a:srgbClr val="FFC000"/>'), line: `<a:ln w="${2 * 12700}">${solid('<a:srgbClr val="7F6000"/>')}</a:ln>` })));
  anchors.push(two(4, 10, 6, 15, sp(7, 'Rounded 6', 'roundRect', box(4, 10, 6, 15), { style: STYLE, text: body(para('Rounded', { algn: 'ctr', sz: 1400, b: true }), 'rtlCol="0" anchor="ctr"') })));
  anchors.push(two(7, 10, 9, 15, sp(8, 'Oval 7', 'ellipse', box(7, 10, 9, 15), { style: STYLE, text: body(para('Oval', { algn: 'ctr' }), 'rtlCol="0" anchor="ctr"') })));
  anchors.push(two(10, 10, 12, 15, sp(9, 'Triangle 8', 'triangle', box(10, 10, 12, 15), { fill: solid('<a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr>'), line: '<a:ln><a:noFill/></a:ln>' })));
  anchors.push(two(1, 17, 4, 22, sp(10, 'Arrow 9', 'rightArrow', box(1, 17, 4, 22), { extra: ' rot="2700000"', fill: solid('<a:schemeClr val="accent6"/>'), line: `<a:ln w="19050">${solid('<a:srgbClr val="375623"/>')}<a:prstDash val="dash"/></a:ln>` })));
  anchors.push(two(5, 17, 9, 22, `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="11" name="TextBox 10"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr>${xfrm(...box(5, 17, 9, 22))}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525" cmpd="sng">${solid('<a:srgbClr val="000000"/>')}</a:ln></xdr:spPr>` +
    `<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>` +
    `${body(para('A text box of two paragraphs, this one long enough to wrap.') + para('Right', { algn: 'r', color: 'C00000' }), 'vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="t"')}</xdr:sp>`));
  const cxn = (id, name, prst, b, ln, extra = '') => `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>${xfrm(...b, extra)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${ln}</xdr:spPr></xdr:cxnSp>`;
  anchors.push(two(10, 17, 13, 22, cxn(12, 'Connector 11', 'straightConnector1', box(10, 17, 13, 22), `<a:ln w="19050">${solid('<a:schemeClr val="accent1"/>')}<a:tailEnd type="triangle"/></a:ln>`)));
  anchors.push(two(1, 24, 4, 29, cxn(13, 'Elbow 12', 'bentConnector3', box(1, 24, 4, 29), `<a:ln w="12700">${solid('<a:srgbClr val="C00000"/>')}<a:headEnd type="triangle"/></a:ln>`, ' flipV="1"')));
  anchors.push(two(5, 24, 8, 29, sp(14, 'Cloud 13', 'cloud', box(5, 24, 8, 29), { fill: solid('<a:srgbClr val="DDEBF7"/>'), line: `<a:ln w="9525">${solid('<a:srgbClr val="5B9BD5"/>')}</a:ln>` })));
  // (A group: its children's space half its own, so each child is drawn twice its file size.)
  const [gx, gy, gw, gh] = box(9, 24, 13, 30);
  const child = (id, name, prst, x, y, w, h, color) => `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${xfrm(x, y, w, h)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${solid(`<a:srgbClr val="${color}"/>`)}<a:ln w="9525">${solid('<a:srgbClr val="000000"/>')}</a:ln></xdr:spPr></xdr:sp>`;
  const cx = 1000;
  const cy = 1000;
  anchors.push(two(9, 24, 13, 30, `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="15" name="Group 14"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="${gx * EMU}" y="${gy * EMU}"/><a:ext cx="${gw * EMU}" cy="${gh * EMU}"/><a:chOff x="${cx * EMU}" y="${cy * EMU}"/><a:chExt cx="${(gw / 2) * EMU}" cy="${(gh / 2) * EMU}"/></a:xfrm></xdr:grpSpPr>` +
    child(16, 'Group Rectangle', 'rect', cx, cy, gw / 4, gh / 4, '70AD47') + child(17, 'Group Oval', 'ellipse', cx + gw / 4, cy + gh / 4, gw / 4, gh / 4, 'ED7D31') + '</xdr:grpSp>'));
  anchors.push(`<xdr:absoluteAnchor><xdr:pos x="${900 * EMU}" y="${220 * EMU}"/><xdr:ext cx="${80 * EMU}" cy="${80 * EMU}"/>${sp(18, 'Star 15', 'star5', [900, 220, 80, 80], { fill: solid('<a:srgbClr val="FFD966"/>'), line: `<a:ln w="12700">${solid('<a:srgbClr val="BF8F00"/>')}</a:ln>` })}<xdr:clientData/></xdr:absoluteAnchor>`);
  anchors.push(two(1, 31, 3, 34, sp(19, 'Hidden 16', 'rect', box(1, 31, 3, 34), { hidden: true, fill: solid('<a:srgbClr val="FF0000"/>') })));

  const drawing = `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">${anchors.join('')}</xdr:wsDr>`;
  const t = (ref, text) => `<c r="${ref}" t="inlineStr"><is><t>${esc(text)}</t></is></c>`;
  const rows = [`<row r="1">${t('A1', 'Pictures, shapes and icons')}</row>`];
  for (let r = 2; r <= 36; r++) rows.push(`<row r="${r}"${r - 1 === HIDDEN_ROW ? ' hidden="1"' : ''}>${t(`A${r}`, `Row ${r}`)}</row>`);
  const sheet = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
    `<cols><col min="3" max="3" width="${(120 - 5) / 7}" customWidth="1"/><col min="7" max="7" width="9.140625" hidden="1" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></worksheet>';
  write('drawings.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
      ['xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'], ...PROPS_CT])
      .replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="xml"'),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Drawings" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': sheet,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('drawing'), '../drawings/drawing1.xml']]),
    'xl/drawings/drawing1.xml': drawing,
    'xl/drawings/_rels/drawing1.xml.rels': rels([['rId1', RT('image'), '../media/image1.png'], ['rId2', RT('image'), '../media/image2.png'], ['rId3', RT('image'), '../media/image3.svg']]),
    'xl/media/image1.png': QUARTERS,
    'xl/media/image2.png': ICON_PNG,
    'xl/media/image3.svg': ICON_SVG,
    'xl/styles.xml': STYLES,
    'xl/theme/theme1.xml': THEME,
  });
}
