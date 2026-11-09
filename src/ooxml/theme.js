// The workbook's theme (xl/theme/theme1.xml): its heading and body fonts and its colour scheme. The one place theme
// fonts and colours are looked up.
import { kids, kid, NS } from './xml.js';

/** The theme a workbook without a theme part gets: Excel 365's current default, whose body font is Aptos Narrow
 *  (Excel's PDF of test/fixtures/basic.xlsx, a workbook with no theme part and fonts of scheme "minor", is set in
 *  Aptos Narrow). Older Excel versions used Calibri here; the heading font is Excel's, not yet measured. */
export const DEFAULT_THEME = { fonts: { major: 'Aptos Display', minor: 'Aptos Narrow' }, colors: [] };

/** The colour scheme's order, which cell colours' `theme` index counts in (dk1 and lt1 first, as Excel counts). */
const SCHEME = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** { fonts: { major, minor }, colors: ['RRGGBB', ...] } from a theme part (null: the default theme). */
export function readTheme(doc) {
  if (!doc) return DEFAULT_THEME;
  const a = NS.a;
  const el = kid(doc.documentElement, 'themeElements', a);
  const fontScheme = kid(el, 'fontScheme', a);
  const latin = (name) => kid(kid(fontScheme, name, a), 'latin', a)?.getAttribute('typeface') || null;
  const clr = kid(el, 'clrScheme', a);
  const colors = SCHEME.map((name) => {
    const c = kid(clr, name, a);
    const srgb = kid(c, 'srgbClr', a);
    const sys = kid(c, 'sysClr', a);
    return srgb ? srgb.getAttribute('val') : sys ? sys.getAttribute('lastClr') : null;
  });
  return {
    fonts: { major: latin('majorFont') || DEFAULT_THEME.fonts.major, minor: latin('minorFont') || DEFAULT_THEME.fonts.minor },
    colors: kids(clr).length ? colors : [],
  };
}
