// The Format Cells box's Number tab: Excel's categories (en-US), the codes each makes from its choices, and the
// category and choices a code comes from (a code none makes is Custom). The Currency, Accounting, Percentage and
// Scientific codes are the ones the Home tab's buttons give (docs/FORMATS.md); the Date, Time and Fraction lists are
// the types Excel's box lists.
const places = (n) => (n > 0 ? `.${'0'.repeat(n)}` : '');

export const NUMBER_CATEGORIES = [
  { name: 'General', code: () => 'General' },
  { name: 'Number', decimals: 2, thousands: false, code: ({ decimals, thousands }) => `${thousands ? '#,##0' : '0'}${places(decimals)}` },
  { name: 'Currency', decimals: 2, code: ({ decimals }) => `"$"#,##0${places(decimals)}` },
  { name: 'Accounting', decimals: 2, code: ({ decimals }) => {
    const d = places(decimals);
    return `_("$"* #,##0${d}_);_("$"* \\(#,##0${d}\\);_("$"* "-"${'?'.repeat(decimals)}_);_(@_)`;
  } },
  { name: 'Date', types: ['m/d/yyyy', '[$-x-sysdate]dddd, mmmm dd, yyyy', 'm/d;@', 'm/d/yy;@', 'mm/dd/yy;@', 'd-mmm;@', 'd-mmm-yy;@', 'dd-mmm-yy;@', 'mmm-yy;@',
    'mmmm-yy;@', 'mmmm d, yyyy;@', 'm/d/yy h:mm;@', 'm/d/yyyy h:mm;@'] },
  { name: 'Time', types: ['[$-x-systime]h:mm:ss AM/PM', 'h:mm;@', '[$-409]h:mm AM/PM;@', 'h:mm:ss;@', '[$-409]h:mm:ss AM/PM;@', 'mm:ss.0;@', '[h]:mm:ss;@'] },
  { name: 'Percentage', decimals: 2, code: ({ decimals }) => `0${places(decimals)}%` },
  { name: 'Fraction', types: ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'] },
  { name: 'Scientific', decimals: 2, code: ({ decimals }) => `0${places(decimals)}E+00` },
  { name: 'Text', code: () => '@' },
  { name: 'Custom', custom: true },
];

/** The code a category makes from the box's choices: { decimals, thousands, type }. */
export function codeFor(category, choices) {
  if (category.types) return choices.type || category.types[0];
  return category.code(choices);
}

/** The category a code comes from, with the choices that make it ({ name, decimals, thousands, ...category }). */
export function categoryOf(code) {
  for (const c of NUMBER_CATEGORIES) {
    if (c.custom) continue;
    if (c.types) { if (c.types.includes(code)) return c; continue; }
    if (c.decimals === undefined) { if (c.code() === code) return c; continue; }
    for (let decimals = 0; decimals <= 30; decimals++) {
      for (const thousands of c.thousands === undefined ? [false] : [false, true]) {
        if (c.code({ decimals, thousands }) === code) return { ...c, decimals, thousands: c.thousands === undefined ? undefined : thousands };
      }
    }
  }
  return NUMBER_CATEGORIES.at(-1);
}
