// Turning a cell's value into the text a sheet shows. Only Excel's General format so far; format codes (`0.00`,
// dates, `[Red]`, sections) come with styles.xml's numFmts.

/** A number in the General format: as many digits as fit in 11 characters (a default-width column), trailing zeros
 *  dropped; scientific with up to 5 decimals ("1.23457E+11") from 1E+11 up and below 0.0001 ("1E-05"). (Excel's
 *  behaviour as commonly described; the exact switch-over points are still to be measured on Excel's output.) */
export function formatGeneral(n) {
  if (!Number.isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const a = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  const intDigits = Math.floor(Math.log10(a)) + 1;
  if (intDigits <= 11 && a >= 1e-4) {
    // ("0." takes two characters below 1; an integer part of d digits leaves 10 - d for the point and decimals.)
    const decimals = intDigits <= 0 ? 9 : Math.max(0, 10 - intDigits);
    let s = a.toFixed(decimals);
    if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
    return sign + s;
  }
  const [m, e] = a.toExponential(5).split('e');
  const ei = Number(e);
  return `${sign}${m.replace(/\.?0+$/, '')}E${ei < 0 ? '-' : '+'}${String(Math.abs(ei)).padStart(2, '0')}`;
}

/** The text a cell shows, from what the reader stored ({ t, v }). */
export function cellText(cell) {
  if (!cell || cell.v == null) return '';
  switch (cell.t) {
    case 'b': return cell.v ? 'TRUE' : 'FALSE';
    case 'n': return formatGeneral(cell.v);
    default: return String(cell.v);
  }
}
