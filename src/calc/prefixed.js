// The newer functions' names as the file has them: Excel 2010 and later functions carry a prefix (_xlfn.).
// Split from tokens.js, which reads and writes formulas with them.

/** Functions newer than Excel 2007, which the file names with a prefix (`_xlfn.CONCAT`) so that older versions show
 *  #NAME? instead of misreading them. Only the ones the engine knows: others typed are kept as typed. */
export const PREFIXED = new Set([
  'CONCAT', 'TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'XLOOKUP', 'XMATCH', 'IFNA', 'DAYS', 'XOR', 'STDEV.S',
  'STDEV.P', 'VAR.S', 'VAR.P', 'CEILING.MATH', 'FLOOR.MATH', 'ISOWEEKNUM', 'PERCENTILE.INC', 'PERCENTILE.EXC',
  'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'TEXTBEFORE', 'TEXTAFTER', 'NUMBERVALUE',
  'UNICODE', 'UNICHAR', 'ENCODEURL', 'FILTERXML', 'SEQUENCE', 'SORTBY', 'UNIQUE', 'LET', 'LAMBDA', 'ANCHORARRAY', 'SINGLE', 'SORT', 'FILTER',
  'TAKE', 'DROP', 'VSTACK', 'HSTACK', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND',
  'TEXTSPLIT', 'ISOMITTED', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'VALUETOTEXT', 'ARRAYTOTEXT',
  'NORM.DIST', 'NORM.S.DIST', 'NORM.INV', 'NORM.S.INV', 'EXPON.DIST', 'POISSON.DIST', 'BINOM.DIST', 'COVARIANCE.P',
  'COVARIANCE.S', 'FORECAST.LINEAR', 'SEC', 'CSC', 'COT', 'ARABIC', 'BASE', 'DECIMAL', 'ISFORMULA', 'FORMULATEXT',
  'SHEET', 'SHEETS', 'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T',
  'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT',
  'F.TEST', 'GAMMA', 'GAMMALN.PRECISE', 'GAMMA.DIST', 'GAMMA.INV', 'BETA.DIST', 'BETA.INV', 'LOGNORM.DIST',
  'LOGNORM.INV', 'WEIBULL.DIST', 'NEGBINOM.DIST', 'HYPGEOM.DIST', 'BINOM.INV', 'CONFIDENCE.NORM', 'CONFIDENCE.T',
  'Z.TEST', 'T.TEST', 'PHI', 'GAUSS', 'AGGREGATE', 'PERCENTRANK.INC', 'PERCENTRANK.EXC', 'MODE.MULT', 'MUNIT', 'BITAND',
  'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT', 'ERF.PRECISE', 'ERFC.PRECISE', 'ACOT', 'ACOTH', 'COTH', 'CSCH', 'SECH',
  'COMBINA', 'PERMUTATIONA', 'PDURATION', 'RRI', 'CEILING.PRECISE', 'FLOOR.PRECISE', 'SKEW.P', 'BINOM.DIST.RANGE',
  'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'GROUPBY', 'PIVOTBY', 'PERCENTOF', 'RANDARRAY', 'TRIMRANGE', '_TRO_TRAILING', '_TRO_LEADING', '_TRO_ALL', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN',
]);
/** Of those, the ones the file names with a second prefix (`_xlfn._xlws.SORT`), as Excel 365 writes them. */
const WORKSHEET_PREFIXED = new Set(['SORT', 'FILTER']);
/** A newer function's name as the file has it. */
export const fileName = (name) => (WORKSHEET_PREFIXED.has(name) ? `_xlfn._xlws.${name}` : PREFIXED.has(name) ? `_xlfn.${name}` : name);
