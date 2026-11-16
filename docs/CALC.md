# Calculation

How formulas are worked out (`src/calc/`), and what Excel was seen to do. As with layout, Excel's own results are
the only measure. `test/fixtures/formulas.xlsx` holds 1926 formula cells with no saved values, and a table (Fruit) for table references. Excel opened it,
worked them out (`tools/excel-compare/excel-values.ps1`, COM, read-only, `CalculateFull`) and its results are in
`test/formulas-excel.json`. `npm test` checks every one: **1926 of 1926 as Excel** (to 15 digits; RATE, IRR, YIELD and ODDFYIELD, searches (Excel promises RATE and IRR to 0.0000001), to 1e-9; XIRR, which it promises to 0.000001%, to 1e-8;
the t, chi-squared, F, gamma, beta, Weibull and negative binomial distributions, their inverses and tests, and VDB, to
1e-13 of the value: series, continued fractions and searches whose last digits are each one's own, Excel's within
2e-14 of these; the IM functions' complex results, each part to 2e-15, since a part falling on a rounding boundary in
the 15th digit takes its last digit from the JavaScript engine's Math.pow: IMPOWER("1.5-2i",2.5) is Excel's
-6.7175144212722-7.24784450716211i in Node 20 and ...212i in Node 24 and Chrome 156).

## How it works

- **Reading formulas.**
  - `tokens.js` reads a formula into tokens; `fileform.js` turns it between the typed form and the file's (`prefixed.js`:
    the newer functions the file names with `_xlfn.`).
  - `parse.js` builds a tree by Excel's precedence. Negation binds tightest, so `=-2^2` is 4, and operators of one
    level work left to right, so `=2^3^2` is 64.
- **Working them out.**
  - `evaluate.js` handles the operators, references and names.
  - About 130 functions in three files, by Excel's categories: `functions.js` (maths, statistics, logic, information
    and the IF/IFS family, plus the helpers the others share), `functions-text.js` (text, and dates and times) and
    `functions-lookup.js` (lookup: VLOOKUP, INDEX, MATCH, XLOOKUP, OFFSET, INDIRECT; and SUMIF, COUNTIFS and the other
    conditional sums and counts).
  - `functions-more.js` holds 54 more: SUBTOTAL (rows a filter hides left out by all, rows hidden by hand by 101 to 111, other subtotals always), RANK,
    PERCENTILE, QUARTILE and MODE; NETWORKDAYS, WORKDAY, DATEDIF, WEEKNUM, ISOWEEKNUM and YEARFRAC; PMT, PV, FV, NPER,
    RATE, NPV and IRR; LOOKUP, XMATCH and ADDRESS; TEXTBEFORE, TEXTAFTER, NUMBERVALUE, FIXED and DOLLAR; trigonometry,
    GCD, LCM, CEILING.MATH and FLOOR.MATH; TYPE and ERROR.TYPE. Excel raises 1 + rate to a whole number of periods by
    repeated squaring in PV and FV but through logarithms in PMT, and each is computed so, to Excel's last digit. RATE
    and IRR are found by Newton's method (IRR on the discount factor, so IRR({-100,10}) is -0.9).
  - `functions-arrays.js`: Excel 365's TAKE, DROP, VSTACK, HSTACK, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW, WRAPROWS,
    WRAPCOLS, EXPAND, TEXTSPLIT, VALUETOTEXT and ARRAYTOTEXT, and the LAMBDA helpers MAP, REDUCE, SCAN, BYROW, BYCOL,
    MAKEARRAY and ISOMITTED (docs/DYNAMIC.md for spilling).
  - `functions-finance.js`: IPMT, PPMT, CUMIPMT, CUMPRINC, SLN, DDB, DB, EFFECT, NOMINAL, XNPV, XIRR; TIMEVALUE (read as an
    entry is read), WORKDAY.INTL and NETWORKDAYS.INTL (Excel's weekend numbers and seven-day strings). As with PV and FV,
    each computed the way that gives Excel's last digit: DDB's book value and EFFECT by repeated squaring, IPMT's
    balance through logarithms.
  - `functions-stats.js`: CORREL, PEARSON, SLOPE, INTERCEPT, RSQ, FORECAST and FORECAST.LINEAR, COVARIANCE.P and .S,
    GEOMEAN, HARMEAN, AVEDEV, DEVSQ, SKEW, KURT, STANDARDIZE, PERMUT, FISHER; NORM.DIST, NORM.S.DIST, NORM.INV, NORM.S.INV
    (W. J. Cody's erfc, the inverse polished by Halley's method: to 15 digits), EXPON.DIST, POISSON.DIST, BINOM.DIST;
    SINH, COSH, TANH, SEC, CSC, COT; ROMAN, ARABIC, BASE, DECIMAL, DEC2BIN, BIN2DEC, DEC2HEX, HEX2DEC, MULTINOMIAL;
    ISFORMULA, FORMULATEXT, SHEET and SHEETS.
  - `functions-dist.js`: the distributions the others lack, on Lanczos' log-gamma and the regularised incomplete gamma
    (series and continued fraction) and beta (continued fraction) functions, inverses by bisection then Newton: T.DIST,
    .2T, .RT, T.INV, .2T; CHISQ.DIST, .RT, CHISQ.INV, .RT; F.DIST, .RT, F.INV, .RT; GAMMA, GAMMALN(.PRECISE), GAMMA.DIST,
    GAMMA.INV; BETA.DIST, BETA.INV; LOGNORM.DIST, LOGNORM.INV; WEIBULL.DIST; NEGBINOM.DIST; HYPGEOM.DIST (the ways
    multiplied out, so exact); BINOM.INV; CONFIDENCE.NORM and .T; Z.TEST, T.TEST (paired, equal and unequal
    variances), CHISQ.TEST, F.TEST; PHI, GAUSS, FISHERINV; and the names of Excel 2007 and before for each (TDIST,
    CHIDIST, FDIST, NORMDIST, CRITBINOM...). Degrees of freedom are cut to whole numbers, T.DIST's too.
  - `functions-agg.js`: AGGREGATE (its options leave out nested subtotals, hidden rows and errors; functions 1 to 13
    take references only, an array #VALUE!), FREQUENCY, TRIMMEAN, PERCENTRANK(.INC, .EXC), MODE.MULT; VARA, VARPA,
    STDEVA, STDEVPA (text 0 and TRUE 1 in ranges, only the numbers of an array), STEYX, PROB; MMULT, MDETERM, MINVERSE,
    MUNIT; SUMX2MY2, SUMX2PY2, SUMXMY2; LINEST, LOGEST, TREND and GROWTH. LINEST takes the xs and ys from their means
    first, as Excel does, so an exact line has residuals of exactly 0 and its F is #NUM!.
  - `functions-eng.js`: BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT; OCT2DEC, DEC2OCT, BIN2HEX, HEX2BIN, BIN2OCT,
    OCT2BIN, HEX2OCT, OCT2HEX; DELTA, GESTEP, ERF, ERFC (.PRECISE); ACOSH, ASINH, ATANH, ACOT, ACOTH, COTH, CSCH, SECH;
    COMBINA, PERMUTATIONA, FACTDOUBLE, SERIESSUM, SQRTPI; DAYS360 (Excel's US method leaves an end on the last of
    February as it is); HYPERLINK, AREAS; FVSCHEDULE, PDURATION, RRI, ISPMT, MIRR (from the two NPVs, as Excel defines
    it), SYD, VDB, DOLLARDE, DOLLARFR, TBILLPRICE, TBILLYIELD, TBILLEQ.
  - `functions-db.js`: the database functions (DSUM, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV,
    DSTDEVP, DVAR, DVARP), their criteria read as Advanced Filter reads them (`criteria.js`, shared with
    `ooxml/advfilter.js`); CEILING.PRECISE, FLOOR.PRECISE, ISO.CEILING, SKEW.P, BINOM.DIST.RANGE, COVAR.
  - `functions-bonds.js`: securities. COUPPCD, COUPNCD, COUPNUM, COUPDAYBS, COUPDAYS, COUPDAYSNC; PRICE, YIELD,
    DURATION, MDURATION; ACCRINT, ACCRINTM; DISC, INTRATE, RECEIVED, PRICEDISC, YIELDDISC, PRICEMAT, YIELDMAT. Coupon
    dates step back from maturity, kept at the month's end when maturity is. What Excel was seen to do (each measured
    with Evaluate in a hidden Excel, then checked by the probe on cases it wasn't worked out from):
    - COUPDAYS with basis 1 runs from the coupon date before settlement to the same day 12/frequency months on
      (February 29 to August 29: 182 days, not the 184 to August 31). PRICE, YIELD and DURATION take the actual period
      (184) instead. COUPDAYSNC counts 30/360 days with basis 4 and takes COUPDAYS - COUPDAYBS with basis 0.
    - PRICE, YIELD and DURATION take the days to the next coupon as COUPDAYS - COUPDAYBS, but for basis 1.
    - ACCRINT (measured on 600 cases over the five bases and both calc_methods, then on 400 it wasn't worked out from,
      387 the same): the coupon dates step from first_interest back past issue, each from the one before (August 30,
      February 28, August 28), at the month's end when first_interest is at one, and on again from the earliest, so the
      period issue falls in can end before first_interest. A period's length L is 360 or 365 over the frequency, or
      with basis 1 the days of the period ending at first_interest (issue's own period its own days). Issue in the
      period ending at first_interest: the days from issue over L. A longer first period accrues at L's rate all
      through it: issue's part over its length, 1 for each whole period after it but the last, and the days left, from
      issue to settlement less those periods', over L, so the answer is linear in settlement with an offset (issue
      October 18, 2005, first_interest February 18, 2007, half-yearly, basis 1: 3 days more than counted, 184 less 181).
      Past first_interest each whole period counts 1 and a part its days over L; a period ending on settlement is a part
      (August 31 to February 28: 178 days of 180 in 30/360). With calc_method FALSE, issue's part and the days from
      the start of the period ending at first_interest to settlement over L: negative when settlement is before that
      start (ACCRINT(DATE(2006,11,1),DATE(2008,1,15),DATE(2007,6,1),0.06,100,4,1,FALSE) is -0.99 in Excel).
    - ACCRINT at the ends of Excel's dates (4251 more cases over COM, `tools/excel-compare/accrint-cases.mjs`, Excel's
      answers in test/accrint-excel.json; 1911 of them #NUM!, every one as Excel's, 4213 as Excel's in all): **#NUM!
      where the coupon walk steps to a date before 1900-01-01 or past 9999-12-31**, whichever dates the case walks
      through. Back from first_interest to issue, on from issue's period to first_interest and on past it to the
      period holding settlement, and with basis 1 the period ending at first_interest. With calc_method FALSE and issue
      on or after first_interest Excel steps back from first_interest once too, whatever the basis. So
      ACCRINT(1,731,731,0.1,1000,4,1,FALSE) is #NUM! (back from 1902-01-01 past 1900), ACCRINT(2958000,2958440,
      2958450,0.1,1000,4) too (the period after 9999-12-06 ends in 10000), but ACCRINT(2958065,2958281,2958464,0.1,
      1000,2,0,TRUE) is worked out (the period after 9999-06-30 ends on 9999-12-31). Before this rule 2091 of the 3951
      near the ends were as Excel's.
    - The coupon walk steps over a February 1900 of 28 days, as EDATE and EOMONTH do (EDATE(31,1) is 59, EOMONTH(40,0)
      59), though serial 60 still reads as February 29, 1900; DATE treats 1900 as a leap year (DATE(1900,2,30) is 61).
      ACCRINT's 30/360 days take February 28, 1900 as the month's last (taking the 29th made 6 more cases wrong).
    - Issue exactly on a coupon date: with calc_method FALSE (and two or more periods to first_interest) issue's
      period counts nothing, the answer the days from the period ending at first_interest's start to settlement over
      L; with TRUE it counts 1, not its days over L, and with settlement before first_interest each period ended
      before settlement counts 1 and the rest its days over L (ACCRINT(DATE(2010,12,20),DATE(2011,12,20),
      DATE(2011,10,18),0.1,1000,2,2,TRUE) is 83.33: 1 and 120/180). 287 of 300 such cases as Excel's (121 before).
    - DISC and its kin with basis 1 take YEARFRAC's year (366 over a February 29); PRICEMAT and YIELDMAT the year up
      to settlement's. DISC is (1 - pr/redemption) x (year / days), to Excel's last digit.
  - `functions-complex.js`: the complex numbers, text such as "3+4i" or "3+4j". COMPLEX, IMREAL, IMAGINARY, IMABS,
    IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER, IMSQRT, IMEXP, IMLN, IMLOG10, IMLOG2, IMSIN,
    IMCOS, IMTAN, IMCOT, IMSEC, IMCSC, IMSINH, IMCOSH, IMSECH, IMCSCH. What Excel was seen to do (about 350 cases with
    Evaluate, then the probe's 105):
    - Reading: the real part, then the imaginary part from the first + or - after the first character that doesn't
      follow an E. Each part is read as a number typed into a cell is (" 4", "$2", "(2)" is -2, "2%", "1,000", "1e3").
      A space may come before that sign but not stand alone after it ("3 +i" is fine, "3+ i" is #NUM!). "3+-4i" is
      3-4i, "++i" and "3+4I" are #NUM!. A number is its real part, an empty cell 0, TRUE #VALUE!.
    - Writing: each part to 15 significant digits, written out in full up to 21 characters (one more than a number
      turned into text gets): "0.0000000000000000001+i", but "1.5E-19+i". A coefficient of 1 is left out ("i", "-i"),
      and a zero part too ("3", "0"). "i" and "j" mixed: IMSUM and IMPRODUCT #VALUE!, IMSUB and IMDIV #NUM!.
    - Past the largest number: IMPRODUCT #VALUE!, the others #NUM!. IMABS is #NUM! once |re| + |im| is
      ("1E308+1E308i"). cosh and sinh are #NUM! once e^x is (IMCOSH("710")), though their value would fit.
    - Powers and roots by the polar form, so IMPOWER("i",2) is -1+1.22514845490862E-16i. The tangent and the
      reciprocals by the double angle (tan(x+iy) = (sin 2x + i sinh 2y) / (cos 2x + cosh 2y)), the real functions'
      own for a real number; ±i and 0 where the double angle is past the largest number (IMTAN("1+400i") is i).
  - `functions-convert.js`: CONVERT, with Excel's 152 unit names in 13 kinds. Each factor is Excel's own, read to
    the last bit from CONVERT(1,unit,base) in a hidden Excel. A conversion is number × factor ÷ factor, which matches
    Excel to the last bit in most cases and to 2 bits in the rest (724 cases). The units that take a prefix were
    measured too: 40 of them, among them mph ("kmph"), ar ("kar" is 1000 ar) and the litre ("cl").
    - A prefix is squared or cubed for a length squared or cubed ("km2" is 1E+06 m2), and only one is allowed ("kmm"
      is #N/A). "da" and "e" are both deka. The binary prefixes (ki, Mi, ... Yi) are for bit and byte only.
    - Names are exact: "M", "FT" and " m" are #N/A, as are units of different kinds. TRUE is #VALUE!, and "2" is 2.
    - Temperatures go through kelvin, with "K" and "kel" taking a prefix ("mK"), and C, F, Rank and Reau not.
  - `functions-oddbonds.js`: ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD, by Microsoft's formulas with Excel's own
    counting, measured on 360 cases with cell values (prices to 1e-15, ODDFYIELD's search to 1e-9), then the
    probe's 54:
    - The quasi-coupon dates of an odd period step from the first coupon back, or from the last interest date on, each
      from the one before on the same day, not kept at the month's end: June 30, December 30, June 30; December 31,
      September 30, June 30, March 30. The days to the next coupon (bases 2 and 3) and from the last (the others)
      take COUPNCD's and COUPPCD's dates, which are.
    - A short first period's N counts the first coupon. A long one's whole quasi-periods from settlement to the first
      coupon are counted on from settlement, moved to its month's end when the first coupon is at one (which counts
      one): settlement on February 10, 2020 and a first coupon on June 30 make 1 half-year, not 0.
    - An odd last period's quasi-periods have their actual days for length with bases 1, 2 and 3, 360 over the
      frequency with 0 and 4.
    - Issue on settlement is #NUM!, as are dates out of order, a negative rate, a frequency other than 1, 2, 4.
  - `functions-bessel.js`: BESSELJ, BESSELY, BESSELI and BESSELK. Excel's are not the exact functions but Numerical
    Recipes' approximations, good to about 1e-8: =BESSELJ(0,0) is 1.00000000283141. The rational and polynomial
    forms for orders 0 and 1, then recurrences: up for Y, K, and J past its order; down (Miller's, 40 for the start)
    for I and J. The coefficients past 8 are two printings: Y0 shows 0.934945152e-7 where J0 has 0.934935152e-7,
    and J0's p -0.1098628267e-2 where Y0's is -0.1098628627e-2; J1's p has -0.3516396496e-4 twice, a term more
    than Y1's (so J1(8) is 0.2346283865070744, 8e-6 from the function). Found from 221 points from 8 to 30, to the
    last bit. Measured on about 970 cases with cell values, then the probe's 60; all as Excel's to the last bit or
    within 1e-15. n is made whole, below 0 #NUM!; Y and K need x above 0; J and Y are #NUM! from 2^27 on, as SIN is.
  - `iterate.js`: circular references (a formula that reads itself, or a circle of them), measured with a hidden
    Excel on 30 cases (each with iteration off, on with 100 and on with 10 iterations; entered, F9, Ctrl+Alt+F9) and
    an 11-step sequence in one workbook, all the same here, to the last bit:
    - Iteration off: a formula in a circle isn't worked out; it keeps its value, and the one that closes the circle
      shows 0 (=A1+1 in A1 is 0). The editor says which formulas are circular.
    - Iteration on (calcPr iterate="1"; iterateCount and iterateDelta written only when not 100 and 0.001): every
      recalculation goes round, whatever changed (=A1+1 is 100, then 200 after any edit). Each pass works out the
      formulas in circles and those reading them, row by row of each sheet in order (B1 before A2), reading the last
      pass's values; one with no value yet reads as 0. The passes stop when no number changed by Maximum Change or
      more, the whole workbook's (one circle still changing keeps every circle going), or after Maximum Iterations:
      =1+A1/2 stops at 1.9990234375 after 11 passes.
  - `threed.js`: 3D references (=SUM(Jan:Apr!A1), 'Feb:My Mar'!A1:B2) and the intersection operator, a space
    between references (=A1:C3 B2:D4 is B2:C3, spilled; =SUM(B:B 2:2) is B2). Measured on 86 formulas over four sheets
    in a hidden Excel, all the same here, then the probe's 17:
    - A 3D reference is the area on each sheet from the first to the last in the workbook's order. SUM, COUNT, COUNTA,
      AVERAGE(A), MAX(A), MIN(A), PRODUCT, STDEV, VAR and their kinds, and in Excel 365 also MEDIAN, LARGE, SMALL,
      PERCENTILE, QUARTILE, SUMSQ, AVEDEV, GEOMEAN, DEVSQ, SKEW, KURT, CONCAT, TEXTJOIN, AND, OR, VSTACK and HSTACK
      (sheet after sheet) and TOCOL take one; IF, CHOOSE and IFERROR pass it on; ISREF says FALSE.
    - Functions that want a reference of one sheet (COUNTIF(S), SUMIF(S), ROWS, INDEX, N, AGGREGATE, SUBTOTAL,
      COUNTBLANK) give #VALUE!; any other, and the reference as a value (=Jan:Apr!A1, +1, SUMPRODUCT, MODE), #REF!.
    - The first name reading as a cell makes it a range, as Excel reads it: S1:S4!A1 is S1 to S4!A1.
    - Intersection binds tighter than negation; no overlap is #NULL!; references on two sheets, or a 3D one, #VALUE!.
    - The file has them as typed: `SUM(Jan:Apr!A1)`, `'Feb:My Mar'!A1:B2`, `A1:C3 B2:D4`.
    - Typed, a 3D reference is put in the workbook's order (`Apr:Jan!A1` is `Jan:Apr!A1`) and one of one sheet made
      plain (`Jan:Jan!A1` is `Jan!A1`). As sheets go and move (16 cases measured, all as Excel's): an end sheet that
      leaves the span's other sheets, deleted or moved out, gives way to the next of them inward (Jan:Apr with Jan
      deleted, or moved after Apr, is Feb:Apr; Feb:Mar with Feb gone is Mar!); one moved but still among them, and any
      other sheet, leaves it as it was; a sheet moved or inserted between the ends is read. Renaming an end renames it.
    - The range operator between references (=SUM(A1:INDEX(A:A,3)), =OFFSET(A1,1,0):A4, =INDEX(..):INDEX(..)):
      the range holding both, measured on 25 formulas, all as Excel's, then the probe's 9. Tightest of all
      (=A1:B2 B2:C3:D4 is B2; =-A1:INDEX(A:A,2) spills -A1 and -A2); references on two sheets, or a 3D one, #VALUE!;
      Jan:'My Mar'!A1 is Jan, a name, to 'My Mar'!A1, so #NAME?; a number or text on one side isn't a formula.
      A formula whose range is worked out (INDEX's) is worked out at every change, as the cells between can be any.
    - Negation over a range of several cells spills in a dynamic formula, =-A1:A2 as -A1 and -A2.
    - Rows or columns inserted or deleted, or cells cut and pasted, on one of its sheets leave a 3D reference as it
      was: it reads what is there now (8 cases measured, all as Excel's).
  - `functions-amort.js`: AMORLINC and AMORDEGRC, the French depreciation. Microsoft's page gives other rules than
    Excel's; these are Excel's, from about 600 cases with cell values in a hidden Excel, then the probe's 48:
    - AMORDEGRC's coefficient by the life (1 / rate): up to 2 years #NUM!, to 4 years 1.5, to 6 years 2, past that 2.5.
    - The life has ceil(1 / rate) periods, the first (prorated) one with them. The last two take half of what is left
      each, and the ones after it nothing. A period starting below salvage takes nothing.
    - Each period's depreciation is rounded as returned, but the value it leaves isn't, except the first period's
      (5000 bought January 1 at 0.3: 1119, then 1746, not 1747).
    - The first period's part of a year: bases 0 and 4 as YEARFRAC; bases 1 and 3 actual days, February 29 counted as
      the 28th, over 365, or for basis 1 over the purchase year's days. None (bought on the first period's last day):
      a whole year, and for AMORDEGRC one period fewer.
    - basis 2 is #NUM!, as are a purchase after the first period, salvage above cost, a negative period, a rate of 0.
  - `functions-regex.js`, `pcre.js`, `pcre-match.js`, `regex.js`: REGEXTEST, REGEXEXTRACT and REGEXREPLACE. Excel's
    patterns are PCRE2's with Unicode properties. `pcre.js` reads a pattern into a tree, `regex.js` turning each
    escape or class that stands for one character into JavaScript's class syntax (the v flag), and `pcre-match.js`
    matches it by backtracking, each one-character class a JavaScript test that can't backtrack itself. Measured on
    about 190 cases with a hidden Excel, then the probe's 125, then 924 more over the features JavaScript hasn't got
    (`tools/excel-compare/regex-cases.mjs`, Excel's answers in test/regex-excel.json; the last 600 random patterns,
    held out): 916 of the 924 as Excel's, where 287 of the 917 that finished were before (the other 7 didn't finish):
    - \w, \d, \s and \b are Unicode's (é is a letter, a no-break space a space, ٣ a digit). $ matches before a final
      newline too, . anything but a newline. (?i), (?s), (?m) and (?x) at the start apply to the whole pattern.
    - Read too: [[:alpha:]] and the other POSIX classes, \Q...\E, (?P<name>...) and (?P=name), \g1 and \g{-1},
      \x{..}, \o{..}, octal, \cJ, \h, \R, \N, \N{U+41}, a{,2}, comments (?#...), atomic groups and possessive
      quantifiers (^a++a matches nothing in "aaa", as in Excel), \p{Greek}, \p{Script=Latin}, \p{Xwd}, \p{Xan},
      \p{Xsp}, \p{Xuc}.
    - PCRE2's own, as Excel has them: \K (=REGEXEXTRACT("price: 42","price: \K\d+") is "42"; REGEXREPLACE replaces
      from \K on); recursion and subroutine calls ((?R), (?1), (?-1), (?+1), (?&name), (?P>name), \g<1>), each call's
      captures put back afterwards, backtracking into it allowed; conditionals on a group ((?(1)..|..), (?(<name>)..)),
      on recursion ((?(R)..), (?(R1)..)), on a lookaround, and (?(DEFINE)..); \G, where the search began; \X, a
      grapheme cluster (an emoji ZWJ family, a flag); branch reset (?|..); options anywhere, for the rest of the group
      ("a(?i)b" matches "aB", not "AB"; (?i:..), (?-i), (?^), (?n), (?U), (?J), (?xx)); the verbs (*FAIL), (*ACCEPT),
      (*COMMIT), (*PRUNE), (*SKIP), (*THEN), (*MARK:x); the start options (*UTF), (*UCP), (*LIMIT_MATCH=n) and the
      newline conventions (*CR), (*LF), (*CRLF), (*ANYCRLF), (*ANY), (*NUL); lookbehind of bounded length
      ((?<=a{1,3})b). A backreference to a group that took no part matches nothing.
    - Case-insensitive, \p{Lu}, \p{Ll} and \p{Lt} match any cased letter (REGEXTEST("a","\p{Lu}",1) is TRUE), but
      [[:upper:]] and [[:lower:]] keep their case (REGEXTEST("a","[[:upper:]]",1) is FALSE).
    - After an empty match the next search starts there and may not be empty where it starts; finding nothing, it
      starts again a character on: =REGEXREPLACE("abc","\G","-") is "-a-b-c-", =REGEXREPLACE("abc","(?!\G)","-")
      "a-b-c-".
    - **A limit on the work**: a match taking more than 5,000,000 steps, or recursions 2,000 deep, is #VALUE!, as
      Excel gives at PCRE2's match limit (=REGEXTEST(REPT("a",30)&"b","^(a+)+$") and
      =REGEXTEST(REPT("ab",20)&"c","^(ab|a|b)*$") are #VALUE! in Excel, and here in well under a second). A plain
      character every match needs, missing from the text, is no match at once, as PCRE2 finds it
      (=REGEXEXTRACT(REPT("a",25)&"x","(a*)*b") is #N/A). So no pattern a workbook brings can hold the sheet up.
    - #VALUE!, as PCRE2 refuses them: an unknown escape (\i), \C, \u0061, \U, \L, an unclosed group, a bad repeat, a
      repeat past 65535 or of nothing (after (?i) or |, or of a repeat: "c{2}(?R)?*?"), a lookbehind of unbounded
      length ((?<=a+)b), \K in a lookaround, a name used twice without (?J), a conditional with three branches, one
      on a group the pattern hasn't got, callouts (?C), [[.a.]] and [[=a=]], a range from or to a set ([\d-z],
      [a-\d]), \p{IsLatin}, \8, a recursion that would loop ((?R) alone). case_sensitivity is 0 or 1, else #VALUE!.
    - REGEXEXTRACT: mode 0 the first match (#N/A if none); 1 every match, in a row; 2 the first match's groups, in a
      row, a group that took no part #N/A; #VALUE! where the pattern has no group, #N/A where it has and nothing
      matches. ARRAYTOTEXT writes an error in an array as its text ("1, #N/A"), an error alone is that error. An array
      given it is cut
      to its first value. REGEXTEST and REGEXREPLACE work out each value of an array.
    - REGEXREPLACE: occurrence 0 every match, n the nth, -n the nth from the end. The replacement is PCRE2's extended
      one: $1, ${1}, ${name}, $name, $0 and $&, $$; \1; \n, \t, \x41; \u, \l, \U...\E, \L...\E; ${2:-default}
      (where group 2 took no part) and ${2:+set:unset}. A group the pattern doesn't have, "$" or "\" alone, and an
      unknown escape (\q) are #VALUE!.
  - `trig.js`: sine and cosine as Excel works them out. The angle is brought near 0 with the x87's 64-bit π, not the
    true one, so =SIN(PI()) is 1.22514845490862E-16 and =COS(PI()/2) 6.1257422745431E-17 (the true values are
    1.22464679914735E-16 and 6.12323399573677E-17). From 2^27 on, #NUM! (=SIN(134217728)). TAN is SIN over COS
    (=TAN(PI()/2) is 1.63245522776191E+16), COT COS over SIN. SIN, COS, TAN, SEC, CSC, COT and the complex functions
    use it.
  - Table references (`ooxml/tables.js`): Fruit[Qty] is the column's data rows; [#All], [#Data], [#Headers], [#Totals]
    and two of them together; [[Qty]:[Price]] the columns between; [#This Row] (shown [@Qty]) the formula's own row,
    from any sheet by the row's number, #VALUE! past the table's rows (Excel's answers in test/formulas-excel.json).
    Names compare case aside. Typing a table reference into a cell isn't read yet.
  - `values.js` defines the values and how each is coerced into another.
- **Recalculating** (`engine.js`):
  - **What a change reaches.** On the first change every worksheet's formulas are read, with what each one reads,
    indexed by column. A change then works out again the formulas it reaches, along with volatile ones (NOW, TODAY,
    RAND, OFFSET, INDIRECT).
  - **Order.** Each formula is worked out after the ones it reads. A depth-first walk with its own stack does this,
    so a column of 10,000 chained formulas doesn't overflow the call stack.
  - **Results.** Each result is written into the cell's `<v>`, and its `t` set to `str`, `b` or `e`. The formula
    itself is kept as written.
- **Formulas Excel saved are trusted until something they read changes,** as Excel trusts them when it opens a file.
- **Opening.** Formulas the file came without a result for are worked out (`Workbook.calculateMissing`), with the
  formulas that read them, as Excel does when it opens such a file; the results saved with the others are kept. They
  are shown at once, and saved with a sheet once it is edited, so a workbook saved unchanged is the same bytes.
- **Shared formulas.** A shared formula's later cells carry only `si`. They read the group's first formula, moved to
  their own position (`shiftFormula`). When one cell of the group changes, every cell is first written out with its
  own formula.
- **Array formulas** (`t="array"`) fill their range. One cell of an array can't be changed, as in Excel. Data
  tables keep the values Excel saved.
- **calcChain.xml** is dropped when a formula is added or removed. A chain that names a cell which no longer holds a
  formula makes Excel offer a repair; without a chain, Excel works out the order again.
- **What the engine can't work out:** an unknown function, a link to another workbook, R1C1 in INDIRECT, a data
  table. These cells keep the value Excel saved (a new entry shows `#NAME?`), and the saved workbook gets
  `<calcPr fullCalcOnLoad="1"/>` so that Excel works everything out again when it opens the file.
- **Typing a formula** (`normalizeFormula`):
  - function names, references and TRUE/FALSE go into capitals
  - newer functions get the `_xlfn.` prefix the file needs, and lose it again on screen (`displayFormula`)
  - missing closing brackets are added
  - a formula that doesn't parse is refused with Excel's words, "There's a problem with this formula", and stays in
    the editor
- **Pointing** (`src/editor/pointing.js`), as Excel's Point mode works:
  - **Where.** A reference can go after `=`, an operator, `(`, `,` or `:`, before the end, `)` or another operator,
    and not inside quotes.
  - **Arrow keys.** In Enter mode (typing over a cell) they point: from the active cell, Shift extends, Ctrl jumps
    over data. In Edit mode (F2, which switches between the two) they move the caret.
  - **Mouse.** A click or a drag on the grid points too, and Shift+click extends.
  - **Ending.** Typing anything else ends the reference, and the next arrow starts a new one.
  - **F4** cycles the reference at the caret: A1, $A$1, A$1, $A1.
  - **Range finder.** The formula's ranges are outlined in colours, and the one being pointed at is dashed. Excel's
    own colours aren't in any file to measure, so these are ours. The reference text isn't coloured in the box.
  - **Other sheets.** A sheet tab clicked where a reference can go shows that sheet. The sheet's name goes into the
    formula at once (`'Notes & more'!`), and the formula goes on in the formula bar. The arrows start from that
    sheet's active cell. Enter puts the formula into its own cell, on its own sheet; Esc goes back without it.
  - **Merged cells.** A merged cell is pointed at by its top-left cell. Excel's behaviour here isn't measured.

## Formula Auditing

Show Formulas (Formulas > Formula Auditing, Ctrl+`), as Excel does it (measured with a hidden Excel's
Window.DisplayFormulas, its Range.Text and the file it saved):

- Saved as `showFormulas="1"` on the sheet's view; a sheet opened shows formulas when its file says so.
- Each column is drawn twice as wide (Excel's 54.75 points to 109.5), the width saved unchanged.
- Each cell shows what it holds, unformatted, as the formula bar has it: =A1*2, 1234.5 for 1,234.50, 45292 for a
  date, 0.25 for 25%, 007 for '007, 100000000000000000000 for 1E+20; an array formula without its braces. Formulas sit
  left, as text does, unless the cell is aligned otherwise.

Trace Precedents and Trace Dependents (`ooxml/trace.js`, `editor/tracing.js`, `editor/arrows.js`), as Excel's arrows
go (measured with a hidden Excel's ShowPrecedents, ShowDependents and NavigateArrow on
=A1+A2+SUM(B1:B3)+Sheet2!A1+Rate+A1 and its dependents; `test/trace.test.js`):

- One arrow for each different cell or range read: A1 read twice is one arrow. A defined name is traced to what it names
  (Rate to D1). Everything on other sheets is one arrow of its own, from a worksheet icon (to it, for dependents).
- Each click goes a level further, from the formulas the last level reached, until there is nothing more. A trace from
  another cell starts afresh, the arrows drawn so far staying. Remove Arrows takes them all away; Remove Precedent
  Arrows and Remove Dependent Arrows those of one kind.
- Drawn as Excel draws them: a dot where the arrow starts and a head where it ends, a range boxed with the arrow from
  its top-left cell, blue, red from a cell holding an error, dashed black to or from another sheet. Each sheet keeps
  its own arrows; none is saved (Excel saves none).
- **Not yet**: Excel takes arrows away when a formula they point to changes, rows or columns are inserted or cells
  moved; here they stay until removed. The arrows' exact look (Excel's line widths, its icon) isn't measured. Double-
  clicking an arrow to go to its other end. How Excel aligns numbers in Show Formulas isn't measured (Range.Text can't
  say); here General's right.

Error indicators (`ooxml/errorcheck.js`): Excel's green triangle in the top-left corner of a cell that breaks one of
the rules Excel checks by default, each cell's message heard as its description (measured with a hidden Excel's
Range.Errors on about 60 cells; `test/errorcheck.test.js` has them):

- **Evaluates to an error**: a formula whose result is an error, but #N/A from a formula calling NA() (=NA(), =NA()+1;
  =1/0+NA() is flagged, and so is =A1 reading an =NA()). An error typed as a value isn't.
- **Number stored as text**: text that reads as a number ('123, '12.5, '1,234, '1e3, ' 7, '50%, '(5), '-3, '0012), not
  a date or time ('1/2/2024, '12:30), 'TRUE, or the text a formula gives.
- **Two-digit year**: text that reads as a date with a two-digit year ('1/2/24, 'Jan 2, 24, '2-Jan-24).
- **Inconsistent formula**: the cells either side, left and right or above and below, hold one formula between them
  (moved as a copy moves it) and this cell another (=E3*F3 between =E2+F2 and =E4+F4).
- **Formula omits adjacent cells**: a number (a date too; not text, TRUE or a formula's result) just before or after a
  range of one row or one column the formula reads, on whatever sheet that is (=SUM(N2:N4) with a number in N1). A
  range of two rows and columns isn't checked, nor a number past a gap.
- **Unlocked formula**: a formula in a cell whose format isn't locked.
- A sheet's ignoredErrors (Excel's Ignore Error, saved in the file) turn a rule off for their cells.
- **The error button** beside the active cell when it has an indicator (`editor/errormenu.js`, `ooxml/errorfix.js`;
  Alt+Shift+F10 too): Excel's menu, headed by the error (Number Stored as Text, Divide by Zero Error...), with its fixes
  (Convert to Number; Copy Formula from Above or Left for an inconsistent formula; Lock Cell), Ignore Error and Edit in
  Formula Bar, for the selected cells with that error, each one step of Undo. Ignore Error is written as Excel writes
  it (measured): one ignoredError a rule, its cells joined into ranges, `<ignoredError sqref="A1:A2 A4"
  numberStoredAsText="1"/>`, after pageMargins.
- **Error Checking** (Formulas > Formula Auditing; `editor/errorbox.js`): Excel's box going through the sheet's cells
  with an indicator in reading order (rows, then columns), from the active cell, each selected in turn: "Error in cell
  G2", what it holds, the error's name and what it means, its fix, Ignore Error, Edit in Formula Bar, Previous, Next;
  a fix or Ignore Error goes on to the next. Past the last it starts again at the top, and back where it started says
  "The error check is complete for the entire sheet." The box is a dialog COM can't drive: its order and words are
  Excel's as its documentation and the box itself show them, not measured; Excel's box lets the sheet be clicked
  while it is open, this one doesn't.
- **Not yet**: the two-digit year's Convert XX to 19XX and 20XX, Update Formula to Include Cells, Show Calculation Steps,
  Help; whether Convert to Number changes a Text number format (not measured: it is kept).
  The rules Excel has off by default (formulas referring to empty cells). One case Excel didn't flag here does: =SUM(M1:M3)
  in M5 with a number in M4, when M7 held =M1+M2 (alone, Excel flags it). The triangle's green isn't measured.

Error Checking Options (the error button's menu, or Options... in the Error Checking box): Excel's own, Enable
background error checking and each rule here in Excel's words (File > Options > Formulas). A rule turned off isn't
checked, and with background checking off no triangles are drawn. As Excel keeps them as its own settings, not in the
workbook, the editor keeps them in the browser; `errorChecking` sets them (`false`, or { background, rules }).

Evaluate Formula (`calc/stepper.js`, `editor/evalbox.js`; `test/stepper.test.js`): the active cell's formula worked out
a step at a time, the next part underlined and the value just worked out in italics; Evaluate, Restart at the end.
The steps go as Microsoft's own example shows them (=IF(AVERAGE(D2:D5)>50,SUM(E2:E5),0) to =IF(40>50,...), to
=IF(FALSE,...), to 0): a reference on its own or as an operand is a step of its own (its value; a range's as an array);
one handed to a function goes with it; a function or operator once its arguments are values; IF works out its
condition, then only the argument it returns, IFERROR its value_if_error only after an error; brackets go with what
is in them. Excel's box is a dialog COM can't drive, so the steps aren't measured beyond that example.

- **Step In** (when the part underlined is one cell holding a formula): that cell's formula on a line of its own
  beneath, named ("Sheet1!$B$1 = "), worked out with Evaluate; **Step Out**: back to the formula above, the cell's
  value put in place of its reference. As Excel's box does (used in its window: each level its own line, the cell
named).
- **Step In on a cell holding a constant**: its value on a line of its own, "The cell currently being evaluated
  contains a constant.", Evaluate and Step In off, Step Out to go back (seen in Excel's own box: =B1+1 into =A1*2 into
  5).
- **Not yet**: a defined name's own steps.

The Watch Window (`ooxml/watches.js`, `editor/watchwindow.js`; `test/watches.test.js`): the watched cells with their Book,
Sheet, Name (a defined name for just that cell), Cell, Value (as the cell shows it) and Formula, kept up to date as
the workbook changes; Add Watch watches the selected cells, each cell of a range on its own as Excel's does; Delete
Watch; a double click or Enter goes to the cell. Saved as Excel saves them (measured: Application.Watches.Add, the
file saved and opened again): a sheet's `<cellWatches><cellWatch r="B2"/>...`, in the order added, after pageMargins.

- **Not yet**: Excel's Add Watch box (here the selection is watched at once); the window floating (it is docked under
  the ribbon here); watches of other open workbooks. Adding or deleting a watch isn't a step of Undo, as in Excel.

## Calculation options

Formulas > Calculation (`src/ooxml/calcmode.js`): Calculation Options (Automatic, Automatic Except for Data Tables,
Manual), Calculate Now (F9), Calculate Sheet (Shift+F9); Ctrl+Alt+F9 works out every formula. Measured in a hidden
Excel (and checked the same way by `test/calcmode.test.js`):

- **Saved** in workbook.xml's calcPr: Manual as `calcMode="manual"`, Automatic Except for Data Tables as
  `calcMode="autoNoTable"`, Automatic as no calcMode; Recalculate workbook before saving off as `calcOnSave="0"`. A
  workbook opened is in the mode its file says.
- **Manual**: a change works out nothing that reads it (=A1*2 stays 10 when A1 goes from 5 to 7, and so does a formula
  on another sheet), and the status bar says Calculate. A formula entered is worked out at once, from what the cells
  hold then (=A2+1 is 11 while A2 still shows 10). Calculate Sheet works out the shown sheet's (A2 14), the other
  sheets' wait; Calculate Now the rest. Back to Automatic, everything waiting is worked out.
- **Saving** a manual workbook works it out first (`save()` in the editor and the agent), unless calcOnSave is off.
- There are no data tables here (What-If Analysis), so Automatic Except for Data Tables works as Automatic.
- **Not yet**: inserting or deleting rows and columns, and moving or copying sheets, still work out their formulas in
  Manual (Excel leaves them waiting); the Options box's iteration and "Recalculate before saving" (calcOnSave can be
  set through `setCalcOnSave`).

## The Function Library

- **The list** (`calc/catalog.js`): every function the engine works out, in Excel's categories as Microsoft's "Excel
  functions (by category)" has them, and its syntax with Excel's argument names, the optional ones in brackets
  (`SUM(number1, [number2], ...)`). Where that list names a function under two categories, it is in one: CONCATENATE
  in Text, FORECAST in Statistical, FLOOR in Math & Trig, LET in Logical. DATEDIF is left out, as Excel leaves it out of
  its lists. `test/catalog.test.js` checks that the list and the engine agree.
- **Insert Function** (Shift+F3, or fx on the Formulas tab; `editor/functionbox.js`): Excel's box. Search for a
  function and Go lists the functions whose names hold the words, under Recommended. A category shows its functions:
  Most Recently Used (Excel's list to start with: SUM, AVERAGE, IF, HYPERLINK, COUNT, MAX, SIN, SUMIF, PMT, STDEV,
  then the ones picked, newest first), All, or one of Excel's. The syntax and description of the one picked are under
  the list. OK goes on to Function Arguments, as in Excel. In a cell being edited, Shift+F3 keeps it open.
- **Function Arguments** (`editor/fnargsbox.js`, `calc/fnargs.js`): after Insert Function, a function picked from the
  Formulas tab's lists, Home or Formulas > AutoSum > More Functions..., or Ctrl+A right after a function's name typed in
  a cell being edited (`=VLOOKUP`, `=VLOOKUP(`, its arguments read into the boxes). A box per argument, named as Excel
  labels them (Lookup_value), the required ones in bold; a repeated argument gets its next box once the last is filled
  (Number3 after Number2; SUMIFS's Criteria_range2 and Criteria2 together), up to Excel's 255. Beside each, its value at
  the cell (`= {1;2;3}`, `= "abc"`, `= TRUE`); under them the function's description, the description of the argument
  with focus, and `Formula result =` for the whole call. OK writes the call (at the caret, or in place of the name and
  arguments it opened on; `=` first in an empty cell) and enters the cell as one step of the history; Cancel leaves the
  cell as it was. The command is `functionArguments` (`insertFunction` leads into it).
  - **Pointing at the sheet** (`editor/pointing.js`, the same pointing as a formula typed in a cell): the box doesn't
    hold the page, so while an argument has focus a click, a drag (Shift+click extends) or the arrow keys (Ctrl jumps,
    Shift extends; from the active cell) put the reference in it, where the box is empty, its text selected (the
    reference replacing it, as when Tab moves into a box) or the caret where a reference can go; F4 cycles the $
    signs. A sheet's tab clicked shows that sheet to point at (`'Notes & more'!B2`); the cell being edited is hidden
    there, and the cell's own sheet comes back when the box closes. The references in the box are outlined on the
    sheet in their colours. The ribbon and the formula bar wait while the box is open; the box sits at the lower right.
    Each argument has Excel's Collapse Dialog button (out of the Tab order), which shrinks the box to that argument's
    text box alone, and Expand Dialog back. Excel also collapses the box by itself while a range is dragged over; here
    it stays as it is.
  - **Hints**: an empty box shows the kind of value it takes, `= number`, `= text`, `= logical`, `= reference`,
    `= array` or `= any`, as Excel's box shows them (`calc/argkinds.js`). Excel's own list of each argument's kind is
    in no file or COM property, so the kinds are given by what each argument means; not measured. A filled box shows
    its value as Evaluate Formula writes it: `= 6`, `= "abc"`, `= TRUE`, `= {1,2,3}` for a row and `= {1;2;3}` for a
    column.
  - **Quotes** (`needsQuotes` in `calc/fnargs.js`): text typed in a text or any box (LEFT's text, COUNTIF's criteria,
    IF's values) is put in quotes when focus leaves the box and when the formula is written, if it isn't a formula:
    `>1` becomes `">1"`, `yes` `"yes"`, `x"y` `"x""y"`. That Excel's box does this is Microsoft's description of it;
    the box can't be driven over COM without sending keys, so which boxes Excel quotes in was not measured. What is
    measured is the rule for "isn't a formula" (`tools/excel-compare/excel-fnquote.ps1`, 58 texts typed after `=` in a
    hidden Excel, `test/fnquote-excel.json`, checked by `test/fnquote.test.js`): Excel refuses `>1`, `<>`, `1,000`,
    `$5`, `*`, `?a`, `1 2` and takes the rest, exactly as this parser does (58 of 58); of those it takes, plain words
    (`abc`, `apple pie`, `Jan`, `XFE1`, `a-b`) are names it doesn't know (`#NAME?`), so they are quoted too, but not a
    defined name or table (`Rate`), a reference (`A1`, `'My Data'!A1`, `Q1`), a number, TRUE, an error value, an array
    or a call (`SUMM(1)`, a function misspelt, stays as typed). Two differ: `R1C1`, which Excel reads as `$A$1`, and
    `N/A`, which it works out to #VALUE!; both are quoted here. LAMBDA, LET and the functions that name their own
    parameters quote nothing. Number, logical, reference and array boxes keep what is typed (`=SUM(abc)`).
- **Descriptions** (`calc/describe.js`): one sentence for every one of the 499 functions listed, and for each argument
  (by its name, or the function's own where the name means something else: PMT's type against T.TEST's), written for
  this project in its own words, not Microsoft's.
- **The Formulas tab's lists**: Recently Used, Financial, Logical, Text, Date & Time, Lookup & Reference, Math & Trig,
  and More Functions (Statistical, Engineering, Information, Compatibility). Database is in Insert Function only, as
  in Excel. A function picked is typed as Insert Function types it. Each one's syntax is its tooltip.
- **The screen tip** (`editor/fntip.js`): inside a function's parentheses in a cell being edited, its syntax under the
  cell, the argument the caret is in in bold. Repeated arguments go on with the last one named (SUMIFS's pairs two by
  two). Text in quotes, sheet names and array constants are passed over.
- **What the box writes, measured** over COM in a hidden Excel, each formula set as a cell's formula
  (`tools/excel-compare/excel-fnargs.ps1`, 47 formulas, `test/fnargs-excel.json`; `excel-fnmin.ps1`, every listed
  function, `test/fnmin-excel.json`; checked by `test/fnargs.test.js`):

  | Rule | Seen |
  |---|---|
  | An argument left empty stays an empty place, kept as written. | `=SUM(A1:A3,,5)` is 11, `=IF(A1>1,,"no")`, `=PMT(0.01,12,1000,,)`, `=VLOOKUP(,,)` all taken as typed. |
  | Too few arguments are refused, empty places counting as arguments. | `=VLOOKUP(2,A1:A3)`, `=MID(B1,2)`, `=ROUND(1.25)` refused; `=VLOOKUP(,,)`, `=MID(B1,2,)`, `=ROUND(1.25,)` taken. |
  | The fewest arguments are the syntax's unbracketed ones, for 491 of the 499. | IF needs 2 (`=IF(A1>1)` refused, `=IF(A1>1,)` taken); DCOUNT, DCOUNTA, PROB, REDUCE and SCAN 3; BYCOL and BYROW 1. |
  | Nothing in a function that needs one argument can't be written. | `=SUM()`, `=LEFT()`, `=AVERAGE()`, `=COUNT()`, `=IF()` refused. |

  So the box writes the arguments up to the last one filled, and at least the fewest Excel takes, the empty ones as
  empty places: VLOOKUP with nothing filled is `=VLOOKUP(,,)`, IF with only a test `=IF(A1>1,)`. OK on a function
  needing one argument with nothing in it (`=MAX()`) leaves the cell being edited with that, as AutoSum leaves
  `=AVERAGE()` (docs/AUTOSUM.md), since Excel would refuse it. The box itself is drawn as Excel's is laid out; what
  Excel's own box writes was not driven over COM (it would need keys sent to a dialog), only the formulas it can write.
  The engine took too few arguments without a word; it still does when typed (Not yet).
- **Not yet**: Function Arguments collapsing by itself while a range is dragged over, moving the box by its title,
  opening it from the formula bar on a cell's formula, and refusing a typed formula with too few arguments as Excel
  does.

## What Excel does (measured)

| Rule | Seen |
|---|---|
| A number becomes text with 15 significant digits. It is written out in full when that takes at most 20 characters, and in scientific notation otherwise. | `=1/3&""` is "0.333333333333333"; "0.000000000000000001", but "1E-19"; "0.000123456789012345", but "1.23456789012345E-05"; "12345678901234500000", but "1E+20". |
| A number written in a formula keeps 15 significant digits; the rest are cut off, not rounded. | `=1234567890123456` is 1234567890123450. |
| When a formula's last operation is + or -, and its two operands agree to 15 significant digits, the result is 0. Brackets, or any later operation, stop this. SUM does the same at each step. | `=0.1+0.2-0.3` is 0, but `=(0.1+0.2-0.3)` and `=1*(0.1+0.2-0.3)` are 5.55E-17. `=1+1E-15-1` is 0, but `=1-1E-15-1` is -9.99E-16. `=SUM(0.1,0.2,-0.3)` is 0. |
| Comparisons treat numbers that agree to 15 significant digits as equal. | `=0.1+0.2=0.3` is TRUE. |
| Outside an array formula, a range given to an operator is cut to the formula's own row or column (implicit intersection). An array constant stays whole. | In row 5, `=SUM(Data!B2:B9*1)` is Data!B5. `=SUM({1,2,3}*2)` is 12, and `={1,2,3}*2` shows 2. |
| SUMPRODUCT's arguments are worked out as arrays. | `=SUMPRODUCT((B2:B9>3)*C2:C9)`, `=SUMPRODUCT(--ISNUMBER(SEARCH("an",A2:A9)))`. |
| A negative number to a fractional power is #NUM!, unless the power is an odd root. | `=(-8)^(1/3)` is -2; `=0^0` is #NUM!. |
| An argument left out counts in COUNT and COUNTA, and counts as 0 in AVERAGE. | `=COUNT(1,)` is 2; `=AVERAGE(1,)` is 0.5. |
| An equal-to criterion matches a number stored as text too; the other operators don't. | `COUNTIF(B2:B9,12)` counts the text "12"; `">=4"` doesn't. |
| A blank criterion with `=` matches blank cells only; with no operator, blank cells and "". A whole-column criterion counts every blank cell. | `COUNTIF(F:F,"")` is 1048576. |
| YEARFRAC's US 30/360 (basis 0) makes a start on the last day of February the 30th, but then leaves an end on the 31st as the 31st; only a start on the 30th or 31st itself makes it the 30th. | `=YEARFRAC(DATE(2020,2,29),DATE(2020,12,31),0)` is 0.836111 (301 days), `=YEARFRAC(DATE(2021,2,28),DATE(2021,3,31),0)` 0.086111 (31). |
| MOD works on the binary values, without the 15-digit rounding. | `=MOD(0.3,0.1)` is 0.09999999999999998. |
| TIME is the fraction of a day with whole days taken off. | `=TIME(25,0,0)` is 0.04166666666666674. |
| ROUND, ROUNDUP and ROUNDDOWN work on the value's 15 significant decimal digits. | `=ROUND(2.675,2)` is 2.68. |
| Text compares without regard to case, in Excel's order. Numbers come before text, and text before FALSE and TRUE. A blank counts as 0, "" or FALSE, whichever the other value is. | `="a"="A"`, `=1<"a"`, `="a"<TRUE`, `=Data!F2=""` (blank) are all TRUE. |
| Approximate lookups on a list sorted up (MATCH 1, VLOOKUP and HLOOKUP TRUE) are a binary search, which on a list that isn't sorted gives its own answer. From the middle: a blank or a value of another type is passed over to the right; an exact match is taken at once, then the last of the equal values after it. On a list sorted down (MATCH -1) it is a scan from the start, blanks and other types passed over, taking the first exact match and stopping at the first value below. | 400 unsorted lists, some with blanks and text, 1,200 answers (`test/lookup-excel.json`, from `tools/excel-compare/excel-lookups.ps1`): all as Excel's. `MATCH(8,{3,6,9,6,6,1,8,7},1)` is 7, not 2 or 8; `MATCH(6,{5,8,7,1},-1)` is #N/A. |

## The byte functions, ASC and DBCS

Measured in Excel 365 in English (language 1033, not a double-byte one), 62 cases built with UNICHAR so no encoding
stands between, each result read a UTF-16 unit at a time:

- **LENB, MIDB, FINDB, SEARCHB, REPLACEB** count UTF-16 units: LENB of an emoji is 2, of two CJK characters 2, and
  MIDB and REPLACEB cut an emoji in two (test/bytes.test.js, as the probe's JSON can't carry half of one).
- **LEFTB and RIGHTB** keep a character whole, as Excel 365's LEFT and RIGHT do: LEFTB of an emoji and 1 is the emoji.
- **ASC and DBCS** give the text as it is: full-width ＡＢＣ stays full-width, ABC stays half-width. They convert only
  where a double-byte language is Excel's default.
- Errors as LEFT, MID, FIND, SEARCH and REPLACE give them.

## Characters: CHAR, CODE, UNICHAR, LEFT and RIGHT

Measured in Excel 365 in English, 96 cases (`src/calc/charsets.js`):

- **CHAR and CODE are code page 1252**, not Unicode: CHAR(128) is €, CODE("€") 128, CODE("Œ") 140, CODE of the curly
  quotes 145 to 148. A character 1252 hasn't got is 63, "?": Ā, α, 日 and an emoji, with no nearest letter taken. The
  control characters U+0080 to U+009F keep their own numbers.
- **UNICHAR refuses what isn't a character** with #N/A: the surrogates (U+D800 to U+DFFF), U+FDD0 to U+FDEF and the last
  two code points of every plane (U+FFFE, U+FFFF, U+1FFFF ... U+10FFFF). Out of range is #VALUE!, as before.
- **LEFT and RIGHT keep an emoji whole**: LEFT of an emoji and 1 is the emoji, LEFT("ab"&emoji,3) all of it. How LEN,
  MID, FIND, SEARCH and REPLACE count one depends on the workbook: see the next section.

## February 29, 1900

Excel's calendar keeps Lotus's leap day, serial 60, but not everywhere alike. Measured in Excel 365 over COM, 513
formulas (`tools/excel-compare/date1900-cases.mjs`, Excel's answers in test/date1900-excel.json, 508 as Excel's; 384
before):

- **DATE counts the days on from the month's first in a calendar with the 29th**: DATE(1900,2,29) is 60,
  DATE(1900,2,30) and DATE(1900,1,61) 61, DATE(1900,3,0) 60. A month before January 1900 is #NUM! whatever the day
  (DATE(1900,0,64)), while DATE(1900,1,0) is 0.
- **DAYS360 reads serials in that calendar too**, February 1900's last day the 29th: DAYS360(59,61) is 3,
  DAYS360(60,61) 1, DAYS360(31,60,TRUE) 29 (each a day off here before, the serials below 61 read a day early).
- **EDATE, EOMONTH and the coupon dates step over a February 1900 of 28 days**: EDATE(31,1) and EOMONTH(40,0) are
  59, EDATE(244,-6) 59; serial 60 still reads as February 29 (MONTH(60) is 2).
- **Not yet: coupon dates before 1900.** COUPDAYS(100,244,2,1) is 185 in Excel (182 here), COUPPCD(61,1000,4,0) 0
  (-5 here), COUPDAYBS(50,881,4,0) 49 (79 here): Excel seems to take a coupon date before 1900 as serial 0; not
  worked out.

## Emoji: the function compatibility version

Measured in Excel 365 in English over COM, 1972 formulas built with UNICHAR (`tools/excel-compare/emoji-cases.mjs`,
`excel-formulas.ps1`; Excel's answers in test/emoji-excel.json), each in basic.xlsx as this editor writes it and in the
same workbook with the version added. The last 600 are random strings of other characters, a held-out check: **1972
of 1972 as Excel in both** (`src/ooxml/calcmode.js` `compatVersion`, `src/calc/functions-text.js`).

- **What decides it** is workbook.xml's `<extLst><ext uri="{D14903EA-33C4-47F7-8F05-3474C54BE107}">` holding
  `<xlwcv:version setVersion="2"/>` (namespace `.../spreadsheetml/2024/workbookCompatibilityVersion`), which Excel 365
  writes into every workbook it makes. With setVersion 2 (or 3), LEN(UNICHAR(128512)) is 1; with 1 or none, 2. Nothing
  else in the workbook matters (calcId, fileVersion, app.xml, the theme, styles, calcChain and the calcFeatures list
  were each tried before and after).
- **Version 2 counts code points, not what a reader sees as one character**: a ZWJ family (man, ZWJ, woman, ZWJ,
  girl) is 5, a flag 2, an emoji with a skin tone 2, a letter and a combining accent 2, ❤ and its variation selector 2.
  Only LEN, MID, REPLACE, FIND and SEARCH change. LEFT and RIGHT keep a pair whole either way; LENB, MIDB, FINDB,
  SEARCHB, REPLACEB, LEFTB and RIGHTB count UTF-16 units either way; CODE, UNICODE, SUBSTITUTE, EXACT, TEXTBEFORE,
  TEXTAFTER and TEXTSPLIT don't count.
- **SEARCH's ? is one UTF-16 unit in both**, so SEARCH("??",emoji) is 1 with version 2; but there a match starts only
  where a code point does: SEARCH("?b","a"&emoji&"b") is 3 without the version, #VALUE! with it.
- **Half an emoji** (its first unit alone, made with MIDB): with version 2 not counted at the very end of a text
  (LEN of it alone is 0, of it and "y" 2). UNICODE of it is #VALUE! in both; UNICODE of the second unit alone is its
  number (56832).
- **FIND may start just past the end, SEARCH may not**: FIND("","abc",4) is 4, SEARCH("","abc",4) and SEARCH("","")
  #VALUE!, in both.
- **UPPER and LOWER leave letters past U+FFFF alone**: UPPER of Deseret's small long i (U+10428) is itself.
- A workbook this editor makes has no version, as before, so it counts as Excel counts it in that workbook; one Excel
  made keeps its extLst byte for byte and is worked out as Excel works it out.

## CELL and INFO

Measured in Excel 365 in English (`src/calc/functions-info.js`, `src/calc/cellinfo.js`; 128 number formats in
test/cellinfo.test.js, 33 cases in the probe):

- **CELL("format")**: F and the decimals for fixed, `,` with a thousands separator, C for currency ($), P percent, S
  scientific, G for General, text, fractions and what isn't one of the rest; `-` after it when the negative section has
  a colour, `()` when the positive one has parentheses. A date or time by the order of its first parts, separators
  aside: day-month-year D1, day-month D2, month-year D3, month-day-year D4, month-day D5; h:mm:ss AM/PM D6, h:mm AM/PM
  D7, h:mm:ss D8, h:mm D9; one starting with the year, one part alone, or elapsed time ([h]) G. A locale prefix
  ([$-409]) counts as a day part, so [$-409]m/d/yy is D2 (as Excel gives it).
- **"color"** and **"parentheses"**: 1 or 0, as above.
- **"prefix"**: for text typed into the cell, ' (General, left, justify), " (right), ^ (centre, centre across), \ (fill);
  "" for numbers and for a formula's text.
- **"type"**: b blank, l text (a formula's too), v anything else. **"contents"**: the value, 0 for a blank cell.
- **"address"**: $B$2; on another sheet with the workbook's name, '[Book1]Other Sheet'!$C$3. **"row"**, **"col"**: the
  first cell's. **"protect"**: 1 locked, 0 not. **"width"**: the column's width in whole characters and whether it is
  the sheet's default, two values ({8,TRUE}).
- **"filename"**: "", as Excel gives it for a workbook not yet saved; a page has no path to give.
- The info type in any case, but not trimmed; another, or a number, is #VALUE!. Without a reference, the formula's own
  cell (Excel takes the active cell).
- **INFO**: "numfile" the workbook's sheets, "origin" $A:$A$1, "osversion" Windows (64-bit) NT 10.00, "recalc"
  Automatic or Manual, "release" 16.0, "system" pcdos; "memavail", "memused" and "totmem" #N/A; another #VALUE!.

## ENCODEURL, BAHTTEXT and PHONETIC

Measured in Excel 365 in English, 51 cases (`src/calc/functions-words.js`; 48 in the probe):

- **ENCODEURL**: the text as UTF-8, each byte but a letter, a digit, - _ and . as %XX in capitals; ~ is encoded too
  (%7E), unlike JavaScript's encodeURIComponent. Written to the file as _xlfn.ENCODEURL.
- **BAHTTEXT**: the number in Thai words, rounded to satang as ROUND does (1.005 is one baht one satang); บาทถ้วน when no
  satang, only สตางค์ when no baht, ลบ before a negative. A 1 in the units place after a higher digit is เอ็ด (11, 101,
  1,000,001), 2 in the tens ยี่สิบ, 1 in the tens สิบ; past a million the groups repeat with ล้าน (1E+15 is a thousand
  million million).
- **PHONETIC**: #N/A. Excel 365 in English gives #N/A for every cell, even one whose text has a phonetic reading saved
  in the file (the reading is shown only with an East Asian editing language). Excel refuses =PHONETIC("abc") as typed;
  here it is #VALUE!.

## FILTERXML

Measured in Excel 365 in English, 55 queries (test/filterxml.test.js; 44 in the probe). `src/calc/xpath.js` is an
XPath 1.0 reader: every axis but namespace, name, * and node-type tests, predicates, the operators and the core
functions.

- What a query finds comes back in a column, one value a node: an element's own text (`<a>x<b>y</b>z</a>` is "xz",
  not its children's), an attribute's value. Each is read as a typed entry is (TRUE, 1,234, 50%, 1e3 and dates become
  values; " 12 " is 12), but text that would be an error or a formula stays text ("#N/A").
- #VALUE!: XML that doesn't parse (or a number given for it), no node found, a result that isn't nodes (count(),
  string(), concat(), name()), a text() node, and an element with no text of its own (an empty one, or `//g/..`) in
  its place in the column.
- Names are matched as written, case and prefix included (`//I` finds nothing); `//*` finds prefixed elements.
- In a formula of the old kind, Excel takes only the first value (as its `@` does with SEQUENCE: docs/DYNAMIC.md);
  here all of them reach the function around it.

## FORECAST.ETS and its family (not yet: what was measured)

FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY and FORECAST.ETS.STAT aren't here yet: how Excel picks
the smoothing parameters isn't found. What was measured, in a hidden Excel 365 over COM, on about 60 series with a
timeline of 1 to n and seasonality 0 (no season), so it needn't be measured again:

- **Matches Excel to 15 digits.** The model is Holt's additive trend in its error-correction form, beta Hyndman's
  (so at most alpha: Excel reports beta 0.099 with alpha 0.1): forecast f = l + b, error e = y - f, then l = l + b +
  alpha e and b = b + beta e. The trend starts at the least-squares slope of the values on the timeline, the level at
  the first value less it, so the first error is 0; for 3 values and for 7 or more (about 30 series). 4, 5 and 6 values
  start from other fixed weights on the values, measured but not explained: [-8, -3, 2, 9]/28, [-26, -13, -5, 13,
  31]/140 and [-90, -124, -53, 18, 89, 160]/980. h steps on, l + h b.
- **The error statistics** (STAT 4 to 7: MASE, SMAPE, MAE, RMSE) are over every one-step error for up to 10 values,
  and for 11 or more over the last third (rounded up) of the same run's errors only, a validation window. MASE is MAE
  over the mean of |y_t - y_t-1|. STAT 8 is the timeline's step.
- **CONFINT** without a season: NORM.S.INV((1 + confidence) / 2) x RMSE x the square root of 1 + the sum over j = 1 to
  h - 1 of (alpha + j beta)^2.
- **Not found: how alpha and beta are chosen.** Excel's answers sit on a few values (alpha 0.002, 0.1, 0.126, 1/6,
  0.25, 0.5, 2/3, 0.75, 0.9, 0.998; beta 0.001, alpha - 0.001, 1/12), and they aren't the least squared or absolute
  error over a 0.001 grid, whether over all the errors, the validation window or the first two thirds: near it for
  10 values or fewer, far from it for some longer series. Constant and straight-line data give 0.9 and 0.001.
- **Not found: a second model.** Some series are forecast as the last value plus the least-squares slope (alpha 1,
  beta 0), whatever STAT reports for them (10 values with alpha shown as 0.998, 13 with 0.666667).
- **Not started:** seasonality detection (FORECAST.ETS.SEASONALITY), the seasonal start values, data completion and
  aggregation.

## Not yet

- **ACCRINT with a date on the 29th, 30th or 31st, or in February, whose coupon dates pass a shorter February**:
  13 of the 400 cases first checked, and 25 of the 3951 near the ends of Excel's dates, differ from Excel's in the
  third or fourth digit (ACCRINT(DATE(2006,3,13),DATE(2006,8,29),DATE(2008,3,28),0.06,100,2,0) is 12.2809 in Excel,
  12.25 here; ACCRINT(2958065,2958434,2958100,0.1,1000,4,0) 9.44 in Excel, 10.28 here); kept out of the probe.
  Tried and dropped: stepping on from first_interest by whole multiples of the period instead of from the date before
  (one case better, two worse), and February 29, 1900 as February's last in the 30/360 days (six worse).
- **ACCRINT with issue on a coupon date, settlement before first_interest and basis 1, 2 or 3**: 13 of the 300
  differ (ACCRINT(DATE(2007,6,14),DATE(2010,6,14),DATE(2009,10,13),0.1,1000,2,3,TRUE) is 233.0137 in Excel, 4 periods
  and 120.5 days of 182.5, where the days are 121; 233.1507 here).
- **ODDLPRICE and ODDLYIELD from a last interest date at a month's end into a February 29 settlement** (August 31,
  2019 and February 29, 2020 measured): 1e-4 from Excel's for some frequencies and bases. Kept out of the probe.
- **AMORDEGRC bought on the first period's last day with a life of 3 years or less**: the rule here (declining, no
  halves, up to the life's periods) fits the two such assets measured, not more.
- **Regular expressions: 8 of the 924 cases measured** aren't Excel's. \G in an alternation after an empty match
  (=REGEXREPLACE("xxa","a|\G","<$0>") is "<>x<>x<a><>" in Excel, "<>xx<a><>" here: Excel seems to start that search
  a character on, as it doesn't for (?!\G)); and three random patterns repeating groups with conditionals, \K and
  \G inside, one of which Excel answers with a number 0. (*SKIP:name) is #VALUE! here (marks aren't kept). Where
  Excel stops for its match limit isn't the same step as here: a pattern needing between about 5 and 10 million
  steps may be worked out in Excel and #VALUE! here. A repeated group of 10,000s of iterations that isn't plain
  characters can run out of the stack, #VALUE! here.
- **IMCSC("1+i")**: Excel's last digits there (0.621518017170428-0.303931001628427i) come from neither form tried;
  here 0.621518017170429-0.303931001628426i. Other arguments agree.
- **Spilling** is in docs/DYNAMIC.md (dynamic arrays, LET, LAMBDA, `D1#`, `@`); its own gaps are listed there.
- **A function Excel doesn't know** (=foo(1), #NAME?) is written in capitals here, FOO(1); Excel keeps it as typed.
  (Telling it from a function Excel has but this editor hasn't needs Excel's whole list.)
- **The range operator and intersection in Evaluate Formula** are one step each, as a reference is (=SUM(A1:INDEX(A:A,3))
  goes straight to its sum); Excel's steps there aren't measured.
- **Iterative calculation turned on or changed** (*measured*: =A1+1 left out at 0; Iteration turned on with Maximum
  Iterations 50, A1 50; Maximum Iterations then 20, 70; off again, 70 left): on, every formula is worked out once
  more, the circles going round; off, nothing is (`setIteration`). Formulas > Calculation Options > Iterative
  Calculation... has Excel's File > Options > Formulas settings for it (`editor/iterbox.js`).
- **Iterative calculation: what isn't yet as Excel.** A circle whose cells change only
  in text or errors isn't measured: here only numbers count toward Maximum Change. Array and dynamic formulas in a
  circle aren't measured either.
- **Dates and times typed as entries** are read as dates and times (docs/ENTRY.md). Inside a formula, text such as
  "2024-01-02" or "1/2/2024" reads as a date.
