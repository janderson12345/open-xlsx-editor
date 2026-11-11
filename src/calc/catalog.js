// The functions as Excel's Insert Function box and the Formulas tab's Function Library list them: each one's category
// (Excel's, from Microsoft's "Excel functions (by category)"; one category each, where that list names a function in
// two: CONCATENATE in Text, FORECAST in Statistical, FLOOR in Math & Trig, LET in Logical) and its syntax with
// Excel's argument names, optional ones in brackets. Only the functions the engine works out (test/catalog.test.js
// checks that each is here). DATEDIF, which Excel leaves out of its lists, is left out too.

const LIST = {
  Financial: `ACCRINT(issue,first_interest,settlement,rate,par,frequency,[basis],[calc_method])
    ACCRINTM(issue,settlement,rate,par,[basis]) AMORDEGRC(cost,date_purchased,first_period,salvage,period,rate,[basis])
    AMORLINC(cost,date_purchased,first_period,salvage,period,rate,[basis]) COUPDAYBS(settlement,maturity,frequency,[basis])
    COUPDAYS(settlement,maturity,frequency,[basis]) COUPDAYSNC(settlement,maturity,frequency,[basis])
    COUPNCD(settlement,maturity,frequency,[basis]) COUPNUM(settlement,maturity,frequency,[basis])
    COUPPCD(settlement,maturity,frequency,[basis]) DISC(settlement,maturity,pr,redemption,[basis])
    DURATION(settlement,maturity,coupon,yld,frequency,[basis]) INTRATE(settlement,maturity,investment,redemption,[basis])
    MDURATION(settlement,maturity,coupon,yld,frequency,[basis]) PRICE(settlement,maturity,rate,yld,redemption,frequency,[basis])
    PRICEDISC(settlement,maturity,discount,redemption,[basis]) PRICEMAT(settlement,maturity,issue,rate,yld,[basis])
    RECEIVED(settlement,maturity,investment,discount,[basis]) YIELD(settlement,maturity,rate,pr,redemption,frequency,[basis])
    YIELDDISC(settlement,maturity,pr,redemption,[basis]) YIELDMAT(settlement,maturity,issue,rate,pr,[basis])
    CUMIPMT(rate,nper,pv,start_period,end_period,type) CUMPRINC(rate,nper,pv,start_period,end_period,type)
    DB(cost,salvage,life,period,[month]) DDB(cost,salvage,life,period,[factor]) DOLLARDE(fractional_dollar,fraction)
    DOLLARFR(decimal_dollar,fraction) EFFECT(nominal_rate,npery) FV(rate,nper,pmt,[pv],[type]) FVSCHEDULE(principal,schedule)
    IPMT(rate,per,nper,pv,[fv],[type]) IRR(values,[guess]) ISPMT(rate,per,nper,pv) MIRR(values,finance_rate,reinvest_rate)
    NOMINAL(effect_rate,npery) NPER(rate,pmt,pv,[fv],[type]) NPV(rate,value1,[value2],...) PDURATION(rate,pv,fv)
    PMT(rate,nper,pv,[fv],[type]) PPMT(rate,per,nper,pv,[fv],[type]) PV(rate,nper,pmt,[fv],[type])
    RATE(nper,pmt,pv,[fv],[type],[guess]) RRI(nper,pv,fv) SLN(cost,salvage,life) SYD(cost,salvage,life,per)
    TBILLEQ(settlement,maturity,discount) TBILLPRICE(settlement,maturity,discount) TBILLYIELD(settlement,maturity,pr)
    VDB(cost,salvage,life,start_period,end_period,[factor],[no_switch]) XIRR(values,dates,[guess]) XNPV(rate,values,dates)`,
  Logical: `AND(logical1,[logical2],...) BYCOL(array,lambda) BYROW(array,lambda) FALSE()
    IF(logical_test,[value_if_true],[value_if_false]) IFERROR(value,value_if_error) IFNA(value,value_if_na)
    IFS(logical_test1,value_if_true1,...) LAMBDA([parameter1,...],calculation) LET(name1,name_value1,calculation_or_name2,...)
    MAKEARRAY(rows,cols,lambda) MAP(array1,lambda_or_array2,...) NOT(logical) OR(logical1,[logical2],...)
    REDUCE([initial_value],array,lambda) SCAN([initial_value],array,lambda) SWITCH(expression,value1,result1,[default_or_value2,result2],...)
    TRUE() XOR(logical1,[logical2],...)`,
  Text: `ARRAYTOTEXT(array,[format]) CHAR(number) CLEAN(text) CODE(text) CONCAT(text1,[text2],...) CONCATENATE(text1,[text2],...)
    DOLLAR(number,[decimals]) EXACT(text1,text2) FIND(find_text,within_text,[start_num]) FIXED(number,[decimals],[no_commas])
    LEFT(text,[num_chars]) LEN(text) LOWER(text) MID(text,start_num,num_chars)
    NUMBERVALUE(text,[decimal_separator],[group_separator]) PROPER(text) REGEXEXTRACT(text,pattern,[return_mode],[case_sensitivity])
    REGEXREPLACE(text,pattern,replacement,[occurrence],[case_sensitivity]) REGEXTEST(text,pattern,[case_sensitivity])
    REPLACE(old_text,start_num,num_chars,new_text) REPT(text,number_times) RIGHT(text,[num_chars]) SEARCH(find_text,within_text,[start_num])
    SUBSTITUTE(text,old_text,new_text,[instance_num]) T(value) TEXT(value,format_text)
    TEXTAFTER(text,delimiter,[instance_num],[match_mode],[match_end],[if_not_found])
    TEXTBEFORE(text,delimiter,[instance_num],[match_mode],[match_end],[if_not_found])
    TEXTJOIN(delimiter,ignore_empty,text1,[text2],...) TEXTSPLIT(text,col_delimiter,[row_delimiter],[ignore_empty],[match_mode],[pad_with])
    TRIM(text) UNICHAR(number) UNICODE(text) UPPER(text) VALUE(text) VALUETOTEXT(value,[format])`,
  'Date & Time': `DATE(year,month,day) DATEVALUE(date_text) DAY(serial_number) DAYS(end_date,start_date)
    DAYS360(start_date,end_date,[method]) EDATE(start_date,months) EOMONTH(start_date,months) HOUR(serial_number)
    ISOWEEKNUM(date) MINUTE(serial_number) MONTH(serial_number) NETWORKDAYS(start_date,end_date,[holidays])
    NETWORKDAYS.INTL(start_date,end_date,[weekend],[holidays]) NOW() SECOND(serial_number) TIME(hour,minute,second)
    TIMEVALUE(time_text) TODAY() WEEKDAY(serial_number,[return_type]) WEEKNUM(serial_number,[return_type])
    WORKDAY(start_date,days,[holidays]) WORKDAY.INTL(start_date,days,[weekend],[holidays]) YEAR(serial_number)
    YEARFRAC(start_date,end_date,[basis])`,
  'Lookup & Reference': `ADDRESS(row_num,column_num,[abs_num],[a1],[sheet_text]) AREAS(reference)
    CHOOSE(index_num,value1,[value2],...) CHOOSECOLS(array,col_num1,[col_num2],...) CHOOSEROWS(array,row_num1,[row_num2],...)
    COLUMN([reference]) COLUMNS(array) DROP(array,rows,[columns]) EXPAND(array,rows,[columns],[pad_with])
    FILTER(array,include,[if_empty]) FORMULATEXT(reference) HLOOKUP(lookup_value,table_array,row_index_num,[range_lookup])
    HSTACK(array1,[array2],...) HYPERLINK(link_location,[friendly_name]) INDEX(array,row_num,[column_num])
    INDIRECT(ref_text,[a1]) LOOKUP(lookup_value,lookup_vector,[result_vector]) MATCH(lookup_value,lookup_array,[match_type])
    OFFSET(reference,rows,cols,[height],[width]) ROW([reference]) ROWS(array) SORT(array,[sort_index],[sort_order],[by_col])
    SORTBY(array,by_array1,[sort_order1],...) TAKE(array,rows,[columns]) TOCOL(array,[ignore],[scan_by_column])
    TOROW(array,[ignore],[scan_by_column]) TRANSPOSE(array) UNIQUE(array,[by_col],[exactly_once])
    VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup]) VSTACK(array1,[array2],...)
    WRAPCOLS(vector,wrap_count,[pad_with]) WRAPROWS(vector,wrap_count,[pad_with])
    XLOOKUP(lookup_value,lookup_array,return_array,[if_not_found],[match_mode],[search_mode])
    XMATCH(lookup_value,lookup_array,[match_mode],[search_mode])`,
  'Math & Trig': `ABS(number) ACOS(number) ACOSH(number) ACOT(number) ACOTH(number) AGGREGATE(function_num,options,ref1,...)
    ARABIC(text) ASIN(number) ASINH(number) ATAN(number) ATAN2(x_num,y_num) ATANH(number) BASE(number,radix,[min_length])
    CEILING(number,significance) CEILING.MATH(number,[significance],[mode]) CEILING.PRECISE(number,[significance]) COMBIN(number,number_chosen)
    COMBINA(number,number_chosen) COS(number) COSH(number) COT(number) COTH(number) CSC(number) CSCH(number)
    DECIMAL(text,radix) DEGREES(angle) EVEN(number) EXP(number) FACT(number) FACTDOUBLE(number) FLOOR(number,significance)
    FLOOR.MATH(number,[significance],[mode]) FLOOR.PRECISE(number,[significance]) ISO.CEILING(number,[significance]) GCD(number1,[number2],...) INT(number) LCM(number1,[number2],...) LN(number)
    LOG(number,[base]) LOG10(number) MDETERM(array) MINVERSE(array) MMULT(array1,array2) MOD(number,divisor)
    MROUND(number,multiple) MULTINOMIAL(number1,[number2],...) MUNIT(dimension) ODD(number) PI() POWER(number,power)
    PRODUCT(number1,[number2],...) QUOTIENT(numerator,denominator) RADIANS(angle) RAND() RANDBETWEEN(bottom,top)
    ROMAN(number,[form]) ROUND(number,num_digits) ROUNDDOWN(number,num_digits) ROUNDUP(number,num_digits) SEC(number)
    SECH(number) SEQUENCE(rows,[columns],[start],[step]) SERIESSUM(x,n,m,coefficients) SIGN(number) SIN(number)
    SINH(number) SQRT(number) SQRTPI(number) SUBTOTAL(function_num,ref1,[ref2],...) SUM(number1,[number2],...)
    SUMIF(range,criteria,[sum_range]) SUMIFS(sum_range,criteria_range1,criteria1,...) SUMPRODUCT(array1,[array2],[array3],...)
    SUMSQ(number1,[number2],...) SUMX2MY2(array_x,array_y) SUMX2PY2(array_x,array_y) SUMXMY2(array_x,array_y) TAN(number)
    TANH(number) TRUNC(number,[num_digits])`,
  Statistical: `AVEDEV(number1,[number2],...) AVERAGE(number1,[number2],...) AVERAGEA(value1,[value2],...)
    AVERAGEIF(range,criteria,[average_range]) AVERAGEIFS(average_range,criteria_range1,criteria1,...)
    BETA.DIST(x,alpha,beta,cumulative,[A],[B]) BETA.INV(probability,alpha,beta,[A],[B])
    BINOM.DIST(number_s,trials,probability_s,cumulative) BINOM.DIST.RANGE(trials,probability_s,number_s,[number_s2])
    BINOM.INV(trials,probability_s,alpha)
    CHISQ.DIST(x,deg_freedom,cumulative) CHISQ.DIST.RT(x,deg_freedom) CHISQ.INV(probability,deg_freedom)
    CHISQ.INV.RT(probability,deg_freedom) CHISQ.TEST(actual_range,expected_range) CONFIDENCE.NORM(alpha,standard_dev,size)
    CONFIDENCE.T(alpha,standard_dev,size) CORREL(array1,array2) COUNT(value1,[value2],...) COUNTA(value1,[value2],...)
    COUNTBLANK(range) COUNTIF(range,criteria) COUNTIFS(criteria_range1,criteria1,...) COVARIANCE.P(array1,array2)
    COVARIANCE.S(array1,array2) DEVSQ(number1,[number2],...) EXPON.DIST(x,lambda,cumulative)
    F.DIST(x,deg_freedom1,deg_freedom2,cumulative) F.DIST.RT(x,deg_freedom1,deg_freedom2)
    F.INV(probability,deg_freedom1,deg_freedom2) F.INV.RT(probability,deg_freedom1,deg_freedom2) F.TEST(array1,array2)
    FISHER(x) FISHERINV(y) FORECAST(x,known_y's,known_x's) FORECAST.LINEAR(x,known_y's,known_x's)
    FREQUENCY(data_array,bins_array) GAMMA(x) GAMMA.DIST(x,alpha,beta,cumulative) GAMMA.INV(probability,alpha,beta)
    GAMMALN(x) GAMMALN.PRECISE(x) GAUSS(x) GEOMEAN(number1,[number2],...) GROWTH(known_y's,[known_x's],[new_x's],[const])
    HARMEAN(number1,[number2],...) HYPGEOM.DIST(sample_s,number_sample,population_s,number_pop,cumulative)
    INTERCEPT(known_y's,known_x's) KURT(number1,[number2],...) LARGE(array,k) LINEST(known_y's,[known_x's],[const],[stats])
    LOGEST(known_y's,[known_x's],[const],[stats]) LOGNORM.DIST(x,mean,standard_dev,cumulative)
    LOGNORM.INV(probability,mean,standard_dev) MAX(number1,[number2],...) MAXA(value1,[value2],...)
    MAXIFS(max_range,criteria_range1,criteria1,...) MEDIAN(number1,[number2],...) MIN(number1,[number2],...)
    MINA(value1,[value2],...) MINIFS(min_range,criteria_range1,criteria1,...) MODE.MULT(number1,[number2],...)
    MODE.SNGL(number1,[number2],...) NEGBINOM.DIST(number_f,number_s,probability_s,cumulative)
    NORM.DIST(x,mean,standard_dev,cumulative) NORM.INV(probability,mean,standard_dev) NORM.S.DIST(z,cumulative)
    NORM.S.INV(probability) PEARSON(array1,array2) PERCENTILE.EXC(array,k) PERCENTILE.INC(array,k)
    PERCENTRANK.EXC(array,x,[significance]) PERCENTRANK.INC(array,x,[significance]) PERMUT(number,number_chosen)
    PERMUTATIONA(number,number_chosen) PHI(x) POISSON.DIST(x,mean,cumulative)
    PROB(x_range,prob_range,[lower_limit],[upper_limit]) QUARTILE.EXC(array,quart) QUARTILE.INC(array,quart)
    RANK.AVG(number,ref,[order]) RANK.EQ(number,ref,[order]) RSQ(known_y's,known_x's) SKEW(number1,[number2],...) SKEW.P(number1,[number2],...)
    SLOPE(known_y's,known_x's) SMALL(array,k) STANDARDIZE(x,mean,standard_dev) STDEV.P(number1,[number2],...)
    STDEV.S(number1,[number2],...) STDEVA(value1,[value2],...) STDEVPA(value1,[value2],...) STEYX(known_y's,known_x's)
    T.DIST(x,deg_freedom,cumulative) T.DIST.2T(x,deg_freedom) T.DIST.RT(x,deg_freedom) T.INV(probability,deg_freedom)
    T.INV.2T(probability,deg_freedom) T.TEST(array1,array2,tails,type) TREND(known_y's,[known_x's],[new_x's],[const])
    TRIMMEAN(array,percent) VAR.P(number1,[number2],...) VAR.S(number1,[number2],...) VARA(value1,[value2],...)
    VARPA(value1,[value2],...) WEIBULL.DIST(x,alpha,beta,cumulative) Z.TEST(array,x,[sigma])`,
  Database: `DAVERAGE(database,field,criteria) DCOUNT(database,[field],criteria) DCOUNTA(database,[field],criteria)
    DGET(database,field,criteria) DMAX(database,field,criteria) DMIN(database,field,criteria) DPRODUCT(database,field,criteria)
    DSTDEV(database,field,criteria) DSTDEVP(database,field,criteria) DSUM(database,field,criteria) DVAR(database,field,criteria)
    DVARP(database,field,criteria)`,
  Engineering: `BESSELI(x,n) BESSELJ(x,n) BESSELK(x,n) BESSELY(x,n) BIN2DEC(number) BIN2HEX(number,[places])
    BIN2OCT(number,[places]) BITAND(number1,number2)
    BITLSHIFT(number,shift_amount) BITOR(number1,number2) BITRSHIFT(number,shift_amount) BITXOR(number1,number2)
    DEC2BIN(number,[places]) DEC2HEX(number,[places]) DEC2OCT(number,[places]) DELTA(number1,[number2])
    COMPLEX(real_num,i_num,[suffix]) CONVERT(number,from_unit,to_unit) ERF(lower_limit,[upper_limit]) ERF.PRECISE(x) ERFC(x) ERFC.PRECISE(x)
    GESTEP(number,[step]) HEX2BIN(number,[places]) HEX2DEC(number) HEX2OCT(number,[places]) IMABS(inumber)
    IMAGINARY(inumber) IMARGUMENT(inumber) IMCONJUGATE(inumber) IMCOS(inumber) IMCOSH(inumber) IMCOT(inumber)
    IMCSC(inumber) IMCSCH(inumber) IMDIV(inumber1,inumber2) IMEXP(inumber) IMLN(inumber) IMLOG10(inumber)
    IMLOG2(inumber) IMPOWER(inumber,number) IMPRODUCT(inumber1,[inumber2],...) IMREAL(inumber) IMSEC(inumber)
    IMSECH(inumber) IMSIN(inumber) IMSINH(inumber) IMSQRT(inumber) IMSUB(inumber1,inumber2)
    IMSUM(inumber1,[inumber2],...) IMTAN(inumber) OCT2BIN(number,[places]) OCT2DEC(number) OCT2HEX(number,[places])`,
  Information: `ERROR.TYPE(error_val) ISBLANK(value) ISERR(value) ISERROR(value) ISEVEN(number) ISFORMULA(reference)
    ISLOGICAL(value) ISNA(value) ISNONTEXT(value) ISNUMBER(value) ISODD(number) ISOMITTED(argument) ISREF(value)
    ISTEXT(value) N(value) NA() SHEET([value]) SHEETS([reference]) TYPE(value)`,
  Compatibility: `BETADIST(x,alpha,beta,[A],[B]) BETAINV(probability,alpha,beta,[A],[B])
    BINOMDIST(number_s,trials,probability_s,cumulative) CHIDIST(x,deg_freedom) CHIINV(probability,deg_freedom)
    CHITEST(actual_range,expected_range) CONFIDENCE(alpha,standard_dev,size) COVAR(array1,array2) CRITBINOM(trials,probability_s,alpha)
    EXPONDIST(x,lambda,cumulative) FDIST(x,deg_freedom1,deg_freedom2) FINV(probability,deg_freedom1,deg_freedom2)
    FTEST(array1,array2) GAMMADIST(x,alpha,beta,cumulative) GAMMAINV(probability,alpha,beta)
    HYPGEOMDIST(sample_s,number_sample,population_s,number_pop) LOGINV(probability,mean,standard_dev)
    LOGNORMDIST(x,mean,standard_dev) MODE(number1,[number2],...) NEGBINOMDIST(number_f,number_s,probability_s)
    NORMDIST(x,mean,standard_dev,cumulative) NORMINV(probability,mean,standard_dev) NORMSDIST(z) NORMSINV(probability)
    PERCENTILE(array,k) PERCENTRANK(array,x,[significance]) POISSON(x,mean,cumulative) QUARTILE(array,quart)
    RANK(number,ref,[order]) STDEV(number1,[number2],...) STDEVP(number1,[number2],...) TDIST(x,deg_freedom,tails)
    TINV(probability,deg_freedom) TTEST(array1,array2,tails,type) VAR(number1,[number2],...) VARP(number1,[number2],...)
    WEIBULL(x,alpha,beta,cumulative) ZTEST(array,x,[sigma])`,
};

/** The categories in the order of Excel's Insert Function box. */
export const CATEGORIES = ['Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text', 'Logical', 'Information', 'Engineering', 'Compatibility'];

/** Excel's Most Recently Used list as it stands before anything is picked. */
export const DEFAULT_RECENT = ['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV'];

const BY_NAME = new Map();
for (const [category, text] of Object.entries(LIST)) {
  for (const sig of text.trim().split(/\s+/)) {
    const open = sig.indexOf('(');
    const name = sig.slice(0, open);
    const args = sig.slice(open + 1, -1);
    BY_NAME.set(name, { name, category, args: args ? args.split(',') : [] });
  }
}

/** Every function listed: { name, category, args }, by name. */
export const catalog = () => [...BY_NAME.values()].sort((a, b) => (a.name < b.name ? -1 : 1));
/** One category's functions, by name. */
export const inCategory = (category) => catalog().filter((f) => f.category === category);
/** A function's entry, or undefined. */
export const functionInfo = (name) => BY_NAME.get(String(name).toUpperCase());
/** A function's syntax as Excel shows it: "SUM(number1, [number2], ...)". */
export const syntaxOf = (name) => { const f = functionInfo(name); return f ? `${f.name}(${f.args.join(', ')})` : ''; };
/** The functions whose names hold `words` (Insert Function's search): names starting with them first. */
export function searchFunctions(words) {
  const q = String(words).trim().toUpperCase();
  if (!q) return [];
  const all = catalog();
  return [...all.filter((f) => f.name.startsWith(q)), ...all.filter((f) => !f.name.startsWith(q) && f.name.includes(q))];
}
