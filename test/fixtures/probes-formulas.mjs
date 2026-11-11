// The formula probe: formulas without saved values, for Excel to work out (see make-fixtures.mjs).
import { normalizeFormula } from '../../src/calc/tokens.js';
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc, col } from './fixture-common.mjs';

// A probe for the calculation engine: formulas without saved values, for Excel to work out when it opens the file, on
// a sheet of mixed data (numbers, text, a number as text, TRUE and FALSE, blanks, an error), with defined names, a
// second sheet whose name needs quotes, a shared formula and a chain read in both directions. Column A shows each
// formula as text, column B holds it.
export const PROBE_FORMULAS = [
  // (These first rows sit beside Data rows 2 to 9: a range in a single-value place is cut to the formula's row.)
  'Data!B2:B9', 'Data!B2:B9*1', 'Data!A2:A9&""', 'SUM(Data!B2:B9*1)', 'Data!A1:E1', '-Data!C2:C9', 'ROUND(Data!C2:C9,0)', 'Data!B2:B9',
  // Arithmetic, precedence, coercion
  '1+2*3', '(1+2)*3', '-2^2', '2^3^2', '2^-1', '10/4', '7/0', '0^0', '(-8)^(1/3)', '2^1024', '5%', '50%*4', '-50%', '2*-3',
  '"3"+1', '"1,234"+1', '" 4 "+1', '"$5"+0', '"(5)"+0', '"50%"+0', '"1e3"+0', '"abc"+1', 'TRUE+1', 'FALSE*5', '"2024-01-02"+0',
  '"1/2/2024"+0', '"10:30"+0', '"2:15 PM"+0', '""&1', '1&2', '"a"&TRUE', 'Data!F2+1', 'Data!F2&"x"',
  // Numbers turned into text
  '1/3&""', '2/3&""', '-1/3&""', '0.1+0.2&""', '1E-5&""', '1E-9&""', '1E-10&""', '1.5E-7&""', '1E-15&""', '1E-19&""', '1E-20&""',
  '123456789012345&""', '1234567890123456&""', '1E15&""', '1E20&""', '1E21&""', '1E22&""', '12345678901234567890&""', '100000*3&""',
  '0.000123456789012345678&""', 'LEN(1/7)',
  // Floating point
  '0.1+0.2-0.3', '(0.1+0.2-0.3)', '1*(0.1+0.2-0.3)', '0.1+0.2=0.3', '1/3*3=1', '0.3-0.1-0.2', '1-0.9-0.1', '3*0.1=0.3',
  '1E15+0.3-1E15', '0.1*3', '1.1-1', '(1.1-1)*10',
  // Comparisons
  '1<2', '"a"="A"', '"a"<"b"', '"B">"a"', '1<"a"', '"a"<TRUE', 'FALSE<TRUE', '1=TRUE', 'Data!F2=0', 'Data!F2=""', 'Data!F2=FALSE',
  '"10"<"9"', '1<>1', '"abc"<"abd"', '"a"<"a "', '"Z"<"a"', '"-"<"a"',
  // References and names
  'Data!B2', 'Data!B2+Data!B3', "'My Data'!A1*2", "'My Data'!A2", 'Rate*10', 'SUM(Items)', 'COUNTA(Items)', 'Data!$C$3', 'D11',
  // Aggregates
  'SUM(Data!B2:B9)', 'SUM(Data!B:B)', 'AVERAGE(Data!B2:B9)', 'MIN(Data!B2:B9)', 'MAX(Data!C2:C9)', 'COUNT(Data!A1:E9)',
  'COUNTA(Data!A1:E9)', 'COUNTBLANK(Data!A1:E9)', 'COUNTBLANK(Data!F:F)', 'PRODUCT(Data!B2:B4)', 'SUM(1,"2",TRUE)', 'SUM(Data!B2:B9,10)',
  'COUNT(1,"2",TRUE,"x")', 'COUNT(1,)', 'COUNTA(1,)', 'AVERAGE(1,)', 'SUM(Data!G2:G3)', 'MEDIAN(Data!B2:B9)', 'MEDIAN(1,2,3,4)',
  'LARGE(Data!B2:B9,2)', 'SMALL(Data!C2:C9,3)', 'STDEV(Data!B2:B9)', 'STDEV.P(Data!B2:B9)', 'VAR(Data!C2:C9)', 'VAR.P(Data!C2:C9)',
  'SUMSQ(1,2,3)', 'SUMPRODUCT(Data!B2:B9,Data!C2:C9)', 'SUMPRODUCT((Data!B2:B9>3)*Data!C2:C9)', 'SUMPRODUCT((Data!A2:A9="Apple")*1)',
  'SUMPRODUCT(--ISNUMBER(SEARCH("an",Data!A2:A9)))', 'SUMPRODUCT({1,2,3},{4,5,6})', 'SUM({1,2,3}*2)', 'MAX(Data!E2:E9)',
  'SUM(Data!E2:E9)', 'AVERAGE(Data!A2:A9)',
  // Math
  'ROUND(2.675,2)', 'ROUND(-2.5,0)', 'ROUND(1234.5678,-2)', 'ROUND(0.1+0.2,15)', 'ROUNDUP(3.14159,2)', 'ROUNDUP(-3.14159,2)',
  'ROUNDDOWN(-3.9,0)', 'INT(-3.5)', 'TRUNC(-3.5)', 'TRUNC(3.14159,3)', 'MOD(10,3)', 'MOD(-10,3)', 'MOD(10,-3)', 'MOD(5.5,2)', 'MOD(1,0)',
  'MOD(0.3,0.1)', 'ABS(-4)', 'SQRT(16)', 'SQRT(-1)', 'POWER(2,10)', 'EXP(1)', 'LN(10)', 'LOG(100)', 'LOG(8,2)', 'LOG10(1000)', 'LN(0)',
  'PI()', 'SIGN(-3)', 'CEILING(2.5,1)', 'CEILING(-2.5,-2)', 'CEILING(2.5,-1)', 'CEILING(-2.5,2)', 'FLOOR(2.5,1)', 'FLOOR(-2.5,-2)',
  'FLOOR(-2.5,2)', 'MROUND(10,3)', 'MROUND(-10,3)', 'EVEN(1.5)', 'ODD(1.5)', 'EVEN(-1.5)', 'ODD(0)', 'FACT(5)', 'FACT(0)',
  'QUOTIENT(7,2)', 'QUOTIENT(-7,2)', 'INT(0.1+0.7)*10',
  // Logic
  'IF(1>2,"y","n")', 'IF(Data!F2,"t","f")', 'IF(1,"a")', 'IF(0,"a")', 'IF(TRUE,)', 'IF("x",1,2)', 'IFERROR(1/0,"err")',
  'IFERROR(5,"err")', 'IFNA(NA(),0)', 'IFNA(1/0,0)', 'AND(TRUE,1)', 'OR(FALSE,0)', 'NOT(1)', 'XOR(TRUE,TRUE,TRUE)', 'AND(Data!A1:A3)',
  'AND(Data!E2:E3)', 'OR(Data!E2:E9)', 'IFS(1>2,"a",TRUE,"b")', 'IFS(FALSE,1)', 'SWITCH(2,1,"one",2,"two")', 'SWITCH(3,1,"one","none")',
  // Information
  'ISBLANK(Data!F2)', 'ISBLANK("")', 'ISNUMBER(Data!B2)', 'ISNUMBER(Data!B4)', 'ISNUMBER("1")', 'ISTEXT(Data!A2)', 'ISERROR(1/0)',
  'ISERR(NA())', 'ISNA(NA())', 'ISLOGICAL(TRUE)', 'ISEVEN(4)', 'ISODD(3.7)', 'N(TRUE)', 'N("x")', 'ISNONTEXT(Data!F2)', 'ISREF(Data!A1)',
  'ISERROR(Data!G2)', 'T(1)', 'T("a")',
  // Text
  'CONCATENATE("a",1,TRUE)', 'CONCAT(Data!A2:A4)', 'CONCAT(Data!B2:C3)', 'TEXTJOIN("-",TRUE,Data!A2:A5,"x")',
  'TEXTJOIN(",",FALSE,Data!E4:E6)', 'LEN("hello")', 'LEFT("hello",2)', 'LEFT("hello")', 'RIGHT("hello",3)', 'MID("hello",2,3)',
  'MID("hello",0,1)', 'MID("hello",9,2)', 'UPPER("abc")', 'LOWER("ABC")', 'PROPER("hello wORLD o\'neil 2nd")', 'TRIM("  a   b  ")',
  'REPT("ab",3)', 'SUBSTITUTE("aaa","a","b")', 'SUBSTITUTE("aaa","a","b",2)', 'REPLACE("abcdef",2,3,"X")', 'FIND("l","hello")',
  'FIND("L","hello")', 'SEARCH("L","hello")', 'SEARCH("h?l","ahelp")', 'SEARCH("~?","a?b")', 'EXACT("a","A")', 'VALUE("1,234.5")',
  'VALUE("abc")', 'VALUE("12%")', 'TEXT(1234.567,"#,##0.00")', 'TEXT(0.25,"0%")', 'TEXT(45678,"yyyy-mm-dd")', 'TEXT("abc","0")',
  'TEXT(-1.5,"0.0;(0.0)")', 'TEXT(TRUE,"0")', 'CHAR(65)', 'CODE("A")', '"x"&Data!F2&"y"', 'LEN(Data!B4)', 'UPPER(Data!B2)',
  // Lookup
  'VLOOKUP("Banana",Data!A2:C9,3,FALSE)', 'VLOOKUP("banana",Data!A2:C9,2,FALSE)', 'VLOOKUP("B*",Data!A2:C9,2,FALSE)',
  'VLOOKUP("Zzz",Data!A2:C9,2,FALSE)', 'VLOOKUP(65,Data!I2:J6,2)', 'VLOOKUP(80,Data!I2:J6,2,TRUE)', 'VLOOKUP(-1,Data!I2:J6,2)',
  'VLOOKUP(65,Data!I2:J6,3)', 'VLOOKUP(65,Data!I2:J6,0)', 'VLOOKUP("Fig",Data!A2:E9,5,0)', 'VLOOKUP("Fig",Data!A2:F9,6,0)',
  'HLOOKUP("Price",Data!A1:E9,3,FALSE)', 'MATCH("Cherry",Data!A2:A9,0)', 'MATCH(65,Data!I2:I6,1)', 'MATCH(65,Data!I2:I6)',
  'MATCH("c*",Data!A2:A9,0)', 'MATCH(99,Data!I2:I6,0)', 'INDEX(Data!A2:C9,3,2)', 'INDEX(Data!A2:A9,4)', 'INDEX(Data!A1:E1,3)',
  'INDEX(Data!A2:C9,9,1)', 'SUM(INDEX(Data!B2:C9,0,1))', 'SUM(INDEX(Data!B2:C9,2,0))', 'INDEX({1,2;3,4},2,1)',
  'XLOOKUP("Cherry",Data!A2:A9,Data!C2:C9)', 'XLOOKUP("nope",Data!A2:A9,Data!C2:C9,"none")', 'XLOOKUP(65,Data!I2:I6,Data!J2:J6,,-1)',
  'XLOOKUP(65,Data!I2:I6,Data!J2:J6,,1)', 'XLOOKUP("Apple",Data!A2:A9,Data!C2:C9,,0,-1)', 'CHOOSE(2,"a","b","c")', 'CHOOSE(4,"a")',
  'ROW()', 'COLUMN()', 'ROW(Data!C5)', 'COLUMN(Data!C5)', 'ROWS(Data!A2:C9)', 'COLUMNS(Data!A2:C9)', 'SUM(OFFSET(Data!B2,1,0,3,1))',
  'OFFSET(Data!A1,2,2)', 'INDIRECT("Data!B3")', 'SUM(INDIRECT("Data!C2:C4"))', 'INDIRECT("Nowhere!A1")',
  // Conditional
  'SUMIF(Data!B2:B9,">3")', 'SUMIF(Data!A2:A9,"Apple",Data!C2:C9)', 'SUMIF(Data!A2:A9,"A*",Data!B2:B9)', 'SUMIF(Data!A2:A9,"apple",Data!C2)',
  'COUNTIF(Data!B2:B9,">=4")', 'COUNTIF(Data!A2:A9,"apple")', 'COUNTIF(Data!A2:A9,"<>Apple")', 'COUNTIF(Data!B2:B9,12)',
  'COUNTIF(Data!B2:B9,"12")', 'COUNTIF(Data!B2:B9,"=12")', 'COUNTIF(Data!F1:F9,"")', 'COUNTIF(Data!F1:F9,"=")', 'COUNTIF(Data!A1:E9,"<>")',
  'COUNTIF(Data!A2:A9,"*")', 'COUNTIF(Data!E2:E9,TRUE)', 'COUNTIF(Data!E2:E9,"TRUE")', 'COUNTIF(Data!G2:G3,"#N/A")',
  'COUNTIF(Data!A2:A9,"<C")', 'COUNTIF(Data!F:F,"")', 'COUNTIF(Data!B2:B9,"<>3")', 'AVERAGEIF(Data!B2:B9,">2")',
  'AVERAGEIF(Data!B2:B9,">100")', 'SUMIFS(Data!C2:C9,Data!A2:A9,"Apple",Data!B2:B9,">2")', 'COUNTIFS(Data!A2:A9,"Apple",Data!E2:E9,FALSE)',
  'AVERAGEIFS(Data!C2:C9,Data!A2:A9,"<>Apple")', 'MAXIFS(Data!C2:C9,Data!A2:A9,"Apple")', 'MINIFS(Data!C2:C9,Data!A2:A9,"Apple")',
  // Dates and times
  'DATE(2024,1,31)', 'DATE(2024,14,1)', 'DATE(2024,1,0)', 'DATE(99,1,1)', 'DATE(1900,3,1)', 'DATE(1900,1,1)', 'YEAR(45678)', 'MONTH(45678)',
  'DAY(45678)', 'WEEKDAY(45678)', 'WEEKDAY(45678,2)', 'WEEKDAY(45678,3)', 'TIME(12,30,0)', 'TIME(25,0,0)', 'HOUR(0.75)', 'MINUTE(0.7604)',
  'SECOND(0.5000115)', 'EDATE(DATE(2024,1,31),1)', 'EOMONTH(DATE(2024,1,15),1)', 'EOMONTH(DATE(2024,1,15),-1)',
  'DAYS(DATE(2024,3,1),DATE(2024,1,1))', 'DATE(2024,2,29)-DATE(2023,2,28)', 'YEAR("2024-03-05")', 'YEAR(-1)', 'DATEVALUE("2024-03-05")',
  'MONTH(Data!D5)',
  // Errors
  '#N/A+1', '#DIV/0!&"x"', 'Data!G2', 'Data!G2+1', '#REF!', 'SUM(1,#VALUE!)', 'IF(#N/A,1,2)', 'ISERROR(#NULL!)',
  // More numbers as text, and the last + or - of a formula near zero
  '1E-16&""', '1E-17&""', '1E-18&""', '1.2345E-14&""', '1.23456E-14&""', '1.5E-18&""', '1.23456789012345E-5&""', '99999999999999999999&""',
  '1.5E+19&""', '-1E-15&""', '0.5-0.4-0.1', '0.1+0.2-0.3+0', '1+1E-15-1', '1-1E-15-1', '2.3-2.2-0.1', 'Data!C2-1.5', '-(0.1+0.2-0.3)',
  '0.1+0.2-0.3>0', '10*(0.1+0.2)-3', '(0.1+0.2)*10-3', '100-99.99', '1E-14+1-1', 'SUM(0.1,0.2,-0.3)', '{1,2,3}*2', 'ROUND({1.2,2.7},0)',
  // SUBTOTAL, ranks, percentiles, quartiles, the mode (functions-more.js)
  'SUBTOTAL(9,Data!C2:C9)', 'SUBTOTAL(1,Data!C2:C9)', 'SUBTOTAL(2,Data!B2:B9)', 'SUBTOTAL(3,Data!A2:A9)', 'SUBTOTAL(4,Data!C2:C9)',
  'SUBTOTAL(109,Data!C2:C9)', 'SUBTOTAL(12,Data!C2:C9)', 'RANK(4,Data!C2:C9)', 'RANK(4,Data!C2:C9,1)', 'RANK.EQ(2,Data!B2:B9)',
  'RANK.AVG(2,Data!B2:B9)', 'RANK(9,Data!C2:C9)', 'PERCENTILE(Data!C2:C9,0.3)', 'PERCENTILE.INC(Data!C2:C9,1)', 'PERCENTILE.EXC(Data!C2:C9,0.25)',
  'PERCENTILE.EXC(Data!C2:C9,0.05)', 'QUARTILE(Data!C2:C9,1)', 'QUARTILE.INC(Data!C2:C9,3)', 'QUARTILE.EXC(Data!C2:C9,2)', 'QUARTILE(Data!C2:C9,5)',
  'MODE(1,2,2,3,3)', 'MODE.SNGL(Data!B2:B9)', 'MODE(1,2,3)', 'AVERAGEA(Data!E2:E9)', 'MAXA(Data!E2:E9)', 'MINA(Data!A2:A3,-1)',
  // Working days, DATEDIF, weeks, YEARFRAC
  'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', 'NETWORKDAYS(DATE(2024,1,31),DATE(2024,1,1))', 'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45293,45296})',
  'WORKDAY(DATE(2024,1,5),1)', 'WORKDAY(DATE(2024,1,8),-3)', 'WORKDAY(DATE(2024,1,1),10,DATE(2024,1,2))',
  'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"Y")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"M")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"D")',
  'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"YM")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"MD")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"YD")',
  'DATEDIF(DATE(2024,3,1),DATE(2024,1,1),"D")', 'WEEKNUM(DATE(2024,1,1))', 'WEEKNUM(DATE(2024,1,7))', 'WEEKNUM(DATE(2024,1,7),2)',
  'WEEKNUM(DATE(2024,12,31),21)', 'ISOWEEKNUM(DATE(2021,1,1))', 'ISOWEEKNUM(DATE(2024,12,30))', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))',
  'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),1)', 'YEARFRAC(DATE(2023,1,31),DATE(2024,2,29),0)', 'YEARFRAC(DATE(2020,3,15),DATE(2024,9,1),1)',
  'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)', 'YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),4)',
  // Loans and investments
  'PMT(0.05/12,360,200000)', 'PMT(0,12,1200)', 'PMT(0.06/12,60,-20000,0,1)', 'PV(0.08/12,240,-500)', 'FV(0.06/12,120,-100)', 'FV(0.05,10,-100,-1000,1)',
  'NPER(0.01,-100,1000)', 'NPER(0,-100,1000)', 'RATE(48,-200,8000)', 'RATE(10,-100,1000)', 'NPV(0.1,-10000,3000,4200,6800)',
  'IRR({-70000,12000,15000,18000,21000,26000})', 'IRR({-100,10})',
  // LOOKUP, XMATCH, ADDRESS
  'LOOKUP(4.5,{1,2,3,4,5},{"a","b","c","d","e"})', 'LOOKUP("Cherry",Data!A2:A9)', 'LOOKUP(0,{1,2})', 'XMATCH(4,Data!C2:C9)', 'XMATCH(5,{1,3,5,7},1)',
  'XMATCH(4,{1,3,5,7},-1)', 'XMATCH(4,{1,3,5,7},1)', 'XMATCH("App*",Data!A2:A9,2)', 'XMATCH(2,{1,2,2,3},0,-1)', 'ADDRESS(2,3)', 'ADDRESS(2,3,4)',
  'ADDRESS(2,3,2,FALSE)', 'ADDRESS(1,28,1,TRUE,"My Data")',
  // Text
  'TEXTBEFORE("a-b-c","-")', 'TEXTBEFORE("a-b-c","-",2)', 'TEXTAFTER("a-b-c","-",-1)', 'TEXTAFTER("a-b-c","x")', 'TEXTAFTER("a-b-c","x",1,0,0,"none")',
  'TEXTBEFORE("A-B","b",1,1)', 'NUMBERVALUE("1.234,5",",",".")', 'NUMBERVALUE("12%")', 'NUMBERVALUE(" 3 000 ")', 'FIXED(1234.567)', 'FIXED(1234.567,1,TRUE)',
  'FIXED(-1234.5,-2)', 'DOLLAR(-1234.567)', 'DOLLAR(0.5,0)', 'UNICODE("A")', 'UNICHAR(9731)',
  // Trigonometry, GCD, LCM, rounding to a multiple, COMBIN, TYPE, ERROR.TYPE
  'SIN(PI()/6)', 'COS(0)', 'TAN(1)', 'ASIN(2)', 'ATAN2(1,1)', 'DEGREES(PI())', 'RADIANS(180)', 'GCD(24,36)', 'GCD(7,0)', 'LCM(4,6)', 'LCM(4,0)',
  'CEILING.MATH(4.3)', 'CEILING.MATH(-4.3)', 'CEILING.MATH(-4.3,1,1)', 'FLOOR.MATH(4.7,2)', 'FLOOR.MATH(-4.7)', 'CEILING.MATH(6.7,0.5)', 'COMBIN(8,3)',
  'COMBIN(3,5)', 'TYPE(1)', 'TYPE("a")', 'TYPE(TRUE)', 'TYPE(#N/A)', 'TYPE({1,2})', 'ERROR.TYPE(#DIV/0!)', 'ERROR.TYPE(1)',
  // Table references (the table Fruit on Data!A1:E9: ooxml/tables.js)
  'SUM(Fruit[Qty])', 'SUM(Fruit[Price])', 'COUNTA(Fruit[Name])', 'ROWS(Fruit[#Data])', 'ROWS(Fruit[#All])', 'INDEX(Fruit[#Headers],1,3)',
  'SUM(Fruit[[Qty]:[Price]])', 'SUMIFS(Fruit[Price],Fruit[Name],"Apple")', 'XLOOKUP("Fig",Fruit[Name],Fruit[Price])', 'SUM(fruit[qty])',
  'Fruit[[#Headers],[Price]]', 'COLUMNS(Fruit[#All])', 'SUBTOTAL(109,Fruit[Price])', 'Fruit[[#This Row],[Name]]',
  'AVERAGE(Fruit[[#Data],[Price]])', 'COUNTA(Fruit[[#Headers],[#Data],[Name]])', 'MATCH("Cherry",Fruit[Name],0)', 'SUM(Fruit[[#All],[Qty]:[Price]])',
  // Excel 365's array functions (each array joined into one text, its shape by ROWS and COLUMNS)
  'TEXTJOIN(",",FALSE,TAKE({1,2,3;4,5,6;7,8,9},2))', 'TEXTJOIN(",",FALSE,TAKE({1,2,3;4,5,6;7,8,9},-1))', 'TEXTJOIN(",",FALSE,TAKE({1,2,3;4,5,6;7,8,9},,2))',
  'TEXTJOIN(",",FALSE,TAKE({1,2,3;4,5,6;7,8,9},-2,-1))', 'ROWS(TAKE({1,2,3;4,5,6;7,8,9},5))', 'TAKE({1,2,3},0)', 'TEXTJOIN(",",FALSE,DROP({1,2,3;4,5,6;7,8,9},1))',
  'TEXTJOIN(",",FALSE,DROP({1,2,3;4,5,6;7,8,9},-1,1))', 'DROP({1,2,3;4,5,6},2)', 'TEXTJOIN(",",FALSE,VSTACK({1,2},{3;4}))', 'ROWS(VSTACK({1,2},{3;4}))',
  'TEXTJOIN(",",FALSE,HSTACK({1;2},{3,4}))', 'COLUMNS(HSTACK({1;2},{3,4}))', 'TEXTJOIN(",",FALSE,VSTACK(Data!A2:A3,"x"))',
  'TEXTJOIN(",",FALSE,CHOOSECOLS({1,2,3;4,5,6},1,3))', 'TEXTJOIN(",",FALSE,CHOOSECOLS({1,2,3;4,5,6},-1))', 'TEXTJOIN(",",FALSE,CHOOSEROWS({1,2,3;4,5,6},2))',
  'CHOOSECOLS({1,2,3;4,5,6},4)', 'TEXTJOIN(",",FALSE,TOCOL({1,2,3;4,5,6}))', 'TEXTJOIN(",",FALSE,TOROW({1,2,3;4,5,6},0,TRUE))', 'TEXTJOIN(",",FALSE,TOCOL({1,#N/A,3},2))',
  'TEXTJOIN(",",FALSE,TOCOL(Data!E2:E9,1))', 'TEXTJOIN(",",FALSE,WRAPROWS({1,2,3,4,5},2))', 'TEXTJOIN(",",FALSE,WRAPROWS({1,2,3,4,5},2,0))',
  'ROWS(WRAPCOLS({1,2,3,4,5},2))', 'TEXTJOIN(",",FALSE,WRAPCOLS({1,2,3,4,5},2,"-"))', 'TEXTJOIN(",",FALSE,EXPAND({1,2},2,3))', 'TEXTJOIN(",",FALSE,EXPAND({1,2},2,3,"-"))',
  'EXPAND({1,2},1)', 'TEXTJOIN("|",FALSE,TEXTSPLIT("a,b,c",","))', 'TEXTJOIN("|",FALSE,TEXTSPLIT("a,b;c,d",",",";"))', 'ROWS(TEXTSPLIT("a,b;c,d",",",";"))',
  'TEXTJOIN("|",FALSE,TEXTSPLIT("a,,b",",",,TRUE))', 'TEXTJOIN("|",FALSE,TEXTSPLIT("a-b_c",{"-","_"}))', 'TEXTJOIN("|",FALSE,TEXTSPLIT("aXbxc","x",,,1))',
  'TEXTJOIN("|",FALSE,TEXTSPLIT("a,b;c",",",";"))', 'TEXTJOIN("|",FALSE,TEXTSPLIT("a,b;c",",",";",,,"-"))',
  'TEXTJOIN(",",FALSE,MAP({1,2,3},LAMBDA(x,x*2)))', 'REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))', 'TEXTJOIN(",",FALSE,SCAN(0,{1,2,3},LAMBDA(a,v,a+v)))',
  'TEXTJOIN(",",FALSE,BYROW({1,2;3,4},LAMBDA(r,SUM(r))))', 'TEXTJOIN(",",FALSE,BYCOL({1,2;3,4},LAMBDA(c,MAX(c))))', 'TEXTJOIN(",",FALSE,MAKEARRAY(2,3,LAMBDA(r,c,r*c)))',
  'LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1,)', 'LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1,2)', 'VALUETOTEXT(123)', 'VALUETOTEXT("a",1)', 'ARRAYTOTEXT({1,"a";TRUE,2})',
  'ARRAYTOTEXT({1,"a"},1)',
  // Loans, depreciation and rates
  'IPMT(0.1/12,1,36,8000)', 'PPMT(0.1/12,1,24,2000)', 'IPMT(0.1,3,3,8000)', 'CUMIPMT(0.09/12,360,125000,13,24,0)', 'CUMPRINC(0.09/12,360,125000,13,24,0)',
  'CUMIPMT(0.09/12,360,125000,1,1,0)', 'SLN(30000,7500,10)', 'DDB(2400,300,10,1)', 'DDB(2400,300,10,10)', 'DDB(2400,300,120,1,2)', 'DB(1000000,100000,6,1,7)',
  'DB(1000000,100000,6,7,7)', 'DB(1000000,100000,6,2)', 'EFFECT(0.0525,4)', 'NOMINAL(0.053543,4)', 'XNPV(0.09,{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})',
  'XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})',
  // Statistics
  'CORREL({3,2,4,5,6},{9,7,12,15,17})', 'PEARSON({3,2,4,5,6},{9,7,12,15,17})', 'SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 'INTERCEPT({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})',
  'RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 'FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})', 'FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})',
  'COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})', 'COVARIANCE.S({3,2,4,5,6},{9,7,12,15,17})', 'CORREL(Data!B2:B9,Data!C2:C9)', 'GEOMEAN(4,5,8,7,11,4,3)',
  'HARMEAN(4,5,8,7,11,4,3)', 'AVEDEV(4,5,6,7,5,4,3)', 'DEVSQ(4,5,8,7,11,4,3)', 'SKEW(3,4,5,2,3,4,5,6,4,7)', 'KURT(3,4,5,2,3,4,5,6,4,7)', 'STANDARDIZE(42,40,1.5)',
  'PERMUT(100,3)', 'FISHER(0.75)', 'NORM.DIST(42,40,1.5,TRUE)', 'NORM.DIST(42,40,1.5,FALSE)', 'NORM.S.DIST(1.333333,TRUE)', 'NORM.S.DIST(1.333333,FALSE)',
  'NORM.INV(0.908789,40,1.5)', 'NORM.S.INV(0.908789)', 'NORM.S.INV(0.001)', 'EXPON.DIST(0.2,10,TRUE)', 'EXPON.DIST(0.2,10,FALSE)', 'POISSON.DIST(2,5,FALSE)',
  'POISSON.DIST(2,5,TRUE)', 'BINOM.DIST(6,10,0.5,FALSE)', 'BINOM.DIST(6,10,0.5,TRUE)',
  // Hyperbolic and reciprocal trigonometry, numerals and bases
  'SINH(1)', 'COSH(1)', 'TANH(0.5)', 'SEC(1)', 'CSC(1)', 'COT(1)', 'COT(0)', 'ROMAN(499)', 'ROMAN(1999)', 'ROMAN(4000)', 'ARABIC("MCMXCIX")', 'ARABIC("-IV")',
  'BASE(255,16)', 'BASE(7,2,8)', 'DECIMAL("FF",16)', 'DECIMAL("111",2)', 'DEC2BIN(9)', 'DEC2BIN(-1)', 'DEC2BIN(9,8)', 'BIN2DEC("1111111111")', 'DEC2HEX(255)',
  'DEC2HEX(-1)', 'HEX2DEC("FFFFFFFFFF")', 'HEX2DEC("A5")', 'MULTINOMIAL(2,3,4)',
  // Formulas, sheets, times and working days
  'ISFORMULA(B2)', 'ISFORMULA(A2)', 'FORMULATEXT(B10)', 'FORMULATEXT(A2)', 'SHEET()', 'SHEET("My Data")', 'SHEETS()', 'TIMEVALUE("6:35 PM")',
  'TIMEVALUE("22-Aug-2011 6:35 AM")', 'WORKDAY.INTL(DATE(2024,1,1),10,7)', 'WORKDAY.INTL(DATE(2024,1,1),10,"0000011")', 'NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),11)',
  'NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"0010001")',
  // Distributions: t, chi-squared, F, gamma, beta, lognormal, Weibull, negative binomial, hypergeometric
  'T.DIST(1.96,10,TRUE)', 'T.DIST(1.96,10,FALSE)', 'T.DIST(-0.5,3,TRUE)', 'T.DIST(2,2.5,TRUE)', 'T.DIST.2T(1.96,10)', 'T.DIST.2T(0.5,1)', 'T.DIST.RT(1.96,10)',
  'T.DIST.RT(-1,30)', 'T.INV(0.75,2)', 'T.INV(0.025,10)', 'T.INV(0.999,5)', 'T.INV.2T(0.05,10)', 'T.INV.2T(0.5,60)', 'TDIST(1.96,10,2)', 'TDIST(1.96,10,1)', 'TINV(0.05,10)',
  'CHISQ.DIST(0.5,1,TRUE)', 'CHISQ.DIST(2,3,FALSE)', 'CHISQ.DIST(18.307,10,TRUE)', 'CHISQ.DIST.RT(18.307,10)', 'CHISQ.DIST.RT(3.5,4)', 'CHISQ.INV(0.93,1)', 'CHISQ.INV(0.5,10)',
  'CHISQ.INV.RT(0.05,10)', 'CHISQ.INV.RT(0.5,2)', 'CHIDIST(18.307,10)', 'CHIINV(0.05,10)',
  'F.DIST(15.2069,6,4,TRUE)', 'F.DIST(15.2069,6,4,FALSE)', 'F.DIST(1,10,20,TRUE)', 'F.DIST.RT(15.2069,6,4)', 'F.DIST.RT(2.5,3,30)', 'F.INV(0.01,6,4)', 'F.INV(0.5,10,20)',
  'F.INV.RT(0.01,6,4)', 'F.INV.RT(0.05,3,30)', 'FDIST(15.2069,6,4)', 'FINV(0.01,6,4)',
  'GAMMA(2.5)', 'GAMMA(5)', 'GAMMA(-0.75)', 'GAMMA(0.1)', 'GAMMA(0)', 'GAMMALN(4.5)', 'GAMMALN(100)', 'GAMMALN.PRECISE(0.5)', 'GAMMA.DIST(10.00001131,9,2,TRUE)',
  'GAMMA.DIST(10.00001131,9,2,FALSE)', 'GAMMA.DIST(3,0.5,1,TRUE)', 'GAMMA.INV(0.068094,9,2)', 'GAMMA.INV(0.5,2,3)', 'GAMMADIST(10,9,2,TRUE)', 'GAMMAINV(0.068094,9,2)',
  'BETA.DIST(2,8,10,TRUE,1,3)', 'BETA.DIST(2,8,10,FALSE,1,3)', 'BETA.DIST(0.3,2,5,TRUE)', 'BETA.INV(0.685470581,8,10,1,3)', 'BETA.INV(0.5,2,5)', 'BETADIST(2,8,10,1,3)',
  'BETAINV(0.685470581,8,10,1,3)', 'LOGNORM.DIST(4,3.5,1.2,TRUE)', 'LOGNORM.DIST(4,3.5,1.2,FALSE)', 'LOGNORM.INV(0.039084,3.5,1.2)', 'LOGNORMDIST(4,3.5,1.2)', 'LOGINV(0.039084,3.5,1.2)',
  'WEIBULL.DIST(105,20,100,TRUE)', 'WEIBULL.DIST(105,20,100,FALSE)', 'WEIBULL(105,20,100,TRUE)', 'NEGBINOM.DIST(10,5,0.25,TRUE)', 'NEGBINOM.DIST(10,5,0.25,FALSE)',
  'NEGBINOMDIST(10,5,0.25)', 'HYPGEOM.DIST(1,4,8,20,TRUE)', 'HYPGEOM.DIST(1,4,8,20,FALSE)', 'HYPGEOMDIST(1,4,8,20)', 'BINOM.INV(6,0.5,0.75)', 'BINOM.INV(100,0.3,0.5)', 'CRITBINOM(6,0.5,0.75)',
  'POISSON(2,5,TRUE)', 'BINOMDIST(6,10,0.5,FALSE)', 'NORMDIST(42,40,1.5,TRUE)', 'NORMINV(0.908789,40,1.5)', 'NORMSDIST(1.333333)', 'NORMSINV(0.908789)', 'EXPONDIST(0.2,10,TRUE)',
  // Tests and confidence
  'CONFIDENCE.NORM(0.05,2.5,50)', 'CONFIDENCE.T(0.05,1,50)', 'CONFIDENCE(0.05,2.5,50)', 'Z.TEST({3,6,7,8,6,5,4,2,1,9},4)', 'Z.TEST({3,6,7,8,6,5,4,2,1,9},6,2)', 'ZTEST({3,6,7,8,6,5,4,2,1,9},4)',
  'T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,1)', 'T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},1,2)', 'T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,3)',
  'TTEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,2)', 'CHISQ.TEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})', 'CHITEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})',
  'F.TEST({6,7,9,15,21},{20,28,31,38,40})', 'FTEST({6,7,9,15,21},{20,28,31,38,40})', 'PHI(0.75)', 'GAUSS(2)', 'FISHERINV(0.972955)',
  // Aggregates, ranks, deviations
  'AGGREGATE(4,6,Data!B2:B9)', 'AGGREGATE(9,0,Data!B2:B9)', 'AGGREGATE(14,6,Data!B2:B9,2)', 'AGGREGATE(15,6,Data!C2:C9,1)', 'AGGREGATE(1,6,Data!B2:B9,Data!C2:C9)', 'AGGREGATE(12,6,Data!B2:B9)',
  'AGGREGATE(16,6,{1,2,3,4,5},0.4)', 'AGGREGATE(19,6,{1,2,3,4,5,6,7},1)', 'AGGREGATE(2,0,Data!A2:E9)', 'AGGREGATE(9,6,{1,2,3})', 'AGGREGATE(20,0,Data!B2:B9)',
  'INDEX(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),1)', 'INDEX(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),3)', 'INDEX(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),4)',
  'TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)', 'TRIMMEAN({1,2,3,4,100},0.4)', 'PERCENTRANK({13,12,11,8,4,3,2,1,1,1},2)', 'PERCENTRANK({13,12,11,8,4,3,2,1,1,1},4)', 'PERCENTRANK({13,12,11,8,4,3,2,1,1,1},8)',
  'PERCENTRANK({13,12,11,8,4,3,2,1,1,1},5.75)', 'PERCENTRANK.INC({1,2,3,6,6,6,7,8,9},6,4)', 'PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},7)', 'PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43)',
  'INDEX(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}),1)', 'INDEX(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}),3)', 'STDEVA(Data!B2:B9,TRUE)', 'STDEVPA({1,2,3,"a",TRUE})', 'VARA(1,2,3,TRUE)',
  'VARPA(Data!A2:B9)', 'STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 'PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)', 'PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)',
  // Regression and matrices
  'INDEX(LINEST({1,9,5,7},{0,4,2,3}),1)', 'INDEX(LINEST({1,9,5,7},{0,4,2,3}),2)', 'INDEX(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE),2,1)', 'INDEX(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE),3,1)',
  'INDEX(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE),3,2)', 'INDEX(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE),4,1)', 'INDEX(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE),5,2)', 'INDEX(LINEST({1,9,5,7},{0,4,2,3},FALSE),1)',
  'INDEX(LINEST({142000;144000;151000;150000;139000;169000;126000;142900;163000;169000;149000},{2310,2,2,20;2333,2,2,12;2356,3,1.5,33;2379,3,2,43;2402,2,3,53;2425,4,2,23;2448,2,1.5,99;2471,2,2,34;2494,3,3,23;2517,4,4,55;2540,2,3,22}),1)',
  'INDEX(LINEST({142000;144000;151000;150000;139000;169000;126000;142900;163000;169000;149000},{2310,2,2,20;2333,2,2,12;2356,3,1.5,33;2379,3,2,43;2402,2,3,53;2425,4,2,23;2448,2,1.5,99;2471,2,2,34;2494,3,3,23;2517,4,4,55;2540,2,3,22}),5)',
  'INDEX(LOGEST({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16}),1)', 'INDEX(LOGEST({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16}),2)',
  'INDEX(TREND({133890,135000,135790,137300,138130,139100,139900,141120,141890,143230,144000,145290},{1,2,3,4,5,6,7,8,9,10,11,12},{13,14,15}),1)',
  'INDEX(TREND({133890,135000,135790,137300,138130,139100,139900,141120,141890,143230,144000,145290},{1,2,3,4,5,6,7,8,9,10,11,12},{13,14,15}),3)', 'INDEX(TREND({1,3,4}),1)',
  'INDEX(GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17,18}),1)', 'INDEX(GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17,18}),2)',
  'INDEX(MMULT({1,3;7,2},{2,0;0,2}),2,1)', 'INDEX(MMULT({1,2,3},{4;5;6}),1)', 'MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})', 'MDETERM({3,6,1;1,1,0;3,10,2})', 'MDETERM({1,2;2,4})',
  'INDEX(MINVERSE({4,-1;2,0}),1,2)', 'INDEX(MINVERSE({1,2,1;3,4,-1;0,2,0}),3,2)', 'MINVERSE({1,2;2,4})', 'INDEX(MUNIT(3),2,2)', 'INDEX(MUNIT(3),1,3)',
  'SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 'SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 'SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})',
  // Engineering, hyperbolic, counting
  'BITAND(13,25)', 'BITOR(23,10)', 'BITXOR(5,3)', 'BITLSHIFT(4,2)', 'BITRSHIFT(13,2)', 'BITAND(2^47,2^47+1)', 'BITLSHIFT(1,48)', 'OCT2DEC(54)', 'OCT2DEC(7777777533)', 'DEC2OCT(58,3)', 'DEC2OCT(-100)',
  'BIN2HEX(11111011,4)', 'BIN2HEX(1110)', 'HEX2BIN("F",8)', 'HEX2BIN("FFFFFFFE00")', 'BIN2OCT(1001,3)', 'OCT2BIN(3,3)', 'HEX2OCT("F",3)', 'OCT2HEX(100,4)', 'OCT2HEX(7777777533)',
  'DELTA(5,4)', 'DELTA(5,5)', 'DELTA(0.5)', 'GESTEP(5,4)', 'GESTEP(-4,-5)', 'GESTEP(-1)', 'ERF(0.745)', 'ERF(1,2)', 'ERF.PRECISE(1)', 'ERFC(1)', 'ERFC.PRECISE(-0.5)',
  'ACOSH(10)', 'ASINH(-2.5)', 'ATANH(0.76159416)', 'ACOT(2)', 'ACOT(-2)', 'ACOTH(6)', 'COTH(2)', 'CSCH(1.5)', 'SECH(1)',
  'COMBINA(4,3)', 'COMBINA(10,3)', 'PERMUTATIONA(3,2)', 'FACTDOUBLE(6)', 'FACTDOUBLE(7)', 'FACTDOUBLE(-1)', 'SERIESSUM(PI()/4,0,2,{1,-0.5,0.0416666666666667,-0.00138888888888889})', 'SQRTPI(2)',
  // Dates, links, areas and money
  'DAYS360(DATE(2011,1,30),DATE(2011,12,31))', 'DAYS360(DATE(2011,2,28),DATE(2011,3,31))', 'DAYS360(DATE(2011,2,28),DATE(2011,3,31),TRUE)', 'DAYS360(DATE(2012,2,29),DATE(2013,2,28))',
  'DAYS360(DATE(2011,1,31),DATE(2011,1,1))', 'HYPERLINK("http://example.com","Example")', 'HYPERLINK("http://example.com")', 'AREAS(Data!B2:D4)', 'AREAS((Data!B2:D4,Data!E5,Data!F6:I9))',
  'FVSCHEDULE(1,{0.09,0.11,0.1})', 'PDURATION(0.025,2000,2200)', 'RRI(96,10000,11000)', 'ISPMT(0.1/12,1,36,8000000)', 'MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)',
  'SYD(30000,7500,10,1)', 'SYD(30000,7500,10,10)', 'VDB(2400,300,10*365,0,1)', 'VDB(2400,300,10,0,1)', 'VDB(2400,300,10,0,0.875,1.5)', 'VDB(2400,300,10,6,10,2,TRUE)', 'VDB(2400,300,10,3,8)',
  'DOLLARDE(1.02,16)', 'DOLLARDE(1.1,32)', 'DOLLARFR(1.125,16)', 'DOLLARFR(1.125,32)', 'TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)', 'TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)',
  'TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)', 'TBILLEQ(DATE(2008,1,1),DATE(2008,12,1),0.09)',
];
export function formulas() {
  const str = (ref, s) => `<c r="${ref}" t="inlineStr"><is><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</t></is></c>`;
  const num = (ref, v) => `<c r="${ref}"><v>${v}</v></c>`;
  const cell = (ref, v) => (v == null ? '' : typeof v === 'number' ? num(ref, v) : typeof v === 'boolean' ? `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>` :
    typeof v === 'object' ? `<c r="${ref}" t="e"><v>${v.err}</v></c>` : str(ref, v));
  const data = [
    ['Name', 'Qty', 'Price', 'Date', 'Flag', null, 'Err', null, 'Score', 'Grade'],
    ['Apple', 3, 1.5, 45292, true, null, { err: '#N/A' }, null, 0, 'F'],
    ['Banana', 5, 0.25, 45323, false, null, 5, null, 50, 'D'],
    ['Cherry', '12', 4, 45352, true, null, null, null, 60, 'C'],
    ['Apple', 2, 2, 45383, null, null, null, null, 70, 'B'],
    ['Elder', 8, 3.5, 45413, false, null, null, null, 80, 'A'],
    ['Fig', null, 6, 45444, true],
    ['Grape', 4, 0.75, 45474, 'yes'],
    ['Apple', 6, 1.25, 45505, false],
  ];
  const dataRows = data.map((row, i) => `<row r="${i + 1}">${row.map((v, c) => cell(`${col(c)}${i + 1}`, v)).join('')}</row>`);
  const calcRows = PROBE_FORMULAS.map((f, i) => {
    const r = i + 2;
    // (D2:D10 a chain, D3:D10 as one shared formula; D11 their sum; E2 reads E3, below it.)
    const chain = r === 2 ? num('D2', 1) : r <= 10 ? `<c r="D${r}">${r === 3 ? '<f t="shared" ref="D3:D10" si="0">D2+1</f>' : '<f t="shared" si="0"/>'}</c>` :
      r === 11 ? '<c r="D11"><f>SUM(D2:D10)</f></c>' : '';
    const fwd = r === 2 ? '<c r="E2"><f>E3*2</f></c>' : r === 3 ? num('E3', 5) : '';
    // (F2:F4 and F12: a table's This Row from another sheet, by the row's number; row 12 is past the table's rows.)
    const thisRow = r <= 4 || r === 12 ? `<c r="F${r}"><f>Fruit[[#This Row],[Price]]*2</f></c>` : '';
    // (As the file has them: newer functions with their _xlfn. prefix.)
    return `<row r="${r}">${str(`A${r}`, f)}<c r="B${r}"><f>${esc(normalizeFormula(f))}</f></c>${chain}${fwd}${thisRow}</row>`;
  });
  const sheet = (rows, cols, tail = '') => `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rows.join('')}</sheetData>${tail}</worksheet>`;
  const table = `${DECL}<table xmlns="${X}" id="1" name="Fruit" displayName="Fruit" ref="A1:E9" totalsRowShown="0"><autoFilter ref="A1:E9"/>` +
    '<tableColumns count="5"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Price"/><tableColumn id="4" name="Date"/><tableColumn id="5" name="Flag"/></tableColumns>' +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
  write('formulas.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/worksheets/sheet2.xml', CT('worksheet')],
      ['xl/worksheets/sheet3.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/tables/table1.xml', CT('table')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Calc" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/>` +
      '<sheet name="My Data" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="Items">Data!$A$2:$A$9</definedName>' +
      '<definedName name="Rate">Data!$C$2</definedName></definedNames></workbook>',
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('worksheet'), 'worksheets/sheet2.xml'],
      ['rId3', RT('worksheet'), 'worksheets/sheet3.xml'], ['rId4', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': sheet(calcRows, '<cols><col min="1" max="1" width="50.7109375" customWidth="1"/><col min="2" max="2" width="30.7109375" customWidth="1"/></cols>'),
    'xl/worksheets/sheet2.xml': sheet(dataRows, '', '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'),
    'xl/worksheets/_rels/sheet2.xml.rels': rels([['rId1', RT('table'), '../tables/table1.xml']]),
    'xl/tables/table1.xml': table,
    'xl/worksheets/sheet3.xml': sheet([`<row r="1">${num('A1', 100)}</row>`, `<row r="2">${str('A2', 'text')}</row>`], ''),
    'xl/styles.xml': STYLES,
  });
}
