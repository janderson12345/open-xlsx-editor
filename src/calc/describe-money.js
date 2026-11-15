// What each Financial and Date & Time function does, one sentence each, for the Function Arguments box (describe.js).
// Written for this project in its own words. One line per function: "NAME: sentence".

export const FINANCIAL = `
ACCRINT: Works out the interest a security that pays interest periodically has built up since its issue.
ACCRINTM: Works out the interest a security that pays all its interest at maturity has built up.
AMORDEGRC: Gives the depreciation of an asset for one accounting period, on the French declining system with a coefficient.
AMORLINC: Gives the depreciation of an asset for one accounting period, on the French straight-line system.
COUPDAYBS: Counts the days from the start of the coupon period to the settlement date.
COUPDAYS: Counts the days in the coupon period that holds the settlement date.
COUPDAYSNC: Counts the days from the settlement date to the next coupon date.
COUPNCD: Gives the first coupon date after the settlement date.
COUPNUM: Counts the coupons still to be paid between settlement and maturity.
COUPPCD: Gives the last coupon date on or before the settlement date.
CUMIPMT: Adds up the interest paid on a loan between two periods.
CUMPRINC: Adds up the principal repaid on a loan between two periods.
DB: Gives an asset's depreciation for a period by the fixed-declining-balance method.
DDB: Gives an asset's depreciation for a period by the double-declining-balance method, or another rate you set.
DISC: Gives the discount rate of a security.
DOLLARDE: Turns a price written as a whole number and a fraction into a decimal number.
DOLLARFR: Turns a decimal price into a whole number and a fraction.
DURATION: Gives the Macaulay duration of a security that pays interest periodically, for a face value of 100.
EFFECT: Gives the effective yearly interest rate from a nominal rate and how often it compounds.
FV: Works out the future value of an investment of equal payments at a constant interest rate.
FVSCHEDULE: Works out the future value of a principal after a series of interest rates is applied in turn.
INTRATE: Gives the interest rate of a security that is fully invested.
IPMT: Gives the interest part of one payment of a loan or investment.
IRR: Gives the internal rate of return of a series of cash flows.
ISPMT: Gives the interest paid in one period of a loan repaid in equal parts of principal.
MDURATION: Gives the modified duration of a security, for a face value of 100.
MIRR: Gives the internal rate of return when money paid out and money taken in earn different rates.
NOMINAL: Gives the nominal yearly interest rate from an effective rate and how often it compounds.
NPER: Works out how many periods a loan or investment takes with equal payments at a constant rate.
NPV: Gives the net present value of future cash flows at a discount rate.
ODDFPRICE: Gives the price per 100 of face value of a security whose first period is odd.
ODDFYIELD: Gives the yield of a security whose first period is odd.
ODDLPRICE: Gives the price per 100 of face value of a security whose last period is odd.
ODDLYIELD: Gives the yield of a security whose last period is odd.
PDURATION: Works out how many periods an investment needs to grow to a given value.
PMT: Works out the payment of a loan with equal payments at a constant interest rate.
PPMT: Gives the principal part of one payment of a loan or investment.
PRICE: Gives the price per 100 of face value of a security that pays interest periodically.
PRICEDISC: Gives the price per 100 of face value of a discounted security.
PRICEMAT: Gives the price per 100 of face value of a security that pays its interest at maturity.
PV: Works out the present value of a series of equal future payments.
RATE: Works out the interest rate per period of a loan or annuity.
RECEIVED: Gives the amount paid back at maturity for a fully invested security.
RRI: Gives the interest rate that makes an investment grow from one value to another.
SLN: Gives an asset's depreciation for one period by the straight-line method.
SYD: Gives an asset's depreciation for a period by the sum-of-years' digits method.
TBILLEQ: Gives the bond-equivalent yield of a Treasury bill.
TBILLPRICE: Gives the price per 100 of face value of a Treasury bill.
TBILLYIELD: Gives the yield of a Treasury bill.
VDB: Gives an asset's depreciation over any span of periods by a declining-balance method.
XIRR: Gives the internal rate of return of cash flows on dates that need not be evenly spaced.
XNPV: Gives the net present value of cash flows on dates that need not be evenly spaced.
YIELD: Gives the yield of a security that pays interest periodically.
YIELDDISC: Gives the yearly yield of a discounted security.
YIELDMAT: Gives the yearly yield of a security that pays its interest at maturity.
`;

export const DATES = `
DATE: Makes a date from a year, a month and a day.
DATEVALUE: Turns a date written as text into a date serial number.
DAY: Gives the day of the month of a date, from 1 to 31.
DAYS: Counts the days between two dates.
DAYS360: Counts the days between two dates on a year of twelve 30-day months.
EDATE: Gives the date a number of months before or after a start date.
EOMONTH: Gives the last day of the month a number of months before or after a start date.
HOUR: Gives the hour of a time, from 0 to 23.
ISOWEEKNUM: Gives the ISO week number of the year for a date.
MINUTE: Gives the minute of a time, from 0 to 59.
MONTH: Gives the month of a date, from 1 (January) to 12 (December).
NETWORKDAYS: Counts the working days between two dates, leaving out weekends and holidays.
NETWORKDAYS.INTL: Counts the working days between two dates, with the weekend days you choose.
NOW: Gives the current date and time.
SECOND: Gives the second of a time, from 0 to 59.
TIME: Makes a time from an hour, a minute and a second.
TIMEVALUE: Turns a time written as text into a fraction of a day.
TODAY: Gives the current date.
WEEKDAY: Gives the day of the week of a date as a number.
WEEKNUM: Gives the week number of the year for a date.
WORKDAY: Gives the date a number of working days before or after a start date.
WORKDAY.INTL: Gives the date a number of working days away, with the weekend days you choose.
YEAR: Gives the year of a date.
YEARFRAC: Gives the part of a year between two dates.
`;
