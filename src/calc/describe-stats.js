// What each Statistical, Compatibility and Database function does, one sentence each, for the Function Arguments box
// (describe.js). Written for this project in its own words. One line per function: "NAME: sentence".

export const STATISTICAL = `
AVEDEV: Gives the average of how far values are from their mean.
AVERAGE: Gives the average (arithmetic mean) of its arguments.
AVERAGEA: Gives the average of its arguments, counting text as 0 and TRUE as 1.
AVERAGEIF: Gives the average of the cells that meet a condition.
AVERAGEIFS: Gives the average of the cells that meet every one of a set of conditions.
BETA.DIST: Gives the beta distribution, cumulative or as a density.
BETA.INV: Gives the inverse of the cumulative beta distribution.
BINOM.DIST: Gives the binomial distribution probability of a number of successes.
BINOM.DIST.RANGE: Gives the probability of a number of successes, or of a span of them, by the binomial distribution.
BINOM.INV: Gives the smallest number of successes whose cumulative binomial probability reaches a criterion.
CHISQ.DIST: Gives the left-tailed chi-squared distribution.
CHISQ.DIST.RT: Gives the right-tailed chi-squared distribution.
CHISQ.INV: Gives the inverse of the left-tailed chi-squared distribution.
CHISQ.INV.RT: Gives the inverse of the right-tailed chi-squared distribution.
CHISQ.TEST: Gives the chi-squared test of independence of observed against expected values.
CONFIDENCE.NORM: Gives the half-width of a confidence interval for a mean, by the normal distribution.
CONFIDENCE.T: Gives the half-width of a confidence interval for a mean, by Student's t-distribution.
CORREL: Gives the correlation coefficient of two sets of values.
COUNT: Counts the numbers among its arguments.
COUNTA: Counts the values among its arguments that aren't empty.
COUNTBLANK: Counts the empty cells in a range.
COUNTIF: Counts the cells in a range that meet a condition.
COUNTIFS: Counts the cells that meet every one of a set of conditions.
COVARIANCE.P: Gives the covariance of two sets of values, taken as a whole population.
COVARIANCE.S: Gives the covariance of two sets of values, taken as a sample.
DEVSQ: Adds up the squares of how far values are from their mean.
EXPON.DIST: Gives the exponential distribution.
F.DIST: Gives the left-tailed F probability distribution.
F.DIST.RT: Gives the right-tailed F probability distribution.
F.INV: Gives the inverse of the left-tailed F distribution.
F.INV.RT: Gives the inverse of the right-tailed F distribution.
F.TEST: Gives the result of an F-test, whether two sets of values differ in variance.
FISHER: Gives the Fisher transformation of a number.
FISHERINV: Gives the inverse of the Fisher transformation.
FORECAST: Predicts a value on the straight line that best fits known values.
FORECAST.LINEAR: Predicts a value on the straight line that best fits known values.
FREQUENCY: Counts how many values fall into each of a set of bins, as a column.
GAMMA: Gives the gamma function of a number.
GAMMA.DIST: Gives the gamma distribution.
GAMMA.INV: Gives the inverse of the cumulative gamma distribution.
GAMMALN: Gives the natural logarithm of the gamma function.
GAMMALN.PRECISE: Gives the natural logarithm of the gamma function.
GAUSS: Gives the probability that a standard normal value lies between the mean and a number of deviations from it.
GEOMEAN: Gives the geometric mean of positive numbers.
GROWTH: Gives values on the exponential curve that best fits known values.
HARMEAN: Gives the harmonic mean of positive numbers.
HYPGEOM.DIST: Gives the hypergeometric distribution.
INTERCEPT: Gives where the line that best fits known values crosses the y axis.
KURT: Gives the kurtosis of a set of values.
LARGE: Gives the k-th largest value of a set.
LINEST: Gives the statistics of the straight line that best fits known values.
LOGEST: Gives the statistics of the exponential curve that best fits known values.
LOGNORM.DIST: Gives the lognormal distribution.
LOGNORM.INV: Gives the inverse of the cumulative lognormal distribution.
MAX: Gives the largest number among its arguments.
MAXA: Gives the largest value among its arguments, counting text as 0 and TRUE as 1.
MAXIFS: Gives the largest of the cells that meet every one of a set of conditions.
MEDIAN: Gives the median, the number in the middle, of its arguments.
MIN: Gives the smallest number among its arguments.
MINA: Gives the smallest value among its arguments, counting text as 0 and TRUE as 1.
MINIFS: Gives the smallest of the cells that meet every one of a set of conditions.
MODE.MULT: Gives a column of the values that occur most often.
MODE.SNGL: Gives the value that occurs most often.
NEGBINOM.DIST: Gives the negative binomial distribution.
NORM.DIST: Gives the normal distribution for a mean and standard deviation.
NORM.INV: Gives the inverse of the cumulative normal distribution.
NORM.S.DIST: Gives the standard normal distribution, of mean 0 and standard deviation 1.
NORM.S.INV: Gives the inverse of the cumulative standard normal distribution.
PEARSON: Gives the Pearson correlation coefficient of two sets of values.
PERCENTILE.EXC: Gives the k-th percentile of values, k strictly between 0 and 1.
PERCENTILE.INC: Gives the k-th percentile of values, k from 0 to 1.
PERCENTRANK.EXC: Gives the rank of a value in a set as a share strictly between 0 and 1.
PERCENTRANK.INC: Gives the rank of a value in a set as a share from 0 to 1.
PERMUT: Counts the ordered ways to choose a number of items from a set.
PERMUTATIONA: Counts the ordered ways to choose a number of items from a set when items may repeat.
PHI: Gives the density of the standard normal distribution.
POISSON.DIST: Gives the Poisson distribution.
PROB: Gives the probability that values lie between two limits.
QUARTILE.EXC: Gives a quartile of a set of values, from percentiles strictly between 0 and 1.
QUARTILE.INC: Gives a quartile of a set of values, from percentiles 0 to 1.
RANK.AVG: Gives the rank of a number in a list, ties given their average rank.
RANK.EQ: Gives the rank of a number in a list, ties given the top rank.
RSQ: Gives the square of the Pearson correlation coefficient.
SKEW: Gives the skewness of a distribution, from a sample.
SKEW.P: Gives the skewness of a distribution, from a whole population.
SLOPE: Gives the slope of the line that best fits known values.
SMALL: Gives the k-th smallest value of a set.
STANDARDIZE: Gives a value as a number of standard deviations from a mean.
STDEV.P: Gives the standard deviation of a whole population.
STDEV.S: Gives the standard deviation of a sample.
STDEVA: Gives the standard deviation of a sample, counting text as 0 and TRUE as 1.
STDEVPA: Gives the standard deviation of a whole population, counting text as 0 and TRUE as 1.
STEYX: Gives the standard error of the predicted y values of a linear fit.
T.DIST: Gives the left-tailed Student's t-distribution.
T.DIST.2T: Gives the two-tailed Student's t-distribution.
T.DIST.RT: Gives the right-tailed Student's t-distribution.
T.INV: Gives the left-tailed inverse of Student's t-distribution.
T.INV.2T: Gives the two-tailed inverse of Student's t-distribution.
T.TEST: Gives the probability of a Student's t-test.
TREND: Gives values on the straight line that best fits known values.
TRIMMEAN: Gives the mean of values with a share of the highest and lowest left out.
VAR.P: Gives the variance of a whole population.
VAR.S: Gives the variance of a sample.
VARA: Gives the variance of a sample, counting text as 0 and TRUE as 1.
VARPA: Gives the variance of a whole population, counting text as 0 and TRUE as 1.
WEIBULL.DIST: Gives the Weibull distribution.
Z.TEST: Gives the one-tailed probability of a z-test.
`;

export const COMPATIBILITY = `
BETADIST: Gives the cumulative beta distribution (the older name of BETA.DIST).
BETAINV: Gives the inverse of the cumulative beta distribution (the older name of BETA.INV).
BINOMDIST: Gives the binomial distribution probability (the older name of BINOM.DIST).
CHIDIST: Gives the right-tailed chi-squared distribution (the older name of CHISQ.DIST.RT).
CHIINV: Gives the inverse of the right-tailed chi-squared distribution (the older name of CHISQ.INV.RT).
CHITEST: Gives the chi-squared test of independence (the older name of CHISQ.TEST).
CONFIDENCE: Gives a confidence interval's half-width by the normal distribution (the older name of CONFIDENCE.NORM).
COVAR: Gives the population covariance of two sets of values (the older name of COVARIANCE.P).
CRITBINOM: Gives the smallest number of successes reaching a cumulative probability (the older name of BINOM.INV).
EXPONDIST: Gives the exponential distribution (the older name of EXPON.DIST).
FDIST: Gives the right-tailed F distribution (the older name of F.DIST.RT).
FINV: Gives the inverse of the right-tailed F distribution (the older name of F.INV.RT).
FTEST: Gives the result of an F-test (the older name of F.TEST).
GAMMADIST: Gives the gamma distribution (the older name of GAMMA.DIST).
GAMMAINV: Gives the inverse of the cumulative gamma distribution (the older name of GAMMA.INV).
HYPGEOMDIST: Gives the hypergeometric distribution (the older form of HYPGEOM.DIST).
LOGINV: Gives the inverse of the cumulative lognormal distribution (the older name of LOGNORM.INV).
LOGNORMDIST: Gives the cumulative lognormal distribution (the older form of LOGNORM.DIST).
MODE: Gives the value that occurs most often (the older name of MODE.SNGL).
NEGBINOMDIST: Gives the negative binomial distribution (the older form of NEGBINOM.DIST).
NORMDIST: Gives the normal distribution (the older name of NORM.DIST).
NORMINV: Gives the inverse of the cumulative normal distribution (the older name of NORM.INV).
NORMSDIST: Gives the cumulative standard normal distribution (the older form of NORM.S.DIST).
NORMSINV: Gives the inverse of the standard normal distribution (the older name of NORM.S.INV).
PERCENTILE: Gives the k-th percentile of values (the older name of PERCENTILE.INC).
PERCENTRANK: Gives the rank of a value as a share of the set (the older name of PERCENTRANK.INC).
POISSON: Gives the Poisson distribution (the older name of POISSON.DIST).
QUARTILE: Gives a quartile of a set of values (the older name of QUARTILE.INC).
RANK: Gives the rank of a number in a list (the older name of RANK.EQ).
STDEV: Gives the standard deviation of a sample (the older name of STDEV.S).
STDEVP: Gives the standard deviation of a whole population (the older name of STDEV.P).
TDIST: Gives Student's t-distribution, one- or two-tailed (the older form of T.DIST.RT and T.DIST.2T).
TINV: Gives the two-tailed inverse of Student's t-distribution (the older name of T.INV.2T).
TTEST: Gives the probability of a Student's t-test (the older name of T.TEST).
VAR: Gives the variance of a sample (the older name of VAR.S).
VARP: Gives the variance of a whole population (the older name of VAR.P).
WEIBULL: Gives the Weibull distribution (the older name of WEIBULL.DIST).
ZTEST: Gives the one-tailed probability of a z-test (the older name of Z.TEST).
`;

export const DATABASE = `
DAVERAGE: Gives the average of a field over the records of a list that meet criteria.
DCOUNT: Counts the numbers in a field over the records of a list that meet criteria.
DCOUNTA: Counts the cells that aren't empty in a field over the records that meet criteria.
DGET: Gives the one value of a field from the record that meets criteria.
DMAX: Gives the largest value of a field over the records that meet criteria.
DMIN: Gives the smallest value of a field over the records that meet criteria.
DPRODUCT: Multiplies the values of a field over the records that meet criteria.
DSTDEV: Gives the sample standard deviation of a field over the records that meet criteria.
DSTDEVP: Gives the population standard deviation of a field over the records that meet criteria.
DSUM: Adds up a field over the records of a list that meet criteria.
DVAR: Gives the sample variance of a field over the records that meet criteria.
DVARP: Gives the population variance of a field over the records that meet criteria.
`;
