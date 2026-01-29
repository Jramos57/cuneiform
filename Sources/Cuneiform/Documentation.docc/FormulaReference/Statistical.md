# Statistical Functions

Comprehensive statistical analysis and data modeling functions for spreadsheet calculations in Cuneiform.

## Overview

Cuneiform provides a complete suite of statistical functions compatible with Excel formulas. These functions cover basic descriptive statistics, counting operations, probability distributions, regression analysis, hypothesis testing, and advanced statistical computations.

This is the largest category of functions in Cuneiform's formula evaluator, with over 100 statistical functions spanning basic data aggregation to sophisticated distribution analysis.

All statistical functions are implemented in the `FormulaEvaluator` and can be used in cell formulas throughout your spreadsheet application.

## Quick Reference

### Basic Statistics
- ``MIN`` - Minimum value in dataset
- ``MAX`` - Maximum value in dataset
- ``AVERAGE`` - Arithmetic mean
- ``MEDIAN`` - Middle value
- ``MODE``, ``MODE.SNGL`` - Most frequent single value
- ``MODE.MULT`` - Multiple most frequent values
- ``GEOMEAN`` - Geometric mean
- ``HARMEAN`` - Harmonic mean
- ``TRIMMEAN`` - Mean excluding extreme values

### Counting Functions
- ``COUNT`` - Count numbers
- ``COUNTA`` - Count non-empty cells
- ``COUNTBLANK`` - Count empty cells
- ``COUNTIF`` - Count with single condition
- ``COUNTIFS`` - Count with multiple conditions
- ``FREQUENCY`` - Frequency distribution

### Conditional Aggregation
- ``AVERAGEIF`` - Average with single condition
- ``AVERAGEIFS`` - Average with multiple conditions
- ``MAXIFS`` - Maximum with conditions
- ``MINIFS`` - Minimum with conditions

### Variability & Dispersion
- ``VAR``, ``VAR.S`` - Sample variance
- ``VAR.P`` - Population variance
- ``VARA`` - Sample variance (includes text/logical)
- ``VARPA`` - Population variance (includes text/logical)
- ``STDEV``, ``STDEV.S`` - Sample standard deviation
- ``STDEV.P`` - Population standard deviation
- ``STDEVA`` - Sample std dev (includes text/logical)
- ``STDEVPA`` - Population std dev (includes text/logical)
- ``AVEDEV`` - Average deviation
- ``DEVSQ`` - Sum of squared deviations

### Percentiles & Quartiles
- ``PERCENTILE``, ``PERCENTILE.INC`` - k-th percentile (inclusive)
- ``PERCENTILE.EXC`` - k-th percentile (exclusive)
- ``PERCENTRANK.INC`` - Rank as percentage (inclusive)
- ``PERCENTRANK.EXC`` - Rank as percentage (exclusive)
- ``QUARTILE``, ``QUARTILE.INC`` - Quartile value (inclusive)
- ``QUARTILE.EXC`` - Quartile value (exclusive)

### Ranking & Position
- ``RANK``, ``RANK.EQ`` - Rank with equal values handling
- ``RANK.AVG`` - Rank using average for ties
- ``LARGE`` - k-th largest value
- ``SMALL`` - k-th smallest value

### Distribution Shape
- ``SKEW`` - Sample skewness
- ``SKEW.P`` - Population skewness
- ``KURT`` - Kurtosis

### Normal Distribution
- ``NORM.DIST`` - Normal distribution
- ``NORM.INV`` - Inverse normal distribution
- ``NORM.S.DIST`` - Standard normal distribution
- ``NORM.S.INV`` - Inverse standard normal
- ``STANDARDIZE`` - Z-score calculation
- ``NORMDIST``, ``NORMINV`` - Legacy normal functions
- ``NORMSDIST``, ``NORMSINV`` - Legacy standard normal functions

### Discrete Distributions
- ``BINOM.DIST`` - Binomial distribution
- ``BINOM.INV`` - Inverse binomial
- ``POISSON.DIST``, ``POISSON`` - Poisson distribution
- ``HYPGEOM.DIST``, ``HYPGEOMDIST`` - Hypergeometric distribution
- ``NEGBINOM.DIST``, ``NEGBINOMDIST`` - Negative binomial distribution
- ``CRITBINOM`` - Critical value for binomial

### Continuous Distributions
- ``EXPON.DIST``, ``EXPONDIST`` - Exponential distribution
- ``LOGNORM.DIST``, ``LOGNORMDIST`` - Lognormal distribution
- ``LOGNORM.INV``, ``LOGNORMINV``, ``LOGINV`` - Inverse lognormal
- ``WEIBULL.DIST``, ``WEIBULL`` - Weibull distribution
- ``GAMMA.DIST``, ``GAMMADIST`` - Gamma distribution
- ``GAMMA.INV``, ``GAMMAINV`` - Inverse gamma distribution
- ``GAMMA`` - Gamma function value
- ``GAMMALN``, ``GAMMALN.PRECISE`` - Natural log of gamma function
- ``BETA.DIST``, ``BETADIST`` - Beta distribution
- ``BETA.INV``, ``BETAINV`` - Inverse beta distribution

### Chi-Square, T, and F Distributions
- ``CHISQ.DIST`` - Chi-square distribution
- ``CHISQ.INV`` - Inverse chi-square
- ``T.DIST`` - Student's t-distribution
- ``T.DIST.RT`` - Right-tailed t-distribution
- ``T.DIST.2T`` - Two-tailed t-distribution
- ``T.INV`` - Inverse t-distribution
- ``T.INV.2T`` - Two-tailed inverse t
- ``F.DIST`` - F probability distribution
- ``F.INV`` - Inverse F distribution

### Hypothesis Testing
- ``CHISQ.TEST``, ``CHITEST`` - Chi-square test
- ``F.TEST``, ``FTEST`` - F-test
- ``T.TEST``, ``TTEST`` - T-test
- ``Z.TEST``, ``ZTEST`` - Z-test

### Correlation & Covariance
- ``CORREL`` - Correlation coefficient
- ``PEARSON`` - Pearson correlation
- ``COVARIANCE.P``, ``COVAR`` - Population covariance
- ``COVARIANCE.S`` - Sample covariance
- ``FISHER`` - Fisher transformation
- ``FISHERINV`` - Inverse Fisher transformation

### Regression Analysis
- ``SLOPE`` - Slope of linear regression
- ``INTERCEPT`` - Y-intercept of linear regression
- ``STEYX`` - Standard error of regression
- ``RSQ`` - R-squared coefficient of determination
- ``FORECAST``, ``FORECAST.LINEAR`` - Linear forecast
- ``TREND`` - Linear trend values
- ``GROWTH`` - Exponential growth trend
- ``LINEST`` - Linear regression statistics
- ``LOGEST`` - Exponential regression statistics

### Time Series Forecasting
- ``FORECAST.ETS`` - Exponential smoothing forecast
- ``FORECAST.ETS.CONFINT`` - Confidence interval for forecast
- ``FORECAST.ETS.SEASONALITY`` - Seasonality length detection
- ``FORECAST.ETS.STAT`` - Statistical values for forecast

### Confidence Intervals
- ``CONFIDENCE.NORM``, ``CONFIDENCE`` - Normal confidence interval
- ``CONFIDENCE.T`` - Student's t confidence interval

### Probability
- ``PROB`` - Probability in range

### Aggregation Functions
- ``SUBTOTAL`` - Subtotal with function number
- ``AGGREGATE`` - Aggregate with options
- ``AVERAGEA`` - Average including text/logical
- ``MINA`` - Minimum including text/logical
- ``MAXA`` - Maximum including text/logical

### Database Statistical Functions
- ``DSTDEV`` - Database standard deviation
- ``DVAR`` - Database variance

## Function Details

### AGGREGATE

Returns an aggregate calculation (like SUM, AVERAGE, etc.) with options to ignore hidden rows, error values, or nested subtotals.

**Syntax:** `AGGREGATE(function_num, options, ref1, [ref2], ...)`

**Parameters:**
- `function_num`: The aggregation function (1-19)
  - 1: AVERAGE, 2: COUNT, 3: COUNTA, 4: MAX, 5: MIN
  - 6: PRODUCT, 7: STDEV.S, 8: STDEV.P, 9: SUM
  - 10: VAR.S, 11: VAR.P, 12: MEDIAN, 13: MODE.SNGL
  - 14: LARGE, 15: SMALL, 16: PERCENTILE.INC, 17: QUARTILE.INC
  - 18: PERCENTILE.EXC, 19: QUARTILE.EXC
- `options`: Behavior flags (0-7)
  - 0/1: Ignore nested SUBTOTAL and AGGREGATE
  - 2/3: Ignore hidden rows
  - 4/5: Ignore error values
  - 6/7: Combine hidden rows + errors
- `ref1, ref2, ...`: Data ranges

**Returns:** Number - The aggregated result

**Examples:**
```swift
let result1 = evaluator.evaluate("=AGGREGATE(1, 0, A1:A10)")  // Average ignoring subtotals
let result2 = evaluator.evaluate("=AGGREGATE(9, 6, B1:B100)") // Sum ignoring errors+hidden
let result3 = evaluator.evaluate("=AGGREGATE(4, 5, C:C)")     // Max ignoring errors
```

**Excel Documentation:** [AGGREGATE function](https://support.microsoft.com/en-us/office/aggregate-function-43b9278e-6aa7-4f17-92b6-e19993fa26df)

**Implementation Status:** ✅ Full implementation

---

### AVEDEV

Returns the average of the absolute deviations of data points from their mean.

**Syntax:** `AVEDEV(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges to analyze

**Returns:** Number - Average absolute deviation

**Examples:**
```swift
let result1 = evaluator.evaluate("=AVEDEV(4, 5, 6, 7, 5, 4, 3)")  // ≈1.02
let result2 = evaluator.evaluate("=AVEDEV(A1:A10)")               // Variability measure
```

**Excel Documentation:** [AVEDEV function](https://support.microsoft.com/en-us/office/avedev-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639)

**Implementation Status:** ✅ Full implementation

---

### AVERAGE

Returns the arithmetic mean of a set of numbers.

**Syntax:** `AVERAGE(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers, ranges, or arrays to average (ignores text and logical values)

**Returns:** Number - The mean

**Examples:**
```swift
let result1 = evaluator.evaluate("=AVERAGE(10, 20, 30)")      // 20
let result2 = evaluator.evaluate("=AVERAGE(A1:A10)")          // Average of range
let result3 = evaluator.evaluate("=AVERAGE(1, 2, \"3\")")     // 1.5 (text ignored)
```

**Excel Documentation:** [AVERAGE function](https://support.microsoft.com/en-us/office/average-function-047bac88-d466-426c-a32b-8f33eb960cf6)

**Implementation Status:** ✅ Full implementation

---

### AVERAGEA

Returns the average including text and logical values (text=0, TRUE=1, FALSE=0).

**Syntax:** `AVERAGEA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values to average (includes text and logical)

**Returns:** Number - The mean

**Examples:**
```swift
let result1 = evaluator.evaluate("=AVERAGEA(10, 20, TRUE)")       // 10.33
let result2 = evaluator.evaluate("=AVERAGEA(5, \"text\", 15)")    // 6.67
```

**Excel Documentation:** [AVERAGEA function](https://support.microsoft.com/en-us/office/averagea-function-f5f84098-d453-4f4c-bbba-3d2c66356091)

**Implementation Status:** ✅ Full implementation

---

### AVERAGEIF

Returns the average of cells that meet a single criterion.

**Syntax:** `AVERAGEIF(range, criteria, [average_range])`

**Parameters:**
- `range`: The range to evaluate against the criteria
- `criteria`: The condition to test (number, text, expression like ">10")
- `average_range` *(optional)*: Cells to average (defaults to range)

**Returns:** Number - The conditional average

**Examples:**
```swift
let result1 = evaluator.evaluate("=AVERAGEIF(A1:A10, \">50\")")           // Avg > 50
let result2 = evaluator.evaluate("=AVERAGEIF(B1:B10, \"Yes\", C1:C10)")   // Conditional avg
let result3 = evaluator.evaluate("=AVERAGEIF(D:D, \"<>0\")")              // Avg non-zero
```

**Excel Documentation:** [AVERAGEIF function](https://support.microsoft.com/en-us/office/averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642)

**Implementation Status:** ✅ Full implementation

---

### AVERAGEIFS

Returns the average of cells that meet multiple criteria.

**Syntax:** `AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Parameters:**
- `average_range`: The range of cells to average
- `criteria_range1, criteria1`: First condition range and criteria
- `criteria_range2, criteria2, ...`: Additional condition pairs

**Returns:** Number - The conditional average

**Examples:**
```swift
let result1 = evaluator.evaluate("=AVERAGEIFS(A1:A10, B1:B10, \">50\", C1:C10, \"Yes\")")
let result2 = evaluator.evaluate("=AVERAGEIFS(Sales, Region, \"West\", Date, \">=\"&DATE(2023,1,1))")
```

**Excel Documentation:** [AVERAGEIFS function](https://support.microsoft.com/en-us/office/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690)

**Implementation Status:** ✅ Full implementation

---

### BETA.DIST

Returns the beta cumulative distribution function or probability density function.

**Syntax:** `BETA.DIST(x, alpha, beta, cumulative, [A], [B])`

**Parameters:**
- `x`: Value at which to evaluate (between A and B)
- `alpha`: Shape parameter (α > 0)
- `beta`: Shape parameter (β > 0)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability density
- `A` *(optional)*: Lower bound of interval (default: 0)
- `B` *(optional)*: Upper bound of interval (default: 1)

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BETA.DIST(0.4, 2, 5, TRUE)")     // CDF ≈0.4059
let result2 = evaluator.evaluate("=BETA.DIST(0.5, 3, 3, FALSE)")    // PDF ≈1.875
let result3 = evaluator.evaluate("=BETA.DIST(7, 2, 3, TRUE, 0, 10)") // Scaled
```

**Excel Documentation:** [BETA.DIST function](https://support.microsoft.com/en-us/office/beta-dist-function-11188c9c-780a-42c7-ba43-9ecb5a878d31)

**Implementation Status:** ✅ Full implementation

**Aliases:** BETADIST (legacy)

---

### BETA.INV

Returns the inverse of the beta cumulative distribution function.

**Syntax:** `BETA.INV(probability, alpha, beta, [A], [B])`

**Parameters:**
- `probability`: Probability associated with the beta distribution (0-1)
- `alpha`: Shape parameter (α > 0)
- `beta`: Shape parameter (β > 0)
- `A` *(optional)*: Lower bound (default: 0)
- `B` *(optional)*: Upper bound (default: 1)

**Returns:** Number - Value x such that BETA.DIST(x, alpha, beta, TRUE, A, B) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=BETA.INV(0.5, 2, 3)")         // ≈0.406
let result2 = evaluator.evaluate("=BETA.INV(0.95, 3, 4, 0, 10)") // Scaled inverse
```

**Excel Documentation:** [BETA.INV function](https://support.microsoft.com/en-us/office/beta-inv-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb)

**Implementation Status:** ✅ Full implementation

**Aliases:** BETAINV (legacy)

---

### BINOM.DIST

Returns the binomial distribution probability.

**Syntax:** `BINOM.DIST(number_s, trials, probability_s, cumulative)`

**Parameters:**
- `number_s`: Number of successes
- `trials`: Number of independent trials
- `probability_s`: Probability of success on each trial (0-1)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability mass

**Returns:** Number - Probability value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BINOM.DIST(6, 10, 0.5, FALSE)") // P(X=6) ≈0.205
let result2 = evaluator.evaluate("=BINOM.DIST(6, 10, 0.5, TRUE)")  // P(X≤6) ≈0.828
let result3 = evaluator.evaluate("=BINOM.DIST(3, 5, 0.6, FALSE)")  // ≈0.3456
```

**Excel Documentation:** [BINOM.DIST function](https://support.microsoft.com/en-us/office/binom-dist-function-c5ae37b6-f39c-4be2-94c2-509a1480770c)

**Implementation Status:** ✅ Full implementation

**Aliases:** BINOMDIST (legacy)

---

### BINOM.INV

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.

**Syntax:** `BINOM.INV(trials, probability_s, alpha)`

**Parameters:**
- `trials`: Number of independent trials
- `probability_s`: Probability of success on each trial (0-1)
- `alpha`: Criterion value (0-1)

**Returns:** Number - Smallest integer x such that BINOM.DIST(x, trials, probability_s, TRUE) ≥ alpha

**Examples:**
```swift
let result1 = evaluator.evaluate("=BINOM.INV(100, 0.5, 0.95)")  // ≈58
let result2 = evaluator.evaluate("=BINOM.INV(50, 0.3, 0.5)")    // ≈15
```

**Excel Documentation:** [BINOM.INV function](https://support.microsoft.com/en-us/office/binom-inv-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9)

**Implementation Status:** ✅ Full implementation

---

### CHISQ.DIST

Returns the chi-squared distribution.

**Syntax:** `CHISQ.DIST(x, deg_freedom, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate (x ≥ 0)
- `deg_freedom`: Degrees of freedom (1-10^10)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=CHISQ.DIST(5, 3, TRUE)")   // CDF ≈0.828
let result2 = evaluator.evaluate("=CHISQ.DIST(5, 3, FALSE)")  // PDF ≈0.154
```

**Excel Documentation:** [CHISQ.DIST function](https://support.microsoft.com/en-us/office/chisq-dist-function-8486b05e-5c05-4942-a9ea-f6b341518732)

**Implementation Status:** ✅ Full implementation

---

### CHISQ.INV

Returns the inverse of the chi-squared cumulative distribution.

**Syntax:** `CHISQ.INV(probability, deg_freedom)`

**Parameters:**
- `probability`: Probability associated with chi-squared distribution (0-1)
- `deg_freedom`: Degrees of freedom (1-10^10)

**Returns:** Number - Value x such that CHISQ.DIST(x, deg_freedom, TRUE) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=CHISQ.INV(0.95, 5)")    // ≈11.07
let result2 = evaluator.evaluate("=CHISQ.INV(0.05, 10)")   // ≈3.94
```

**Excel Documentation:** [CHISQ.INV function](https://support.microsoft.com/en-us/office/chisq-inv-function-400db556-62b3-472d-80b3-254723e7092f)

**Implementation Status:** ✅ Full implementation

---

### CHISQ.TEST

Returns the test for independence using the chi-squared distribution.

**Syntax:** `CHISQ.TEST(actual_range, expected_range)`

**Parameters:**
- `actual_range`: Range of observed data
- `expected_range`: Range of expected values

**Returns:** Number - P-value for the chi-squared test

**Examples:**
```swift
let result = evaluator.evaluate("=CHISQ.TEST(A1:C2, A4:C5)")  // Test independence
```

**Excel Documentation:** [CHISQ.TEST function](https://support.microsoft.com/en-us/office/chisq-test-function-2e8a7861-b14a-4985-aa93-fb88de3f260f)

**Implementation Status:** ✅ Full implementation

**Aliases:** CHITEST (legacy)

---

### CONFIDENCE.NORM

Returns the confidence interval for a population mean using normal distribution.

**Syntax:** `CONFIDENCE.NORM(alpha, standard_dev, size)`

**Parameters:**
- `alpha`: Significance level (1 - confidence level, e.g., 0.05 for 95%)
- `standard_dev`: Population standard deviation
- `size`: Sample size

**Returns:** Number - Confidence interval margin (±)

**Examples:**
```swift
let result1 = evaluator.evaluate("=CONFIDENCE.NORM(0.05, 2.5, 50)")  // ≈0.693
let result2 = evaluator.evaluate("=CONFIDENCE.NORM(0.01, 10, 100)")  // ≈2.576
```

**Excel Documentation:** [CONFIDENCE.NORM function](https://support.microsoft.com/en-us/office/confidence-norm-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4)

**Implementation Status:** ✅ Full implementation

**Aliases:** CONFIDENCE (legacy)

---

### CONFIDENCE.T

Returns the confidence interval using Student's t-distribution.

**Syntax:** `CONFIDENCE.T(alpha, standard_dev, size)`

**Parameters:**
- `alpha`: Significance level (1 - confidence level)
- `standard_dev`: Sample standard deviation
- `size`: Sample size

**Returns:** Number - Confidence interval margin (±)

**Examples:**
```swift
let result1 = evaluator.evaluate("=CONFIDENCE.T(0.05, 3, 20)")   // ≈1.408
let result2 = evaluator.evaluate("=CONFIDENCE.T(0.01, 5, 30)")   // ≈2.462
```

**Excel Documentation:** [CONFIDENCE.T function](https://support.microsoft.com/en-us/office/confidence-t-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53)

**Implementation Status:** ✅ Full implementation

---

### CORREL

Returns the correlation coefficient between two datasets.

**Syntax:** `CORREL(array1, array2)`

**Parameters:**
- `array1`: First array or range of values
- `array2`: Second array or range of values

**Returns:** Number - Correlation coefficient (-1 to 1)

**Examples:**
```swift
let result1 = evaluator.evaluate("=CORREL(A1:A10, B1:B10)")       // Linear correlation
let result2 = evaluator.evaluate("=CORREL({1,2,3}, {2,4,6})")     // Perfect positive: 1
let result3 = evaluator.evaluate("=CORREL({1,2,3}, {3,2,1})")     // Perfect negative: -1
```

**Excel Documentation:** [CORREL function](https://support.microsoft.com/en-us/office/correl-function-995dcef7-0c0a-4bed-a3fb-239d7b68ca92)

**Implementation Status:** ✅ Full implementation

---

### COUNT

Counts the number of cells containing numbers.

**Syntax:** `COUNT(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values, ranges, or arrays (counts only numeric values)

**Returns:** Number - Count of numeric cells

**Examples:**
```swift
let result1 = evaluator.evaluate("=COUNT(A1:A10)")                // Count numbers
let result2 = evaluator.evaluate("=COUNT(1, 2, \"text\", 3)")     // 3 (text ignored)
let result3 = evaluator.evaluate("=COUNT(A:A)")                   // Count column
```

**Excel Documentation:** [COUNT function](https://support.microsoft.com/en-us/office/count-function-a59cd7fc-b623-4d93-87a4-d23bf411294c)

**Implementation Status:** ✅ Full implementation

---

### COUNTA

Counts the number of non-empty cells.

**Syntax:** `COUNTA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values, ranges, or arrays to count (counts all non-empty)

**Returns:** Number - Count of non-empty cells

**Examples:**
```swift
let result1 = evaluator.evaluate("=COUNTA(A1:A10)")              // Count all non-empty
let result2 = evaluator.evaluate("=COUNTA(1, \"text\", TRUE)")   // 3
let result3 = evaluator.evaluate("=COUNTA(A:A)")                 // Count populated cells
```

**Excel Documentation:** [COUNTA function](https://support.microsoft.com/en-us/office/counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509)

**Implementation Status:** ✅ Full implementation

---

### COUNTBLANK

Counts the number of empty cells in a range.

**Syntax:** `COUNTBLANK(range)`

**Parameters:**
- `range`: The range to check for empty cells

**Returns:** Number - Count of blank cells

**Examples:**
```swift
let result1 = evaluator.evaluate("=COUNTBLANK(A1:A10)")    // Count empty cells
let result2 = evaluator.evaluate("=COUNTBLANK(B:B)")       // Empty cells in column
```

**Excel Documentation:** [COUNTBLANK function](https://support.microsoft.com/en-us/office/countblank-function-6a92d772-675c-4bee-b346-24af6bd3ac22)

**Implementation Status:** ✅ Full implementation

---

### COUNTIF

Counts cells that meet a single criterion.

**Syntax:** `COUNTIF(range, criteria)`

**Parameters:**
- `range`: The range to evaluate
- `criteria`: The condition to test (number, text, expression like ">10")

**Returns:** Number - Count of matching cells

**Examples:**
```swift
let result1 = evaluator.evaluate("=COUNTIF(A1:A10, \">50\")")      // Count > 50
let result2 = evaluator.evaluate("=COUNTIF(B1:B10, \"Yes\")")      // Count "Yes"
let result3 = evaluator.evaluate("=COUNTIF(C:C, \"<>0\")")         // Count non-zero
```

**Excel Documentation:** [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34)

**Implementation Status:** ✅ Full implementation

---

### COUNTIFS

Counts cells that meet multiple criteria.

**Syntax:** `COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Parameters:**
- `criteria_range1, criteria1`: First condition range and criteria
- `criteria_range2, criteria2, ...`: Additional condition pairs

**Returns:** Number - Count of cells meeting all criteria

**Examples:**
```swift
let result1 = evaluator.evaluate("=COUNTIFS(A1:A10, \">50\", B1:B10, \"Yes\")")
let result2 = evaluator.evaluate("=COUNTIFS(Sales, \">1000\", Region, \"West\", Status, \"Active\")")
```

**Excel Documentation:** [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842)

**Implementation Status:** ✅ Full implementation

---

### COVARIANCE.P

Returns population covariance, the average of the products of deviations.

**Syntax:** `COVARIANCE.P(array1, array2)`

**Parameters:**
- `array1`: First array or range
- `array2`: Second array or range

**Returns:** Number - Population covariance

**Examples:**
```swift
let result1 = evaluator.evaluate("=COVARIANCE.P(A1:A10, B1:B10)")
let result2 = evaluator.evaluate("=COVARIANCE.P({1,2,3}, {2,4,5})")  // ≈1.556
```

**Excel Documentation:** [COVARIANCE.P function](https://support.microsoft.com/en-us/office/covariance-p-function-6f0e1e6d-956d-4e4b-9943-cfef0bf9edfc)

**Implementation Status:** ✅ Full implementation

**Aliases:** COVAR (legacy)

---

### COVARIANCE.S

Returns sample covariance.

**Syntax:** `COVARIANCE.S(array1, array2)`

**Parameters:**
- `array1`: First sample array or range
- `array2`: Second sample array or range

**Returns:** Number - Sample covariance

**Examples:**
```swift
let result = evaluator.evaluate("=COVARIANCE.S(A1:A10, B1:B10)")  // Sample covariance
```

**Excel Documentation:** [COVARIANCE.S function](https://support.microsoft.com/en-us/office/covariance-s-function-0a539b74-7371-42aa-a18f-1f5320314977)

**Implementation Status:** ✅ Full implementation

---

### CRITBINOM

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value (legacy function).

**Syntax:** `CRITBINOM(trials, probability_s, alpha)`

**Parameters:**
- `trials`: Number of Bernoulli trials
- `probability_s`: Probability of success
- `alpha`: Criterion value

**Returns:** Number - Critical value

**Examples:**
```swift
let result = evaluator.evaluate("=CRITBINOM(100, 0.5, 0.95)")  // ≈58
```

**Excel Documentation:** [CRITBINOM function](https://support.microsoft.com/en-us/office/critbinom-function-eb6b871d-796b-4d21-b69b-e4350d5f407b)

**Implementation Status:** ✅ Full implementation

**Note:** Consider using BINOM.INV for new workbooks

---

### DEVSQ

Returns the sum of squares of deviations from the sample mean.

**Syntax:** `DEVSQ(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges

**Returns:** Number - Sum of squared deviations

**Examples:**
```swift
let result1 = evaluator.evaluate("=DEVSQ(4, 5, 6, 7, 5, 4, 3)")  // Sum of (x-mean)²
let result2 = evaluator.evaluate("=DEVSQ(A1:A10)")               // Variability measure
```

**Excel Documentation:** [DEVSQ function](https://support.microsoft.com/en-us/office/devsq-function-8b739616-8376-4df5-8bd0-cfe0a6caf444)

**Implementation Status:** ✅ Full implementation

---

### DSTDEV

Estimates the standard deviation of a population based on a sample using database criteria.

**Syntax:** `DSTDEV(database, field, criteria)`

**Parameters:**
- `database`: The range containing the database
- `field`: Column to aggregate (name or index)
- `criteria`: Range containing conditions

**Returns:** Number - Sample standard deviation

**Examples:**
```swift
let result = evaluator.evaluate("=DSTDEV(A1:E100, \"Salary\", G1:G2)")
```

**Excel Documentation:** [DSTDEV function](https://support.microsoft.com/en-us/office/dstdev-function-026b8c73-616d-4b5e-b072-241871c4ab96)

**Implementation Status:** ✅ Full implementation

---

### DVAR

Estimates the variance of a population based on a sample using database criteria.

**Syntax:** `DVAR(database, field, criteria)`

**Parameters:**
- `database`: The range containing the database
- `field`: Column to aggregate
- `criteria`: Range containing conditions

**Returns:** Number - Sample variance

**Examples:**
```swift
let result = evaluator.evaluate("=DVAR(A1:E100, \"Sales\", G1:G2)")
```

**Excel Documentation:** [DVAR function](https://support.microsoft.com/en-us/office/dvar-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1)

**Implementation Status:** ✅ Full implementation

---

### EXPON.DIST

Returns the exponential distribution.

**Syntax:** `EXPON.DIST(x, lambda, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate (x ≥ 0)
- `lambda`: Parameter value (λ > 0)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=EXPON.DIST(1, 1, TRUE)")   // CDF ≈0.632
let result2 = evaluator.evaluate("=EXPON.DIST(1, 1, FALSE)")  // PDF ≈0.368
let result3 = evaluator.evaluate("=EXPON.DIST(0.5, 2, TRUE)") // ≈0.632
```

**Excel Documentation:** [EXPON.DIST function](https://support.microsoft.com/en-us/office/expon-dist-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e)

**Implementation Status:** ✅ Full implementation

**Aliases:** EXPONDIST (legacy)

---

### F.DIST

Returns the F probability distribution.

**Syntax:** `F.DIST(x, deg_freedom1, deg_freedom2, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate (x ≥ 0)
- `deg_freedom1`: Numerator degrees of freedom
- `deg_freedom2`: Denominator degrees of freedom
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=F.DIST(15.2, 6, 4, TRUE)")   // CDF
let result2 = evaluator.evaluate("=F.DIST(2, 5, 10, FALSE)")    // PDF
```

**Excel Documentation:** [F.DIST function](https://support.microsoft.com/en-us/office/f-dist-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d)

**Implementation Status:** ✅ Full implementation

---

### F.INV

Returns the inverse of the F cumulative distribution.

**Syntax:** `F.INV(probability, deg_freedom1, deg_freedom2)`

**Parameters:**
- `probability`: Probability associated with F distribution (0-1)
- `deg_freedom1`: Numerator degrees of freedom
- `deg_freedom2`: Denominator degrees of freedom

**Returns:** Number - Inverse value

**Examples:**
```swift
let result1 = evaluator.evaluate("=F.INV(0.95, 6, 4)")    // Critical value
let result2 = evaluator.evaluate("=F.INV(0.01, 10, 20)")  // ≈0.286
```

**Excel Documentation:** [F.INV function](https://support.microsoft.com/en-us/office/f-inv-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe)

**Implementation Status:** ✅ Full implementation

---

### F.TEST

Returns the result of an F-test for comparing variances.

**Syntax:** `F.TEST(array1, array2)`

**Parameters:**
- `array1`: First dataset
- `array2`: Second dataset

**Returns:** Number - Two-tailed p-value

**Examples:**
```swift
let result = evaluator.evaluate("=F.TEST(A1:A10, B1:B10)")  // Variance comparison
```

**Excel Documentation:** [F.TEST function](https://support.microsoft.com/en-us/office/f-test-function-100a59e7-4108-46f8-8443-78ffacb6c0a7)

**Implementation Status:** ✅ Full implementation

**Aliases:** FTEST (legacy)

---

### FISHER

Returns the Fisher transformation.

**Syntax:** `FISHER(x)`

**Parameters:**
- `x`: Correlation coefficient (-1 < x < 1)

**Returns:** Number - Fisher transformation value

**Examples:**
```swift
let result1 = evaluator.evaluate("=FISHER(0.5)")     // ≈0.549
let result2 = evaluator.evaluate("=FISHER(0.75)")    // ≈0.973
let result3 = evaluator.evaluate("=FISHER(-0.5)")    // ≈-0.549
```

**Excel Documentation:** [FISHER function](https://support.microsoft.com/en-us/office/fisher-function-d656523c-5076-4f95-b87b-7741bf236c69)

**Implementation Status:** ✅ Full implementation

---

### FISHERINV

Returns the inverse of the Fisher transformation.

**Syntax:** `FISHERINV(y)`

**Parameters:**
- `y`: Value for which to compute the inverse

**Returns:** Number - Inverse Fisher transformation

**Examples:**
```swift
let result1 = evaluator.evaluate("=FISHERINV(0.549)")   // ≈0.5
let result2 = evaluator.evaluate("=FISHERINV(0.973)")   // ≈0.75
```

**Excel Documentation:** [FISHERINV function](https://support.microsoft.com/en-us/office/fisherinv-function-62504b39-415a-4284-a285-19c8e82f86bb)

**Implementation Status:** ✅ Full implementation

---

### FORECAST

Predicts a value based on linear regression (legacy function name).

**Syntax:** `FORECAST(x, known_y's, known_x's)`

**Parameters:**
- `x`: Data point for which to predict a value
- `known_y's`: Dependent array or range
- `known_x's`: Independent array or range

**Returns:** Number - Predicted value

**Examples:**
```swift
let result1 = evaluator.evaluate("=FORECAST(30, A1:A10, B1:B10)")  // Predict y at x=30
let result2 = evaluator.evaluate("=FORECAST(5, {10,20,30}, {1,2,3})") // Linear prediction
```

**Excel Documentation:** [FORECAST function](https://support.microsoft.com/en-us/office/forecast-function-50ca49c9-7b40-4892-94e4-7ad38bbeda99)

**Implementation Status:** ✅ Full implementation

**Aliases:** FORECAST.LINEAR (current name)

---

### FORECAST.ETS

Returns a future value based on exponential smoothing.

**Syntax:** `FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])`

**Parameters:**
- `target_date`: Date for which to predict
- `values`: Historical values
- `timeline`: Date array corresponding to values
- `seasonality` *(optional)*: Length of seasonal pattern (0=auto)
- `data_completion` *(optional)*: How to handle missing points
- `aggregation` *(optional)*: Method for aggregating multiple values

**Returns:** Number - Forecasted value

**Examples:**
```swift
let result = evaluator.evaluate("=FORECAST.ETS(DATE(2024,1,1), A1:A100, B1:B100)")
```

**Excel Documentation:** [FORECAST.ETS function](https://support.microsoft.com/en-us/office/forecast-ets-function-15389b8b-677e-4fbd-bd95-21d464333f41)

**Implementation Status:** ✅ Full implementation

---

### FORECAST.ETS.CONFINT

Returns a confidence interval for the forecast value.

**Syntax:** `FORECAST.ETS.CONFINT(target_date, values, timeline, [confidence_level], [seasonality], [data_completion], [aggregation])`

**Parameters:**
- `target_date`: Date for prediction
- `values`: Historical values
- `timeline`: Date array
- `confidence_level` *(optional)*: Confidence level (default: 0.95)
- Additional parameters as in FORECAST.ETS

**Returns:** Number - Confidence interval

**Examples:**
```swift
let result = evaluator.evaluate("=FORECAST.ETS.CONFINT(DATE(2024,6,1), A1:A50, B1:B50, 0.95)")
```

**Excel Documentation:** [FORECAST.ETS.CONFINT function](https://support.microsoft.com/en-us/office/forecast-ets-confint-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)

**Implementation Status:** ✅ Full implementation

---

### FORECAST.ETS.SEASONALITY

Returns the detected seasonality length.

**Syntax:** `FORECAST.ETS.SEASONALITY(values, timeline, [data_completion], [aggregation])`

**Parameters:**
- `values`: Historical values
- `timeline`: Date array
- `data_completion` *(optional)*: Missing point handling
- `aggregation` *(optional)*: Aggregation method

**Returns:** Number - Detected seasonal cycle length

**Examples:**
```swift
let result = evaluator.evaluate("=FORECAST.ETS.SEASONALITY(A1:A100, B1:B100)")
```

**Excel Documentation:** [FORECAST.ETS.SEASONALITY function](https://support.microsoft.com/en-us/office/forecast-ets-seasonality-function-30db8cfb-2cdb-4cf2-94d4-e02e5698c28c)

**Implementation Status:** ✅ Full implementation

---

### FORECAST.ETS.STAT

Returns a statistical value for the forecast.

**Syntax:** `FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality], [data_completion], [aggregation])`

**Parameters:**
- `values`: Historical values
- `timeline`: Date array
- `statistic_type`: Type of statistic (1-8)
  - 1: Alpha, 2: Beta, 3: Gamma
  - 4: MASE, 5: SMAPE, 6: MAE, 7: RMSE, 8: Step size
- Additional parameters as in FORECAST.ETS

**Returns:** Number - Requested statistic

**Examples:**
```swift
let result = evaluator.evaluate("=FORECAST.ETS.STAT(A1:A100, B1:B100, 1)")  // Get Alpha
```

**Excel Documentation:** [FORECAST.ETS.STAT function](https://support.microsoft.com/en-us/office/forecast-ets-stat-function-eed9cd2e-0564-4dbb-a81f-f6d1e7a73152)

**Implementation Status:** ✅ Full implementation

---

### FREQUENCY

Returns a frequency distribution as a vertical array.

**Syntax:** `FREQUENCY(data_array, bins_array)`

**Parameters:**
- `data_array`: Array of values to count
- `bins_array`: Array of intervals into which to group values

**Returns:** Array - Frequency counts for each bin

**Examples:**
```swift
let result1 = evaluator.evaluate("=FREQUENCY(A1:A20, {10,20,30,40})")  // Count by bins
let result2 = evaluator.evaluate("=FREQUENCY({1,2,3,4,5,6}, {2,4,6})") // {2,2,2,0}
```

**Excel Documentation:** [FREQUENCY function](https://support.microsoft.com/en-us/office/frequency-function-44e3be2b-eca0-42cd-a3f7-fd9ea898fdb9)

**Implementation Status:** ✅ Full implementation

---

### GAMMA

Returns the Gamma function value.

**Syntax:** `GAMMA(number)`

**Parameters:**
- `number`: Value at which to evaluate Gamma function

**Returns:** Number - Γ(number)

**Examples:**
```swift
let result1 = evaluator.evaluate("=GAMMA(5)")      // 24 (same as 4!)
let result2 = evaluator.evaluate("=GAMMA(0.5)")    // √π ≈1.772
let result3 = evaluator.evaluate("=GAMMA(3.5)")    // ≈3.323
```

**Excel Documentation:** [GAMMA function](https://support.microsoft.com/en-us/office/gamma-function-ce1702b1-cf55-471d-8307-f83be0fc5297)

**Implementation Status:** ✅ Full implementation

---

### GAMMA.DIST

Returns the gamma distribution.

**Syntax:** `GAMMA.DIST(x, alpha, beta, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate (x ≥ 0)
- `alpha`: Shape parameter (α > 0)
- `beta`: Scale parameter (β > 0)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=GAMMA.DIST(5, 2, 1, TRUE)")   // CDF
let result2 = evaluator.evaluate("=GAMMA.DIST(5, 2, 1, FALSE)")  // PDF
```

**Excel Documentation:** [GAMMA.DIST function](https://support.microsoft.com/en-us/office/gamma-dist-function-9b6f1538-d11c-4d5f-8966-21f6a2201def)

**Implementation Status:** ✅ Full implementation

**Aliases:** GAMMADIST (legacy)

---

### GAMMA.INV

Returns the inverse of the gamma cumulative distribution.

**Syntax:** `GAMMA.INV(probability, alpha, beta)`

**Parameters:**
- `probability`: Probability associated with gamma distribution (0-1)
- `alpha`: Shape parameter
- `beta`: Scale parameter

**Returns:** Number - Inverse value

**Examples:**
```swift
let result = evaluator.evaluate("=GAMMA.INV(0.95, 2, 3)")  // Critical value
```

**Excel Documentation:** [GAMMA.INV function](https://support.microsoft.com/en-us/office/gamma-inv-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18)

**Implementation Status:** ✅ Full implementation

**Aliases:** GAMMAINV (legacy)

---

### GAMMALN

Returns the natural logarithm of the gamma function.

**Syntax:** `GAMMALN(x)`

**Parameters:**
- `x`: Value at which to evaluate (x > 0)

**Returns:** Number - ln(Γ(x))

**Examples:**
```swift
let result1 = evaluator.evaluate("=GAMMALN(5)")     // ln(24) ≈3.178
let result2 = evaluator.evaluate("=GAMMALN(10)")    // ln(362880) ≈12.802
```

**Excel Documentation:** [GAMMALN function](https://support.microsoft.com/en-us/office/gammaln-function-b838c48b-c65f-484f-9e1d-141c55470eb9)

**Implementation Status:** ✅ Full implementation

**Aliases:** GAMMALN.PRECISE

---

### GEOMEAN

Returns the geometric mean of positive numbers.

**Syntax:** `GEOMEAN(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Positive numbers or ranges

**Returns:** Number - Geometric mean (nth root of product)

**Examples:**
```swift
let result1 = evaluator.evaluate("=GEOMEAN(4, 9)")              // 6
let result2 = evaluator.evaluate("=GEOMEAN(2, 8)")              // 4
let result3 = evaluator.evaluate("=GEOMEAN(A1:A10)")            // Growth rate analysis
```

**Excel Documentation:** [GEOMEAN function](https://support.microsoft.com/en-us/office/geomean-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5)

**Implementation Status:** ✅ Full implementation

---

### GROWTH

Returns values along an exponential trend.

**Syntax:** `GROWTH(known_y's, [known_x's], [new_x's], [const])`

**Parameters:**
- `known_y's`: Known y-values (y = b*m^x)
- `known_x's` *(optional)*: Known x-values (default: 1, 2, 3, ...)
- `new_x's` *(optional)*: New x-values for prediction (default: known_x's)
- `const` *(optional)*: TRUE to calculate b, FALSE to set b=1

**Returns:** Array - Predicted exponential growth values

**Examples:**
```swift
let result1 = evaluator.evaluate("=GROWTH(A1:A10, B1:B10, B11:B15)")  // Exponential forecast
let result2 = evaluator.evaluate("=GROWTH({100,200,400}, {1,2,3})")   // Geometric progression
```

**Excel Documentation:** [GROWTH function](https://support.microsoft.com/en-us/office/growth-function-541a91dc-3d5e-437d-b156-21324e68b80d)

**Implementation Status:** ✅ Full implementation

---

### HARMEAN

Returns the harmonic mean of positive numbers.

**Syntax:** `HARMEAN(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Positive numbers or ranges

**Returns:** Number - Harmonic mean (n / Σ(1/x))

**Examples:**
```swift
let result1 = evaluator.evaluate("=HARMEAN(4, 5, 8, 7, 11, 4, 3)")  // ≈5.028
let result2 = evaluator.evaluate("=HARMEAN(10, 20)")                // ≈13.33
```

**Excel Documentation:** [HARMEAN function](https://support.microsoft.com/en-us/office/harmean-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6)

**Implementation Status:** ✅ Full implementation

---

### HYPGEOM.DIST

Returns the hypergeometric distribution.

**Syntax:** `HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)`

**Parameters:**
- `sample_s`: Number of successes in sample
- `number_sample`: Sample size
- `population_s`: Number of successes in population
- `number_pop`: Population size
- `cumulative`: TRUE for cumulative, FALSE for probability mass

**Returns:** Number - Probability value

**Examples:**
```swift
let result1 = evaluator.evaluate("=HYPGEOM.DIST(1, 4, 8, 20, FALSE)")  // PMF
let result2 = evaluator.evaluate("=HYPGEOM.DIST(1, 4, 8, 20, TRUE)")   // CDF
```

**Excel Documentation:** [HYPGEOM.DIST function](https://support.microsoft.com/en-us/office/hypgeom-dist-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf)

**Implementation Status:** ✅ Full implementation

**Aliases:** HYPGEOMDIST (legacy)

---

### INTERCEPT

Returns the y-intercept of linear regression line.

**Syntax:** `INTERCEPT(known_y's, known_x's)`

**Parameters:**
- `known_y's`: Dependent values
- `known_x's`: Independent values

**Returns:** Number - Y-intercept (b in y = mx + b)

**Examples:**
```swift
let result1 = evaluator.evaluate("=INTERCEPT(A1:A10, B1:B10)")  // Find b
let result2 = evaluator.evaluate("=INTERCEPT({1,9,5,7}, {0,4,2,3})")  // b ≈0.048
```

**Excel Documentation:** [INTERCEPT function](https://support.microsoft.com/en-us/office/intercept-function-2a9b74e2-9d47-4772-b663-3bca70bf63ef)

**Implementation Status:** ✅ Full implementation

---

### KURT

Returns the kurtosis of a dataset (measure of tailedness).

**Syntax:** `KURT(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges (requires ≥4 data points)

**Returns:** Number - Kurtosis coefficient

**Examples:**
```swift
let result1 = evaluator.evaluate("=KURT(3,4,5,2,3,4,5,6,4,7)")  // Tail analysis
let result2 = evaluator.evaluate("=KURT(A1:A100)")               // Distribution shape
```

**Excel Documentation:** [KURT function](https://support.microsoft.com/en-us/office/kurt-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab)

**Implementation Status:** ✅ Full implementation

---

### LARGE

Returns the k-th largest value in a dataset.

**Syntax:** `LARGE(array, k)`

**Parameters:**
- `array`: Array or range of values
- `k`: Position from the largest (1 = largest, 2 = 2nd largest, etc.)

**Returns:** Number - The k-th largest value

**Examples:**
```swift
let result1 = evaluator.evaluate("=LARGE(A1:A10, 1)")           // Largest value
let result2 = evaluator.evaluate("=LARGE(A1:A10, 3)")           // 3rd largest
let result3 = evaluator.evaluate("=LARGE({3,7,2,9,1}, 2)")      // 7
```

**Excel Documentation:** [LARGE function](https://support.microsoft.com/en-us/office/large-function-3af0af19-1190-42bb-bb8b-01672ec00a64)

**Implementation Status:** ✅ Full implementation

---

### LINEST

Returns statistics for a linear trend using least squares method.

**Syntax:** `LINEST(known_y's, [known_x's], [const], [stats])`

**Parameters:**
- `known_y's`: Known y-values
- `known_x's` *(optional)*: Known x-values (default: 1, 2, 3, ...)
- `const` *(optional)*: TRUE to calculate intercept (default: TRUE)
- `stats` *(optional)*: TRUE to return additional statistics

**Returns:** Array - Regression coefficients and optionally statistics

**Examples:**
```swift
let result1 = evaluator.evaluate("=LINEST(A1:A10, B1:B10)")        // Slope and intercept
let result2 = evaluator.evaluate("=LINEST(Y, X, TRUE, TRUE)")      // Full statistics
```

**Excel Documentation:** [LINEST function](https://support.microsoft.com/en-us/office/linest-function-84d7d0d9-6e50-4101-977a-fa7abf772b6d)

**Implementation Status:** ✅ Full implementation

---

### LOGEST

Returns statistics for an exponential curve using least squares.

**Syntax:** `LOGEST(known_y's, [known_x's], [const], [stats])`

**Parameters:**
- `known_y's`: Known y-values (for y = b*m^x)
- `known_x's` *(optional)*: Known x-values
- `const` *(optional)*: TRUE to calculate b
- `stats` *(optional)*: TRUE to return additional statistics

**Returns:** Array - Exponential regression coefficients and statistics

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOGEST(A1:A10, B1:B10)")        // Exponential fit
let result2 = evaluator.evaluate("=LOGEST(Y, X, TRUE, TRUE)")      // Full statistics
```

**Excel Documentation:** [LOGEST function](https://support.microsoft.com/en-us/office/logest-function-f27462d8-3657-4030-866b-a272c1d18b4b)

**Implementation Status:** ✅ Full implementation

---

### LOGNORM.DIST

Returns the lognormal distribution.

**Syntax:** `LOGNORM.DIST(x, mean, standard_dev, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate (x > 0)
- `mean`: Mean of ln(x)
- `standard_dev`: Standard deviation of ln(x)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOGNORM.DIST(4, 3.5, 1.2, TRUE)")   // CDF
let result2 = evaluator.evaluate("=LOGNORM.DIST(4, 3.5, 1.2, FALSE)")  // PDF
```

**Excel Documentation:** [LOGNORM.DIST function](https://support.microsoft.com/en-us/office/lognorm-dist-function-eb60d00b-48a9-4217-be2b-6074aee6b070)

**Implementation Status:** ✅ Full implementation

**Aliases:** LOGNORMDIST (legacy)

---

### LOGNORM.INV

Returns the inverse of the lognormal cumulative distribution.

**Syntax:** `LOGNORM.INV(probability, mean, standard_dev)`

**Parameters:**
- `probability`: Probability associated with lognormal distribution (0-1)
- `mean`: Mean of ln(x)
- `standard_dev`: Standard deviation of ln(x)

**Returns:** Number - Inverse value

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOGNORM.INV(0.95, 3, 1)")   // Critical value
let result2 = evaluator.evaluate("=LOGNORM.INV(0.5, 0, 1)")    // Median
```

**Excel Documentation:** [LOGNORM.INV function](https://support.microsoft.com/en-us/office/lognorm-inv-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600)

**Implementation Status:** ✅ Full implementation

**Aliases:** LOGNORMINV, LOGINV (legacy)

---

### MAX

Returns the largest value in a set of numbers.

**Syntax:** `MAX(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers, ranges, or arrays (ignores text and logical)

**Returns:** Number - Maximum value

**Examples:**
```swift
let result1 = evaluator.evaluate("=MAX(10, 20, 5, 30)")      // 30
let result2 = evaluator.evaluate("=MAX(A1:A10)")             // Largest in range
let result3 = evaluator.evaluate("=MAX(A:A)")                // Largest in column
```

**Excel Documentation:** [MAX function](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098)

**Implementation Status:** ✅ Full implementation

---

### MAXA

Returns the largest value including text and logical values.

**Syntax:** `MAXA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values (text=0, TRUE=1, FALSE=0)

**Returns:** Number - Maximum value

**Examples:**
```swift
let result = evaluator.evaluate("=MAXA(5, TRUE, \"text\", 3)")  // 5
```

**Excel Documentation:** [MAXA function](https://support.microsoft.com/en-us/office/maxa-function-814bda1e-3840-4bff-9365-2f59ac2ee62d)

**Implementation Status:** ✅ Full implementation

---

### MAXIFS

Returns the maximum value among cells specified by multiple criteria.

**Syntax:** `MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Parameters:**
- `max_range`: Range to find maximum from
- `criteria_range1, criteria1`: First condition range and criteria
- Additional condition pairs

**Returns:** Number - Maximum of cells meeting all criteria

**Examples:**
```swift
let result1 = evaluator.evaluate("=MAXIFS(A1:A10, B1:B10, \">50\", C1:C10, \"Yes\")")
let result2 = evaluator.evaluate("=MAXIFS(Sales, Region, \"West\", Year, 2023)")
```

**Excel Documentation:** [MAXIFS function](https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883)

**Implementation Status:** ✅ Full implementation

---

### MEDIAN

Returns the median (middle value) of a set of numbers.

**Syntax:** `MEDIAN(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges

**Returns:** Number - The median value

**Examples:**
```swift
let result1 = evaluator.evaluate("=MEDIAN(1, 2, 3, 4, 5)")      // 3
let result2 = evaluator.evaluate("=MEDIAN(1, 2, 3, 4)")         // 2.5 (average of middle)
let result3 = evaluator.evaluate("=MEDIAN(A1:A100)")            // 50th percentile
```

**Excel Documentation:** [MEDIAN function](https://support.microsoft.com/en-us/office/median-function-d0916313-4753-414c-8537-ce85bdd967d2)

**Implementation Status:** ✅ Full implementation

---

### MIN

Returns the smallest value in a set of numbers.

**Syntax:** `MIN(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers, ranges, or arrays (ignores text and logical)

**Returns:** Number - Minimum value

**Examples:**
```swift
let result1 = evaluator.evaluate("=MIN(10, 20, 5, 30)")      // 5
let result2 = evaluator.evaluate("=MIN(A1:A10)")             // Smallest in range
let result3 = evaluator.evaluate("=MIN(A:A)")                // Smallest in column
```

**Excel Documentation:** [MIN function](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152)

**Implementation Status:** ✅ Full implementation

---

### MINA

Returns the smallest value including text and logical values.

**Syntax:** `MINA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values (text=0, TRUE=1, FALSE=0)

**Returns:** Number - Minimum value

**Examples:**
```swift
let result = evaluator.evaluate("=MINA(5, TRUE, \"text\", 3)")  // 0 (text=0)
```

**Excel Documentation:** [MINA function](https://support.microsoft.com/en-us/office/mina-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3)

**Implementation Status:** ✅ Full implementation

---

### MINIFS

Returns the minimum value among cells specified by multiple criteria.

**Syntax:** `MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Parameters:**
- `min_range`: Range to find minimum from
- `criteria_range1, criteria1`: First condition range and criteria
- Additional condition pairs

**Returns:** Number - Minimum of cells meeting all criteria

**Examples:**
```swift
let result1 = evaluator.evaluate("=MINIFS(A1:A10, B1:B10, \">50\", C1:C10, \"Yes\")")
let result2 = evaluator.evaluate("=MINIFS(Prices, Category, \"Electronics\", Stock, \">0\")")
```

**Excel Documentation:** [MINIFS function](https://support.microsoft.com/en-us/office/minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599)

**Implementation Status:** ✅ Full implementation

---

### MODE

Returns the most frequently occurring value in a dataset.

**Syntax:** `MODE(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges (requires at least 2 values)

**Returns:** Number - Most common value

**Examples:**
```swift
let result1 = evaluator.evaluate("=MODE(1, 2, 3, 3, 4, 5)")      // 3
let result2 = evaluator.evaluate("=MODE(A1:A100)")                // Most frequent value
```

**Excel Documentation:** [MODE function](https://support.microsoft.com/en-us/office/mode-function-e45192ce-9122-4980-82ed-4bdc34973120)

**Implementation Status:** ✅ Full implementation

**Aliases:** MODE.SNGL (current name)

---

### MODE.MULT

Returns a vertical array of the most frequently occurring values.

**Syntax:** `MODE.MULT(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges

**Returns:** Array - Multiple mode values

**Examples:**
```swift
let result = evaluator.evaluate("=MODE.MULT({1,2,2,3,3,4})")  // {2,3}
```

**Excel Documentation:** [MODE.MULT function](https://support.microsoft.com/en-us/office/mode-mult-function-50fd9464-b2ba-4191-b57a-39446689ae8c)

**Implementation Status:** ✅ Full implementation

---

### NEGBINOM.DIST

Returns the negative binomial distribution.

**Syntax:** `NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)`

**Parameters:**
- `number_f`: Number of failures
- `number_s`: Threshold number of successes
- `probability_s`: Probability of success
- `cumulative`: TRUE for cumulative, FALSE for probability mass

**Returns:** Number - Probability value

**Examples:**
```swift
let result1 = evaluator.evaluate("=NEGBINOM.DIST(10, 5, 0.25, FALSE)")  // PMF
let result2 = evaluator.evaluate("=NEGBINOM.DIST(10, 5, 0.25, TRUE)")   // CDF
```

**Excel Documentation:** [NEGBINOM.DIST function](https://support.microsoft.com/en-us/office/negbinom-dist-function-c8239f89-c2d0-45bd-b6af-172e570f8599)

**Implementation Status:** ✅ Full implementation

**Aliases:** NEGBINOMDIST (legacy)

---

### NORM.DIST

Returns the normal distribution.

**Syntax:** `NORM.DIST(x, mean, standard_dev, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate
- `mean`: Arithmetic mean of the distribution
- `standard_dev`: Standard deviation (> 0)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORM.DIST(42, 40, 1.5, TRUE)")   // CDF ≈0.908
let result2 = evaluator.evaluate("=NORM.DIST(42, 40, 1.5, FALSE)")  // PDF ≈0.109
let result3 = evaluator.evaluate("=NORM.DIST(50, 50, 10, TRUE)")    // 0.5 (at mean)
```

**Excel Documentation:** [NORM.DIST function](https://support.microsoft.com/en-us/office/norm-dist-function-edb1cc14-a21c-4e53-839d-8082074c9f8d)

**Implementation Status:** ✅ Full implementation

**Aliases:** NORMDIST (legacy)

---

### NORM.INV

Returns the inverse of the normal cumulative distribution.

**Syntax:** `NORM.INV(probability, mean, standard_dev)`

**Parameters:**
- `probability`: Probability corresponding to normal distribution (0-1)
- `mean`: Arithmetic mean
- `standard_dev`: Standard deviation

**Returns:** Number - Value x such that NORM.DIST(x, mean, standard_dev, TRUE) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORM.INV(0.95, 100, 15)")  // ≈124.67
let result2 = evaluator.evaluate("=NORM.INV(0.5, 50, 10)")    // 50 (median=mean)
```

**Excel Documentation:** [NORM.INV function](https://support.microsoft.com/en-us/office/norm-inv-function-54b30935-fee7-493c-bedb-2278a9db7e13)

**Implementation Status:** ✅ Full implementation

**Aliases:** NORMINV (legacy)

---

### NORM.S.DIST

Returns the standard normal distribution (mean=0, stdev=1).

**Syntax:** `NORM.S.DIST(z, cumulative)`

**Parameters:**
- `z`: Z-score (standardized value)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORM.S.DIST(1.96, TRUE)")   // ≈0.975 (95th percentile)
let result2 = evaluator.evaluate("=NORM.S.DIST(0, TRUE)")      // 0.5
let result3 = evaluator.evaluate("=NORM.S.DIST(1, FALSE)")     // ≈0.242 (PDF at z=1)
```

**Excel Documentation:** [NORM.S.DIST function](https://support.microsoft.com/en-us/office/norm-s-dist-function-1e787282-3832-4520-a9ae-bd2a8d99ba88)

**Implementation Status:** ✅ Full implementation

**Aliases:** NORMSDIST (legacy)

---

### NORM.S.INV

Returns the inverse of the standard normal cumulative distribution.

**Syntax:** `NORM.S.INV(probability)`

**Parameters:**
- `probability`: Probability corresponding to standard normal (0-1)

**Returns:** Number - Z-score such that NORM.S.DIST(z, TRUE) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORM.S.INV(0.975)")   // ≈1.96 (95% CI)
let result2 = evaluator.evaluate("=NORM.S.INV(0.5)")     // 0 (median)
let result3 = evaluator.evaluate("=NORM.S.INV(0.025)")   // ≈-1.96
```

**Excel Documentation:** [NORM.S.INV function](https://support.microsoft.com/en-us/office/norm-s-inv-function-d6d556b4-ab7f-49cd-b526-5a20918452b1)

**Implementation Status:** ✅ Full implementation

**Aliases:** NORMSINV (legacy)

---

### PEARSON

Returns the Pearson correlation coefficient (same as CORREL).

**Syntax:** `PEARSON(array1, array2)`

**Parameters:**
- `array1`: First array or range
- `array2`: Second array or range

**Returns:** Number - Pearson's r (-1 to 1)

**Examples:**
```swift
let result1 = evaluator.evaluate("=PEARSON(A1:A10, B1:B10)")  // Correlation
let result2 = evaluator.evaluate("=PEARSON({1,2,3}, {2,4,6})") // 1 (perfect positive)
```

**Excel Documentation:** [PEARSON function](https://support.microsoft.com/en-us/office/pearson-function-0c3e30fc-e5af-49c4-808a-3ef66e034c18)

**Implementation Status:** ✅ Full implementation

---

### PERCENTILE

Returns the k-th percentile of values in a range (inclusive method).

**Syntax:** `PERCENTILE(array, k)`

**Parameters:**
- `array`: Array or range of data
- `k`: Percentile value (0-1)

**Returns:** Number - The k-th percentile

**Examples:**
```swift
let result1 = evaluator.evaluate("=PERCENTILE(A1:A100, 0.5)")   // Median (50th percentile)
let result2 = evaluator.evaluate("=PERCENTILE(A1:A100, 0.95)")  // 95th percentile
let result3 = evaluator.evaluate("=PERCENTILE({1,2,3,4}, 0.3)") // 1.9
```

**Excel Documentation:** [PERCENTILE function](https://support.microsoft.com/en-us/office/percentile-function-91b43a53-543c-4708-93de-d626debdddca)

**Implementation Status:** ✅ Full implementation

**Aliases:** PERCENTILE.INC (current name)

---

### PERCENTILE.EXC

Returns the k-th percentile using exclusive method.

**Syntax:** `PERCENTILE.EXC(array, k)`

**Parameters:**
- `array`: Array or range of data
- `k`: Percentile value (must be between 0 and 1, exclusive: 1/(n+1) < k < n/(n+1))

**Returns:** Number - The k-th percentile

**Examples:**
```swift
let result1 = evaluator.evaluate("=PERCENTILE.EXC(A1:A100, 0.5)")  // Median
let result2 = evaluator.evaluate("=PERCENTILE.EXC({1,2,3,4}, 0.4)") // Exclusive method
```

**Excel Documentation:** [PERCENTILE.EXC function](https://support.microsoft.com/en-us/office/percentile-exc-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba)

**Implementation Status:** ✅ Full implementation

---

### PERCENTRANK.INC

Returns the rank of a value as a percentage (inclusive, 0 to 1).

**Syntax:** `PERCENTRANK.INC(array, x, [significance])`

**Parameters:**
- `array`: Array or range of data
- `x`: Value for which to find the rank
- `significance` *(optional)*: Number of significant digits (default: 3)

**Returns:** Number - Percentage rank (0-1)

**Examples:**
```swift
let result1 = evaluator.evaluate("=PERCENTRANK.INC(A1:A100, 75)")      // Rank of 75
let result2 = evaluator.evaluate("=PERCENTRANK.INC({1,2,3,4,5}, 3)")  // 0.5
```

**Excel Documentation:** [PERCENTRANK.INC function](https://support.microsoft.com/en-us/office/percentrank-inc-function-149592c9-00c0-49ba-86c1-c1f45b80463a)

**Implementation Status:** ✅ Full implementation

**Aliases:** PERCENTRANK (legacy)

---

### PERCENTRANK.EXC

Returns the rank of a value as a percentage (exclusive method).

**Syntax:** `PERCENTRANK.EXC(array, x, [significance])`

**Parameters:**
- `array`: Array or range of data
- `x`: Value to rank
- `significance` *(optional)*: Number of significant digits

**Returns:** Number - Percentage rank

**Examples:**
```swift
let result = evaluator.evaluate("=PERCENTRANK.EXC(A1:A100, 75, 4)")  // Exclusive rank
```

**Excel Documentation:** [PERCENTRANK.EXC function](https://support.microsoft.com/en-us/office/percentrank-exc-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314)

**Implementation Status:** ✅ Full implementation

---

### POISSON.DIST

Returns the Poisson distribution.

**Syntax:** `POISSON.DIST(x, mean, cumulative)`

**Parameters:**
- `x`: Number of events (integer ≥ 0)
- `mean`: Expected number of events (λ > 0)
- `cumulative`: TRUE for cumulative, FALSE for probability mass

**Returns:** Number - Probability value

**Examples:**
```swift
let result1 = evaluator.evaluate("=POISSON.DIST(2, 5, FALSE)")  // P(X=2) ≈0.084
let result2 = evaluator.evaluate("=POISSON.DIST(2, 5, TRUE)")   // P(X≤2) ≈0.125
let result3 = evaluator.evaluate("=POISSON.DIST(10, 10, FALSE)") // ≈0.125
```

**Excel Documentation:** [POISSON.DIST function](https://support.microsoft.com/en-us/office/poisson-dist-function-8fe148ff-39a2-46cb-abf3-7772695d9636)

**Implementation Status:** ✅ Full implementation

**Aliases:** POISSON (legacy)

---

### PROB

Returns the probability that values in a range are between limits.

**Syntax:** `PROB(x_range, prob_range, [lower_limit], [upper_limit])`

**Parameters:**
- `x_range`: Range of numeric values
- `prob_range`: Probabilities associated with x_range
- `lower_limit` *(optional)*: Lower bound (default: minimum of x_range)
- `upper_limit` *(optional)*: Upper bound (default: lower_limit)

**Returns:** Number - Probability sum in range

**Examples:**
```swift
let result1 = evaluator.evaluate("=PROB(A1:A4, B1:B4, 2, 4)")  // P(2 ≤ X ≤ 4)
let result2 = evaluator.evaluate("=PROB({1,2,3}, {0.2,0.5,0.3}, 2)") // P(X=2) = 0.5
```

**Excel Documentation:** [PROB function](https://support.microsoft.com/en-us/office/prob-function-9ac30561-c81c-4259-8253-34f0a238fc49)

**Implementation Status:** ✅ Full implementation

---

### QUARTILE

Returns the quartile of a dataset (inclusive method).

**Syntax:** `QUARTILE(array, quart)`

**Parameters:**
- `array`: Array or range of data
- `quart`: Quartile to return (0=min, 1=Q1, 2=median, 3=Q3, 4=max)

**Returns:** Number - The requested quartile

**Examples:**
```swift
let result1 = evaluator.evaluate("=QUARTILE(A1:A100, 1)")      // Q1 (25th percentile)
let result2 = evaluator.evaluate("=QUARTILE(A1:A100, 2)")      // Median
let result3 = evaluator.evaluate("=QUARTILE(A1:A100, 3)")      // Q3 (75th percentile)
```

**Excel Documentation:** [QUARTILE function](https://support.microsoft.com/en-us/office/quartile-function-93cf8f62-60cd-4fdb-8a92-8451041e1a2a)

**Implementation Status:** ✅ Full implementation

**Aliases:** QUARTILE.INC (current name)

---

### QUARTILE.EXC

Returns the quartile using exclusive method.

**Syntax:** `QUARTILE.EXC(array, quart)`

**Parameters:**
- `array`: Array or range of data
- `quart`: Quartile (1=Q1, 2=median, 3=Q3; 0 and 4 not supported)

**Returns:** Number - The requested quartile

**Examples:**
```swift
let result1 = evaluator.evaluate("=QUARTILE.EXC(A1:A100, 1)")  // Q1 exclusive
let result2 = evaluator.evaluate("=QUARTILE.EXC(A1:A100, 3)")  // Q3 exclusive
```

**Excel Documentation:** [QUARTILE.EXC function](https://support.microsoft.com/en-us/office/quartile-exc-function-5a355b7a-840b-4a01-b0f1-f538c2864cad)

**Implementation Status:** ✅ Full implementation

---

### RANK

Returns the rank of a number in a list.

**Syntax:** `RANK(number, ref, [order])`

**Parameters:**
- `number`: The number to find the rank of
- `ref`: Array or range of numbers
- `order` *(optional)*: 0 or omitted for descending (1=highest), non-zero for ascending

**Returns:** Number - The rank (1 = highest or lowest depending on order)

**Examples:**
```swift
let result1 = evaluator.evaluate("=RANK(7, A1:A10)")          // Rank descending
let result2 = evaluator.evaluate("=RANK(7, A1:A10, 1)")       // Rank ascending
let result3 = evaluator.evaluate("=RANK(5, {1,3,5,7,9})")     // 3 (descending)
```

**Excel Documentation:** [RANK function](https://support.microsoft.com/en-us/office/rank-function-6a2fc49d-1831-4a03-9d8c-c279cf99f723)

**Implementation Status:** ✅ Full implementation

**Aliases:** RANK.EQ (current name)

---

### RANK.AVG

Returns the rank with average handling for ties.

**Syntax:** `RANK.AVG(number, ref, [order])`

**Parameters:**
- `number`: The number to rank
- `ref`: Array or range
- `order` *(optional)*: 0 for descending, non-zero for ascending

**Returns:** Number - Average rank for ties

**Examples:**
```swift
let result = evaluator.evaluate("=RANK.AVG(3, {1,2,3,3,5})")  // 3.5 (avg of ranks 3 and 4)
```

**Excel Documentation:** [RANK.AVG function](https://support.microsoft.com/en-us/office/rank-avg-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a)

**Implementation Status:** ✅ Full implementation

---

### RSQ

Returns the square of the Pearson correlation coefficient (R²).

**Syntax:** `RSQ(known_y's, known_x's)`

**Parameters:**
- `known_y's`: Dependent values
- `known_x's`: Independent values

**Returns:** Number - R-squared (coefficient of determination, 0-1)

**Examples:**
```swift
let result1 = evaluator.evaluate("=RSQ(A1:A10, B1:B10)")      // Goodness of fit
let result2 = evaluator.evaluate("=RSQ({2,4,6}, {1,2,3})")    // 1 (perfect fit)
let result3 = evaluator.evaluate("=RSQ({1,2,3}, {3,2,1})")    // 1 (linear relationship)
```

**Excel Documentation:** [RSQ function](https://support.microsoft.com/en-us/office/rsq-function-d7161715-250d-4a01-b80d-a8364f2be08f)

**Implementation Status:** ✅ Full implementation

---

### SKEW

Returns the skewness of a distribution (sample).

**Syntax:** `SKEW(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers or ranges (requires ≥3 values)

**Returns:** Number - Skewness coefficient

**Examples:**
```swift
let result1 = evaluator.evaluate("=SKEW(3,4,5,2,3,4,5,6,4,7)")  // Distribution asymmetry
let result2 = evaluator.evaluate("=SKEW(A1:A100)")              // Positive = right-skewed
```

**Excel Documentation:** [SKEW function](https://support.microsoft.com/en-us/office/skew-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa)

**Implementation Status:** ✅ Full implementation

---

### SKEW.P

Returns the skewness of a distribution based on a population.

**Syntax:** `SKEW.P(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Population values

**Returns:** Number - Population skewness

**Examples:**
```swift
let result = evaluator.evaluate("=SKEW.P(A1:A1000)")  // Population skewness
```

**Excel Documentation:** [SKEW.P function](https://support.microsoft.com/en-us/office/skew-p-function-76530a5c-99b9-48a1-8392-26632d542fcb)

**Implementation Status:** ✅ Full implementation

---

### SLOPE

Returns the slope of the linear regression line.

**Syntax:** `SLOPE(known_y's, known_x's)`

**Parameters:**
- `known_y's`: Dependent values
- `known_x's`: Independent values

**Returns:** Number - Slope (m in y = mx + b)

**Examples:**
```swift
let result1 = evaluator.evaluate("=SLOPE(A1:A10, B1:B10)")     // Find m
let result2 = evaluator.evaluate("=SLOPE({1,9,5,7}, {0,4,2,3})") // ≈2.077
let result3 = evaluator.evaluate("=SLOPE({2,4,6}, {1,2,3})")   // 2
```

**Excel Documentation:** [SLOPE function](https://support.microsoft.com/en-us/office/slope-function-11fb8f97-3117-4813-98aa-61d7e01276b9)

**Implementation Status:** ✅ Full implementation

---

### SMALL

Returns the k-th smallest value in a dataset.

**Syntax:** `SMALL(array, k)`

**Parameters:**
- `array`: Array or range of values
- `k`: Position from the smallest (1 = smallest, 2 = 2nd smallest, etc.)

**Returns:** Number - The k-th smallest value

**Examples:**
```swift
let result1 = evaluator.evaluate("=SMALL(A1:A10, 1)")          // Minimum
let result2 = evaluator.evaluate("=SMALL(A1:A10, 3)")          // 3rd smallest
let result3 = evaluator.evaluate("=SMALL({3,7,2,9,1}, 2)")     // 2
```

**Excel Documentation:** [SMALL function](https://support.microsoft.com/en-us/office/small-function-17da8222-7c82-42b2-961b-14c45384df07)

**Implementation Status:** ✅ Full implementation

---

### STANDARDIZE

Returns a normalized value (z-score).

**Syntax:** `STANDARDIZE(x, mean, standard_dev)`

**Parameters:**
- `x`: Value to normalize
- `mean`: Arithmetic mean of distribution
- `standard_dev`: Standard deviation

**Returns:** Number - Z-score = (x - mean) / standard_dev

**Examples:**
```swift
let result1 = evaluator.evaluate("=STANDARDIZE(42, 40, 1.5)")  // ≈1.333
let result2 = evaluator.evaluate("=STANDARDIZE(50, 50, 10)")   // 0 (at mean)
let result3 = evaluator.evaluate("=STANDARDIZE(30, 50, 10)")   // -2
```

**Excel Documentation:** [STANDARDIZE function](https://support.microsoft.com/en-us/office/standardize-function-81d66554-2d54-40ec-ba83-6437108ee775)

**Implementation Status:** ✅ Full implementation

---

### STDEV

Returns the sample standard deviation.

**Syntax:** `STDEV(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Sample values (ignores text and logical)

**Returns:** Number - Sample standard deviation

**Examples:**
```swift
let result1 = evaluator.evaluate("=STDEV(1, 2, 3, 4, 5)")      // ≈1.581
let result2 = evaluator.evaluate("=STDEV(A1:A100)")            // Sample variability
```

**Excel Documentation:** [STDEV function](https://support.microsoft.com/en-us/office/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0)

**Implementation Status:** ✅ Full implementation

**Aliases:** STDEV.S (current name)

---

### STDEV.P

Returns the population standard deviation.

**Syntax:** `STDEV.P(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Population values

**Returns:** Number - Population standard deviation

**Examples:**
```swift
let result = evaluator.evaluate("=STDEV.P(A1:A10000)")  // Entire population
```

**Excel Documentation:** [STDEV.P function](https://support.microsoft.com/en-us/office/stdev-p-function-6e917c05-31a0-496f-ade7-4f4e7462f285)

**Implementation Status:** ✅ Full implementation

---

### STDEVA

Returns the sample standard deviation including text and logical values.

**Syntax:** `STDEVA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values (text=0, TRUE=1, FALSE=0)

**Returns:** Number - Sample standard deviation

**Examples:**
```swift
let result = evaluator.evaluate("=STDEVA(10, 20, TRUE, \"text\")")  // Includes all
```

**Excel Documentation:** [STDEVA function](https://support.microsoft.com/en-us/office/stdeva-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d)

**Implementation Status:** ✅ Full implementation

---

### STDEVPA

Returns the population standard deviation including text and logical values.

**Syntax:** `STDEVPA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Population values (text=0, TRUE=1, FALSE=0)

**Returns:** Number - Population standard deviation

**Examples:**
```swift
let result = evaluator.evaluate("=STDEVPA(5, 10, TRUE, 20)")  // Population with text
```

**Excel Documentation:** [STDEVPA function](https://support.microsoft.com/en-us/office/stdevpa-function-5578d4d6-455a-4308-9991-d405afe2c28c)

**Implementation Status:** ✅ Full implementation

---

### STEYX

Returns the standard error of the predicted y-value for each x in regression.

**Syntax:** `STEYX(known_y's, known_x's)`

**Parameters:**
- `known_y's`: Dependent values
- `known_x's`: Independent values

**Returns:** Number - Standard error of regression

**Examples:**
```swift
let result1 = evaluator.evaluate("=STEYX(A1:A10, B1:B10)")  // Prediction accuracy
let result2 = evaluator.evaluate("=STEYX({2,3,9,1,8}, {6,5,11,7,5})") // ≈2.11
```

**Excel Documentation:** [STEYX function](https://support.microsoft.com/en-us/office/steyx-function-6ce74b2c-449d-4a6e-b9ac-f9cef5ba48ab)

**Implementation Status:** ✅ Full implementation

---

### SUBTOTAL

Returns a subtotal with function number (can ignore hidden rows).

**Syntax:** `SUBTOTAL(function_num, ref1, [ref2], ...)`

**Parameters:**
- `function_num`: Aggregation function
  - 1/101: AVERAGE, 2/102: COUNT, 3/103: COUNTA, 4/104: MAX, 5/105: MIN
  - 6/106: PRODUCT, 7/107: STDEV, 8/108: STDEV.P, 9/109: SUM
  - 10/110: VAR.S, 11/111: VAR.P
  - (100+ variants ignore manually hidden rows)
- `ref1, ref2, ...`: Ranges to aggregate

**Returns:** Number - Subtotal result

**Examples:**
```swift
let result1 = evaluator.evaluate("=SUBTOTAL(9, A1:A10)")    // SUM
let result2 = evaluator.evaluate("=SUBTOTAL(1, B:B)")       // AVERAGE
let result3 = evaluator.evaluate("=SUBTOTAL(109, C1:C100)") // SUM ignoring hidden
```

**Excel Documentation:** [SUBTOTAL function](https://support.microsoft.com/en-us/office/subtotal-function-7b027003-f060-4ade-9040-e478765b9939)

**Implementation Status:** ✅ Full implementation

---

### T.DIST

Returns the Student's t-distribution.

**Syntax:** `T.DIST(x, deg_freedom, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate
- `deg_freedom`: Degrees of freedom (≥1)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=T.DIST(1.96, 60, TRUE)")   // CDF
let result2 = evaluator.evaluate("=T.DIST(1, 10, FALSE)")     // PDF
```

**Excel Documentation:** [T.DIST function](https://support.microsoft.com/en-us/office/t-dist-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2)

**Implementation Status:** ✅ Full implementation

---

### T.DIST.RT

Returns the right-tailed Student's t-distribution.

**Syntax:** `T.DIST.RT(x, deg_freedom)`

**Parameters:**
- `x`: Value at which to evaluate
- `deg_freedom`: Degrees of freedom

**Returns:** Number - Right-tail probability

**Examples:**
```swift
let result = evaluator.evaluate("=T.DIST.RT(1.96, 30)")  // P(T > 1.96)
```

**Excel Documentation:** [T.DIST.RT function](https://support.microsoft.com/en-us/office/t-dist-rt-function-20a30020-86f9-4b35-af1f-7ef6ae683eda)

**Implementation Status:** ✅ Full implementation

---

### T.DIST.2T

Returns the two-tailed Student's t-distribution.

**Syntax:** `T.DIST.2T(x, deg_freedom)`

**Parameters:**
- `x`: Value at which to evaluate (x ≥ 0)
- `deg_freedom`: Degrees of freedom

**Returns:** Number - Two-tail probability

**Examples:**
```swift
let result = evaluator.evaluate("=T.DIST.2T(1.96, 30)")  // P(|T| > 1.96)
```

**Excel Documentation:** [T.DIST.2T function](https://support.microsoft.com/en-us/office/t-dist-2t-function-198e9340-e360-4230-bd21-f52f22ff5c28)

**Implementation Status:** ✅ Full implementation

---

### T.INV

Returns the inverse of the left-tailed Student's t-distribution.

**Syntax:** `T.INV(probability, deg_freedom)`

**Parameters:**
- `probability`: Probability associated with t-distribution (0-1)
- `deg_freedom`: Degrees of freedom

**Returns:** Number - Inverse value

**Examples:**
```swift
let result = evaluator.evaluate("=T.INV(0.025, 30)")  // Critical value for two-tailed test
```

**Excel Documentation:** [T.INV function](https://support.microsoft.com/en-us/office/t-inv-function-2908272b-4e61-4942-9df9-a25fec9b0e2e)

**Implementation Status:** ✅ Full implementation

---

### T.INV.2T

Returns the inverse of the two-tailed Student's t-distribution.

**Syntax:** `T.INV.2T(probability, deg_freedom)`

**Parameters:**
- `probability`: Two-tailed probability (0-1)
- `deg_freedom`: Degrees of freedom

**Returns:** Number - Inverse value (positive)

**Examples:**
```swift
let result = evaluator.evaluate("=T.INV.2T(0.05, 30)")  // ≈2.042 (95% CI)
```

**Excel Documentation:** [T.INV.2T function](https://support.microsoft.com/en-us/office/t-inv-2t-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17)

**Implementation Status:** ✅ Full implementation

---

### T.TEST

Returns the probability associated with Student's t-test.

**Syntax:** `T.TEST(array1, array2, tails, type)`

**Parameters:**
- `array1`: First data sample
- `array2`: Second data sample
- `tails`: 1 for one-tailed, 2 for two-tailed
- `type`: 1=paired, 2=equal variance, 3=unequal variance

**Returns:** Number - P-value

**Examples:**
```swift
let result1 = evaluator.evaluate("=T.TEST(A1:A10, B1:B10, 2, 2)")  // Two-tailed, equal var
let result2 = evaluator.evaluate("=T.TEST(Before, After, 1, 1)")   // Paired, one-tailed
```

**Excel Documentation:** [T.TEST function](https://support.microsoft.com/en-us/office/t-test-function-d4e08ec3-c545-485f-962e-276f7cbed055)

**Implementation Status:** ✅ Full implementation

**Aliases:** TTEST (legacy)

---

### TREND

Returns values along a linear trend.

**Syntax:** `TREND(known_y's, [known_x's], [new_x's], [const])`

**Parameters:**
- `known_y's`: Known y-values
- `known_x's` *(optional)*: Known x-values (default: 1, 2, 3, ...)
- `new_x's` *(optional)*: New x-values for prediction (default: known_x's)
- `const` *(optional)*: TRUE to calculate intercept (default: TRUE)

**Returns:** Array - Predicted linear trend values

**Examples:**
```swift
let result1 = evaluator.evaluate("=TREND(A1:A10, B1:B10, B11:B15)")  // Linear forecast
let result2 = evaluator.evaluate("=TREND({1,2,4,8}, {1,2,3,4})")     // Fit to exponential data
```

**Excel Documentation:** [TREND function](https://support.microsoft.com/en-us/office/trend-function-e2f135f0-8827-4096-9873-9a7cf7b51ef1)

**Implementation Status:** ✅ Full implementation

---

### TRIMMEAN

Returns the mean of the interior of a dataset (excludes extremes).

**Syntax:** `TRIMMEAN(array, percent)`

**Parameters:**
- `array`: Array or range of values
- `percent`: Fractional number of data points to exclude (0 to 1)

**Returns:** Number - Trimmed mean

**Examples:**
```swift
let result1 = evaluator.evaluate("=TRIMMEAN(A1:A100, 0.2)")  // Exclude 20% extremes
let result2 = evaluator.evaluate("=TRIMMEAN({1,2,3,4,5,100}, 0.3)") // Robust average
```

**Excel Documentation:** [TRIMMEAN function](https://support.microsoft.com/en-us/office/trimmean-function-d90c9878-a119-4746-88fa-63d988f511d3)

**Implementation Status:** ✅ Full implementation

---

### VAR

Returns the sample variance.

**Syntax:** `VAR(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Sample values (ignores text and logical)

**Returns:** Number - Sample variance

**Examples:**
```swift
let result1 = evaluator.evaluate("=VAR(1, 2, 3, 4, 5)")      // ≈2.5
let result2 = evaluator.evaluate("=VAR(A1:A100)")            // Sample variability
```

**Excel Documentation:** [VAR function](https://support.microsoft.com/en-us/office/var-function-1f2b7ab2-954d-4e17-ba2c-9e58b15a7da2)

**Implementation Status:** ✅ Full implementation

**Aliases:** VAR.S (current name)

---

### VAR.P

Returns the population variance.

**Syntax:** `VAR.P(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Population values

**Returns:** Number - Population variance

**Examples:**
```swift
let result = evaluator.evaluate("=VAR.P(A1:A10000)")  // Entire population
```

**Excel Documentation:** [VAR.P function](https://support.microsoft.com/en-us/office/var-p-function-73d1285c-108c-4843-ba5d-a51f90656f3a)

**Implementation Status:** ✅ Full implementation

---

### VARA

Returns the sample variance including text and logical values.

**Syntax:** `VARA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values (text=0, TRUE=1, FALSE=0)

**Returns:** Number - Sample variance

**Examples:**
```swift
let result = evaluator.evaluate("=VARA(10, 20, TRUE, \"text\")")  // Includes all
```

**Excel Documentation:** [VARA function](https://support.microsoft.com/en-us/office/vara-function-3de77469-fa3a-47b4-85fd-81758a1e1d07)

**Implementation Status:** ✅ Full implementation

---

### VARPA

Returns the population variance including text and logical values.

**Syntax:** `VARPA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Population values (text=0, TRUE=1, FALSE=0)

**Returns:** Number - Population variance

**Examples:**
```swift
let result = evaluator.evaluate("=VARPA(5, 10, TRUE, 20)")  // Population with text
```

**Excel Documentation:** [VARPA function](https://support.microsoft.com/en-us/office/varpa-function-59a62635-4e89-4fad-88ac-ce4dc0513b96)

**Implementation Status:** ✅ Full implementation

---

### WEIBULL.DIST

Returns the Weibull distribution.

**Syntax:** `WEIBULL.DIST(x, alpha, beta, cumulative)`

**Parameters:**
- `x`: Value at which to evaluate (x ≥ 0)
- `alpha`: Shape parameter (α > 0)
- `beta`: Scale parameter (β > 0)
- `cumulative`: TRUE for cumulative, FALSE for probability density

**Returns:** Number - Distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=WEIBULL.DIST(105, 20, 100, TRUE)")   // CDF
let result2 = evaluator.evaluate("=WEIBULL.DIST(2, 1.5, 1, FALSE)")     // PDF
```

**Excel Documentation:** [WEIBULL.DIST function](https://support.microsoft.com/en-us/office/weibull-dist-function-4e783c39-9325-49be-bbc9-a83ef82b45db)

**Implementation Status:** ✅ Full implementation

**Aliases:** WEIBULL (legacy)

---

### Z.TEST

Returns the one-tailed P-value of a z-test.

**Syntax:** `Z.TEST(array, x, [sigma])`

**Parameters:**
- `array`: Array or range of data
- `x`: Value to test
- `sigma` *(optional)*: Population standard deviation (if omitted, uses sample stdev)

**Returns:** Number - One-tailed P-value

**Examples:**
```swift
let result1 = evaluator.evaluate("=Z.TEST(A1:A100, 50)")         // Test mean = 50
let result2 = evaluator.evaluate("=Z.TEST(Data, 100, 15)")       // With known σ
```

**Excel Documentation:** [Z.TEST function](https://support.microsoft.com/en-us/office/z-test-function-d633d5a3-2031-4614-a016-92180ad82bee)

**Implementation Status:** ✅ Full implementation

**Aliases:** ZTEST (legacy)

---

## Legacy Function Aliases

The following legacy function names are supported for backward compatibility:

- **BETADIST** → BETA.DIST
- **BETAINV** → BETA.INV
- **BINOMDIST** → BINOM.DIST
- **CHITEST** → CHISQ.TEST
- **CONFIDENCE** → CONFIDENCE.NORM
- **COVAR** → COVARIANCE.P
- **EXPONDIST** → EXPON.DIST
- **FTEST** → F.TEST
- **GAMMADIST** → GAMMA.DIST
- **GAMMAINV** → GAMMA.INV
- **HYPGEOMDIST** → HYPGEOM.DIST
- **LOGINV** → LOGNORM.INV
- **LOGNORMDIST** → LOGNORM.DIST
- **LOGNORMINV** → LOGNORM.INV
- **MODE** → MODE.SNGL
- **NEGBINOMDIST** → NEGBINOM.DIST
- **NORMDIST** → NORM.DIST
- **NORMINV** → NORM.INV
- **NORMSDIST** → NORM.S.DIST
- **NORMSINV** → NORM.S.INV
- **PERCENTILE** → PERCENTILE.INC
- **PERCENTRANK** → PERCENTRANK.INC
- **POISSON** → POISSON.DIST
- **QUARTILE** → QUARTILE.INC
- **RANK** → RANK.EQ
- **STDEV** → STDEV.S
- **TTEST** → T.TEST
- **WEIBULL** → WEIBULL.DIST
- **ZTEST** → Z.TEST

## See Also

- ``MathematicalFunctions`` - Mathematical and trigonometric functions
- ``FinancialFunctions`` - Financial calculation functions
- ``LogicalFunctions`` - Logical operations and conditional functions
- ``TextFunctions`` - Text manipulation functions
- ``FormulaEvaluator`` - The core formula evaluation engine
