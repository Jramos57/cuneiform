# Compatibility Functions

Legacy Excel functions maintained for backward compatibility with older spreadsheet versions.

## Overview

Cuneiform provides compatibility functions that were used in older versions of Excel (primarily Excel 2007 and earlier). These functions have been replaced by newer versions with improved accuracy and functionality, but are maintained for backward compatibility with existing spreadsheets.

Microsoft Excel includes these functions to ensure that spreadsheets created in older versions continue to work correctly. When creating new spreadsheets, it's recommended to use the newer function versions instead (e.g., use `BETA.DIST` instead of `BETADIST`).

All compatibility functions are implemented in the `FormulaEvaluator` and behave identically to their Excel counterparts, ensuring full compatibility when reading legacy spreadsheets.

## Quick Reference

### Statistical Distribution Functions (Legacy)
- ``BETADIST`` - Beta distribution (legacy) → Use BETA.DIST
- ``BETAINV`` - Inverse beta distribution (legacy) → Use BETA.INV
- ``GAMMADIST`` - Gamma distribution (legacy) → Use GAMMA.DIST
- ``GAMMAINV`` - Inverse gamma distribution (legacy) → Use GAMMA.INV
- ``LOGNORMDIST`` - Lognormal distribution (legacy) → Use LOGNORM.DIST
- ``LOGNORMINV`` - Inverse lognormal distribution (legacy) → Use LOGNORM.INV
- ``WEIBULL`` - Weibull distribution (legacy) → Use WEIBULL.DIST
- ``HYPGEOMDIST`` - Hypergeometric distribution (legacy) → Use HYPGEOM.DIST
- ``NEGBINOMDIST`` - Negative binomial distribution (legacy) → Use NEGBINOM.DIST
- ``EXPONDIST`` - Exponential distribution (legacy) → Use EXPON.DIST
- ``BINOMDIST`` - Binomial distribution (legacy) → Use BINOM.DIST

### Normal Distribution Functions (Legacy)
- ``NORMDIST`` - Normal distribution (legacy) → Use NORM.DIST
- ``NORMINV`` - Inverse normal distribution (legacy) → Use NORM.INV
- ``NORMSDIST`` - Standard normal distribution (legacy) → Use NORM.S.DIST
- ``NORMSINV`` - Inverse standard normal distribution (legacy) → Use NORM.S.INV

### Dynamic Array & Lambda Functions
- ``LAMBDA`` - Create custom functions (requires dynamic arrays)
- ``LET`` - Assign names to calculation results (requires dynamic arrays)
- ``MAP`` - Map function to array (requires LAMBDA support)
- ``REDUCE`` - Reduce array to single value (requires LAMBDA support)
- ``SCAN`` - Scan array with function (requires LAMBDA support)
- ``BYROW`` - Apply function to each row (requires LAMBDA support)
- ``BYCOL`` - Apply function to each column (requires LAMBDA support)
- ``MAKEARRAY`` - Create array with function (requires LAMBDA support)

## Function Details

### BETADIST

Returns the beta cumulative distribution function (legacy version).

**Syntax:** `BETADIST(x, alpha, beta, [A], [B])`

**Parameters:**
- `x`: The value at which to evaluate the function
- `alpha`: Parameter of the distribution (must be > 0)
- `beta`: Parameter of the distribution (must be > 0)
- `A` *(optional)*: Lower bound of the interval (default: 0)
- `B` *(optional)*: Upper bound of the interval (default: 1)

**Returns:** Number - The cumulative beta distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BETADIST(0.5, 2, 3)")       // ≈0.6875
let result2 = evaluator.evaluate("=BETADIST(0.4, 8, 10, 0, 1)") // ≈0.4059
```

**Excel Documentation:** [BETADIST function](https://support.microsoft.com/en-us/office/betadist-function-49f1b9a9-a5da-470f-8077-5f1730b5fd47)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `BETA.DIST` for new workbooks

**Notes:** This is the legacy version maintained for compatibility. The newer `BETA.DIST` function provides the same functionality with clearer parameter names.

---

### BETAINV

Returns the inverse of the beta cumulative distribution function (legacy version).

**Syntax:** `BETAINV(probability, alpha, beta, [A], [B])`

**Parameters:**
- `probability`: Probability associated with the beta distribution (0 to 1)
- `alpha`: Parameter of the distribution (must be > 0)
- `beta`: Parameter of the distribution (must be > 0)
- `A` *(optional)*: Lower bound of the interval (default: 0)
- `B` *(optional)*: Upper bound of the interval (default: 1)

**Returns:** Number - The value x such that BETADIST(x, alpha, beta, A, B) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=BETAINV(0.5, 2, 3)")        // ≈0.4134
let result2 = evaluator.evaluate("=BETAINV(0.75, 8, 10, 0, 1)") // ≈0.4899
```

**Excel Documentation:** [BETAINV function](https://support.microsoft.com/en-us/office/betainv-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `BETA.INV` for new workbooks

---

### BINOMDIST

Returns the individual term binomial distribution probability (legacy version).

**Syntax:** `BINOMDIST(number_s, trials, probability_s, cumulative)`

**Parameters:**
- `number_s`: Number of successes in trials
- `trials`: Number of independent trials (must be ≥ number_s)
- `probability_s`: Probability of success on each trial (0 to 1)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability mass

**Returns:** Number - The binomial distribution probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=BINOMDIST(6, 10, 0.5, FALSE)")  // ≈0.2051
let result2 = evaluator.evaluate("=BINOMDIST(6, 10, 0.5, TRUE)")   // ≈0.8281
let result3 = evaluator.evaluate("=BINOMDIST(2, 5, 0.3, FALSE)")   // ≈0.3087
```

**Excel Documentation:** [BINOMDIST function](https://support.microsoft.com/en-us/office/binomdist-function-c5ae37b6-f39c-4be2-94c2-509a1480770c)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `BINOM.DIST` for new workbooks

**Notes:** Calculates the probability of getting exactly k successes in n trials (when cumulative=FALSE) or at most k successes (when cumulative=TRUE).

---

### BYCOL

Applies a LAMBDA function to each column in an array (requires dynamic arrays).

**Syntax:** `BYCOL(array, lambda)`

**Parameters:**
- `array`: The array to process
- `lambda`: A LAMBDA function to apply to each column

**Returns:** Array - Results of applying the function to each column

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =BYCOL(A1:C3, LAMBDA(col, SUM(col)))
```

**Excel Documentation:** [BYCOL function](https://support.microsoft.com/en-us/office/bycol-function-58463999-7de5-49ce-8f38-b7f7a2192bfb)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires LAMBDA support and dynamic arrays. Returns `#CALC!` error as these features require dynamic function creation not yet supported in Cuneiform's formula evaluator.

---

### BYROW

Applies a LAMBDA function to each row in an array (requires dynamic arrays).

**Syntax:** `BYROW(array, lambda)`

**Parameters:**
- `array`: The array to process
- `lambda`: A LAMBDA function to apply to each row

**Returns:** Array - Results of applying the function to each row

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =BYROW(A1:C3, LAMBDA(row, AVERAGE(row)))
```

**Excel Documentation:** [BYROW function](https://support.microsoft.com/en-us/office/byrow-function-2e04c677-78c8-4e6b-8c10-a4602f2602bb)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires LAMBDA support and dynamic arrays. Returns `#CALC!` error as these features require dynamic function creation.

---

### EXPONDIST

Returns the exponential distribution (legacy version).

**Syntax:** `EXPONDIST(x, lambda, cumulative)`

**Parameters:**
- `x`: The value at which to evaluate the function (must be ≥ 0)
- `lambda`: The parameter value (must be > 0)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability density

**Returns:** Number - The exponential distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=EXPONDIST(0.5, 2, TRUE)")   // ≈0.6321
let result2 = evaluator.evaluate("=EXPONDIST(0.5, 2, FALSE)")  // ≈0.7358
let result3 = evaluator.evaluate("=EXPONDIST(1, 1, TRUE)")     // ≈0.6321
```

**Excel Documentation:** [EXPONDIST function](https://support.microsoft.com/en-us/office/expondist-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `EXPON.DIST` for new workbooks

**Notes:** The exponential distribution is commonly used to model time between events in a Poisson process.

---

### GAMMADIST

Returns the gamma distribution (legacy version).

**Syntax:** `GAMMADIST(x, alpha, beta, cumulative)`

**Parameters:**
- `x`: The value at which to evaluate the function (must be > 0)
- `alpha`: Shape parameter of the distribution (must be > 0)
- `beta`: Scale parameter of the distribution (must be > 0)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability density

**Returns:** Number - The gamma distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=GAMMADIST(2, 3, 2, TRUE)")   // ≈0.3233
let result2 = evaluator.evaluate("=GAMMADIST(2, 3, 2, FALSE)")  // ≈0.2707
```

**Excel Documentation:** [GAMMADIST function](https://support.microsoft.com/en-us/office/gammadist-function-7327c94d-0f05-4511-83df-1dd7ed23e19e)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `GAMMA.DIST` for new workbooks

**Notes:** The gamma distribution is useful for modeling waiting times and life data analysis.

---

### GAMMAINV

Returns the inverse of the gamma cumulative distribution (legacy version).

**Syntax:** `GAMMAINV(probability, alpha, beta)`

**Parameters:**
- `probability`: Probability associated with the gamma distribution (0 to 1)
- `alpha`: Shape parameter of the distribution (must be > 0)
- `beta`: Scale parameter of the distribution (must be > 0)

**Returns:** Number - The value x such that GAMMADIST(x, alpha, beta, TRUE) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=GAMMAINV(0.5, 3, 2)")   // ≈5.348
let result2 = evaluator.evaluate("=GAMMAINV(0.75, 3, 2)")  // ≈7.655
```

**Excel Documentation:** [GAMMAINV function](https://support.microsoft.com/en-us/office/gammainv-function-06393558-37ab-47d0-aa63-432f99e7916d)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `GAMMA.INV` for new workbooks

---

### HYPGEOMDIST

Returns the hypergeometric distribution (legacy version).

**Syntax:** `HYPGEOMDIST(sample_s, number_sample, population_s, number_pop)`

**Parameters:**
- `sample_s`: Number of successes in the sample
- `number_sample`: Size of the sample
- `population_s`: Number of successes in the population
- `number_pop`: Population size

**Returns:** Number - The hypergeometric distribution probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=HYPGEOMDIST(1, 4, 8, 20)")  // ≈0.3633
let result2 = evaluator.evaluate("=HYPGEOMDIST(2, 5, 10, 20)") // ≈0.3483
```

**Excel Documentation:** [HYPGEOMDIST function](https://support.microsoft.com/en-us/office/hypgeomdist-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22d40)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `HYPGEOM.DIST` for new workbooks

**Aliases:** Also available as `HYPGEOM.DIST`

**Notes:** Models drawing without replacement from a finite population containing successes and failures.

---

### LAMBDA

Creates a custom function that can be called by name (requires dynamic arrays).

**Syntax:** `LAMBDA([parameter1, parameter2, ...], calculation)`

**Parameters:**
- `parameter1, parameter2, ...`: Names for values to pass to the function
- `calculation`: Formula expression using the parameters

**Returns:** Function - A custom function

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =LAMBDA(x, y, x + y)(1, 2)  // Returns 3
```

**Excel Documentation:** [LAMBDA function](https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires dynamic function creation not yet supported in Cuneiform. Returns `#CALC!` error. This is an advanced Excel 365 feature.

---

### LET

Assigns names to calculation results for improved formula readability (requires dynamic arrays).

**Syntax:** `LET(name1, value1, [name2, value2, ...], calculation)`

**Parameters:**
- `name1, value1, ...`: Name-value pairs to define
- `calculation`: Formula that uses the defined names

**Returns:** Any - Result of the calculation

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =LET(x, A1*2, y, x+10, x+y)
```

**Excel Documentation:** [LET function](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires variable binding not yet supported. Returns `#CALC!` error. This Excel 365 feature improves formula performance and readability.

---

### LOGNORMDIST

Returns the lognormal distribution (legacy version).

**Syntax:** `LOGNORMDIST(x, mean, standard_dev)`

**Parameters:**
- `x`: The value at which to evaluate the function (must be > 0)
- `mean`: Mean of ln(x)
- `standard_dev`: Standard deviation of ln(x) (must be > 0)

**Returns:** Number - The cumulative lognormal distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOGNORMDIST(4, 3.5, 1.2)")  // ≈0.0390
let result2 = evaluator.evaluate("=LOGNORMDIST(10, 2, 0.5)")   // ≈0.9981
```

**Excel Documentation:** [LOGNORMDIST function](https://support.microsoft.com/en-us/office/lognormdist-function-eb60d00b-48a9-4217-be2b-6074aee6b070)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `LOGNORM.DIST` for new workbooks

**Notes:** Models data where the logarithm of the variable is normally distributed, common in financial and biological applications.

---

### LOGNORMINV

Returns the inverse of the lognormal cumulative distribution (legacy version).

**Syntax:** `LOGNORMINV(probability, mean, standard_dev)`

**Parameters:**
- `probability`: Probability associated with the lognormal distribution (0 to 1)
- `mean`: Mean of ln(x)
- `standard_dev`: Standard deviation of ln(x) (must be > 0)

**Returns:** Number - The value x such that LOGNORMDIST(x, mean, standard_dev) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOGNORMINV(0.5, 3.5, 1.2)")  // ≈33.12
let result2 = evaluator.evaluate("=LOGNORMINV(0.9, 2, 0.5)")    // ≈11.02
```

**Excel Documentation:** [LOGNORMINV function](https://support.microsoft.com/en-us/office/lognorminv-function-0b1fc8d7-0ee8-44c6-b232-74cab59e6fa1)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `LOGNORM.INV` for new workbooks

---

### MAKEARRAY

Creates a calculated array of specified dimensions (requires LAMBDA support).

**Syntax:** `MAKEARRAY(rows, columns, lambda)`

**Parameters:**
- `rows`: Number of rows in the array
- `columns`: Number of columns in the array
- `lambda`: LAMBDA function to calculate each element

**Returns:** Array - Calculated array

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =MAKEARRAY(3, 3, LAMBDA(r, c, r*c))
```

**Excel Documentation:** [MAKEARRAY function](https://support.microsoft.com/en-us/office/makearray-function-b80da5ad-b338-4149-a523-5b221da09097)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires LAMBDA support. Returns `#CALC!` error. This is an Excel 365 dynamic array function.

---

### MAP

Maps each value in array(s) by applying a LAMBDA function (requires LAMBDA support).

**Syntax:** `MAP(array1, [array2, ...], lambda)`

**Parameters:**
- `array1, array2, ...`: Arrays to map
- `lambda`: LAMBDA function to apply to each set of values

**Returns:** Array - Mapped results

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =MAP(A1:A10, LAMBDA(x, x*2))
```

**Excel Documentation:** [MAP function](https://support.microsoft.com/en-us/office/map-function-48006093-f97c-47c1-bfcc-749263bb1f01)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires LAMBDA support. Returns `#CALC!` error.

---

### NEGBINOMDIST

Returns the negative binomial distribution (legacy version).

**Syntax:** `NEGBINOMDIST(number_f, number_s, probability_s)`

**Parameters:**
- `number_f`: Number of failures
- `number_s`: Threshold number of successes
- `probability_s`: Probability of success

**Returns:** Number - The negative binomial distribution probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=NEGBINOMDIST(10, 5, 0.25)")  // ≈0.0550
let result2 = evaluator.evaluate("=NEGBINOMDIST(5, 3, 0.5)")    // ≈0.0547
```

**Excel Documentation:** [NEGBINOMDIST function](https://support.microsoft.com/en-us/office/negbinomdist-function-c8239f89-c2d0-45bd-b6af-172e570f8599)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `NEGBINOM.DIST` for new workbooks

**Aliases:** Also available as `NEGBINOM.DIST`

**Notes:** Models the number of failures before a specified number of successes occurs.

---

### NORMDIST

Returns the normal distribution (legacy version).

**Syntax:** `NORMDIST(x, mean, standard_dev, cumulative)`

**Parameters:**
- `x`: The value for which you want the distribution
- `mean`: Arithmetic mean of the distribution
- `standard_dev`: Standard deviation of the distribution (must be > 0)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability density

**Returns:** Number - The normal distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORMDIST(42, 40, 1.5, TRUE)")   // ≈0.9088
let result2 = evaluator.evaluate("=NORMDIST(42, 40, 1.5, FALSE)")  // ≈0.1080
```

**Excel Documentation:** [NORMDIST function](https://support.microsoft.com/en-us/office/normdist-function-126db625-c53e-4591-9a22-c9ff422d6d58)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `NORM.DIST` for new workbooks

**Notes:** One of the most commonly used statistical distributions. The bell curve.

---

### NORMINV

Returns the inverse of the normal cumulative distribution (legacy version).

**Syntax:** `NORMINV(probability, mean, standard_dev)`

**Parameters:**
- `probability`: Probability corresponding to the normal distribution (0 to 1)
- `mean`: Arithmetic mean of the distribution
- `standard_dev`: Standard deviation of the distribution (must be > 0)

**Returns:** Number - The value x such that NORMDIST(x, mean, standard_dev, TRUE) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORMINV(0.9088, 40, 1.5)")  // ≈42
let result2 = evaluator.evaluate("=NORMINV(0.5, 100, 10)")     // 100
```

**Excel Documentation:** [NORMINV function](https://support.microsoft.com/en-us/office/norminv-function-87981ab8-2de0-4cb0-b1aa-e21d4cb879b8)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `NORM.INV` for new workbooks

---

### NORMSDIST

Returns the standard normal cumulative distribution (legacy version).

**Syntax:** `NORMSDIST(z)`

**Parameters:**
- `z`: The value for which you want the distribution

**Returns:** Number - The standard normal cumulative distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORMSDIST(0)")      // 0.5
let result2 = evaluator.evaluate("=NORMSDIST(1.96)")   // ≈0.975
let result3 = evaluator.evaluate("=NORMSDIST(-1.96)")  // ≈0.025
```

**Excel Documentation:** [NORMSDIST function](https://support.microsoft.com/en-us/office/normsdist-function-463369ea-0345-445d-802a-4ff0d6ce7cac)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `NORM.S.DIST` for new workbooks

**Notes:** Equivalent to `NORMDIST(z, 0, 1, TRUE)` - the standard normal distribution with mean 0 and standard deviation 1.

---

### NORMSINV

Returns the inverse of the standard normal cumulative distribution (legacy version).

**Syntax:** `NORMSINV(probability)`

**Parameters:**
- `probability`: Probability corresponding to the normal distribution (0 to 1)

**Returns:** Number - The value z such that NORMSDIST(z) = probability

**Examples:**
```swift
let result1 = evaluator.evaluate("=NORMSINV(0.5)")     // 0
let result2 = evaluator.evaluate("=NORMSINV(0.975)")   // ≈1.96
let result3 = evaluator.evaluate("=NORMSINV(0.025)")   // ≈-1.96
```

**Excel Documentation:** [NORMSINV function](https://support.microsoft.com/en-us/office/normsinv-function-8d1bce66-8e4d-4f3b-967c-30eed61f019d)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `NORM.S.INV` for new workbooks

**Notes:** Commonly used to find z-scores and critical values for confidence intervals.

---

### REDUCE

Reduces an array to an accumulated value by applying a LAMBDA function (requires LAMBDA support).

**Syntax:** `REDUCE([initial_value], array, lambda)`

**Parameters:**
- `initial_value` *(optional)*: Starting value for the accumulator
- `array`: Array to reduce
- `lambda`: LAMBDA function to apply (takes accumulator and value)

**Returns:** Any - The accumulated value

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =REDUCE(0, A1:A10, LAMBDA(acc, val, acc+val))
```

**Excel Documentation:** [REDUCE function](https://support.microsoft.com/en-us/office/reduce-function-42e39910-b345-45f3-84b8-0642b568b7cb)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires LAMBDA support. Returns `#CALC!` error. Similar to functional programming's fold/reduce operations.

---

### SCAN

Scans an array by applying a LAMBDA function and returns each intermediate value (requires LAMBDA support).

**Syntax:** `SCAN([initial_value], array, lambda)`

**Parameters:**
- `initial_value` *(optional)*: Starting value for the accumulator
- `array`: Array to scan
- `lambda`: LAMBDA function to apply (takes accumulator and value)

**Returns:** Array - All intermediate accumulated values

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =SCAN(0, A1:A10, LAMBDA(acc, val, acc+val))  // Running sum
```

**Excel Documentation:** [SCAN function](https://support.microsoft.com/en-us/office/scan-function-d58dfd11-9969-4439-b2dc-e7062724de29)

**Implementation Status:** ❌ Not implemented (returns #CALC! error)

**Notes:** Requires LAMBDA support. Returns `#CALC!` error. Like REDUCE but returns all intermediate results.

---

### WEIBULL

Returns the Weibull distribution (legacy version).

**Syntax:** `WEIBULL(x, alpha, beta, cumulative)`

**Parameters:**
- `x`: The value at which to evaluate the function (must be ≥ 0)
- `alpha`: Shape parameter of the distribution (must be > 0)
- `beta`: Scale parameter of the distribution (must be > 0)
- `cumulative`: TRUE for cumulative distribution, FALSE for probability density

**Returns:** Number - The Weibull distribution value

**Examples:**
```swift
let result1 = evaluator.evaluate("=WEIBULL(2, 1.5, 1, TRUE)")   // ≈0.9502
let result2 = evaluator.evaluate("=WEIBULL(2, 1.5, 1, FALSE)")  // ≈0.2492
```

**Excel Documentation:** [WEIBULL function](https://support.microsoft.com/en-us/office/weibull-function-b83dc2c6-260b-4754-bef2-633196f6fdcc)

**Implementation Status:** ✅ Full implementation

**Replacement:** Use `WEIBULL.DIST` for new workbooks

**Notes:** Commonly used in reliability analysis and failure time modeling.

---

## Migration Guide

When updating spreadsheets from older Excel versions, consider replacing compatibility functions with their modern equivalents:

| Legacy Function | Modern Replacement | Benefits |
|----------------|-------------------|----------|
| `BETADIST` | `BETA.DIST` | Clearer syntax, consistent naming |
| `BETAINV` | `BETA.INV` | Better parameter validation |
| `BINOMDIST` | `BINOM.DIST` | Improved accuracy |
| `EXPONDIST` | `EXPON.DIST` | Consistent with other .DIST functions |
| `GAMMADIST` | `GAMMA.DIST` | More accurate calculations |
| `GAMMAINV` | `GAMMA.INV` | Better numerical stability |
| `HYPGEOMDIST` | `HYPGEOM.DIST` | Additional cumulative parameter |
| `LOGNORMDIST` | `LOGNORM.DIST` | PDF support added |
| `LOGNORMINV` | `LOGNORM.INV` | Improved precision |
| `NEGBINOMDIST` | `NEGBINOM.DIST` | Cumulative distribution support |
| `NORMDIST` | `NORM.DIST` | Consistent naming |
| `NORMINV` | `NORM.INV` | Improved algorithm |
| `NORMSDIST` | `NORM.S.DIST` | Better performance |
| `NORMSINV` | `NORM.S.INV` | More accurate for extreme values |
| `WEIBULL` | `WEIBULL.DIST` | Standardized parameters |

**Note:** Cuneiform supports both the legacy and modern versions, so no immediate migration is required for compatibility.

## Dynamic Array Functions

The LAMBDA and related functions (LET, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY) are part of Excel 365's dynamic array features. These functions:

- Enable functional programming patterns in spreadsheets
- Require runtime creation of custom functions
- Are not yet supported in Cuneiform's formula evaluator
- Return `#CALC!` error when encountered

These advanced features may be considered for future implementation as Cuneiform's formula engine evolves.

## See Also

- ``StatisticalFunctions`` - Modern statistical functions
- ``MathematicalFunctions`` - Core mathematical operations
- ``FormulaEvaluator`` - The formula evaluation engine
- <doc:WebService> - Web and external data functions
