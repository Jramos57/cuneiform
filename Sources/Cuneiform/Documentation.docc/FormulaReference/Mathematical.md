# Mathematical Functions

Comprehensive mathematical and trigonometric functions for spreadsheet calculations in Cuneiform.

## Overview

Cuneiform provides a complete set of mathematical and trigonometric functions compatible with Excel formulas. These functions cover basic arithmetic operations, rounding, trigonometry, logarithms, combinatorics, and advanced mathematical calculations.

All mathematical functions are implemented in the `FormulaEvaluator` and can be used in cell formulas throughout your spreadsheet application.

## Quick Reference

### Basic Math
- ``SUM`` - Sum of numbers
- ``AVERAGE`` - Average of numbers
- ``PRODUCT`` - Product of numbers (not yet implemented)
- ``ABS`` - Absolute value
- ``SIGN`` - Sign of a number

### Rounding Functions
- ``ROUND`` - Round to specified digits
- ``ROUNDUP`` - Round away from zero
- ``ROUNDDOWN`` - Round toward zero
- ``CEILING`` - Round up to multiple
- ``FLOOR`` - Round down to multiple
- ``TRUNC`` - Truncate to integer
- ``INT`` - Round down to integer
- ``MROUND`` - Round to nearest multiple
- ``EVEN`` - Round to nearest even integer
- ``ODD`` - Round to nearest odd integer

### Trigonometric Functions
- ``SIN``, ``COS``, ``TAN`` - Basic trigonometric functions
- ``ASIN``, ``ACOS``, ``ATAN``, ``ATAN2`` - Inverse trigonometric functions
- ``SEC``, ``CSC``, ``COT`` - Secant, cosecant, cotangent
- ``ACOT`` - Inverse cotangent
- ``SINH``, ``COSH``, ``TANH`` - Hyperbolic functions (partial support)
- ``ASINH``, ``ATANH``, ``ACOTH`` - Inverse hyperbolic functions
- ``SECH``, ``CSCH``, ``COTH`` - Hyperbolic reciprocal functions
- ``RADIANS``, ``DEGREES`` - Angle conversions
- ``PI`` - Value of π

### Exponential & Logarithmic
- ``EXP`` - e raised to power
- ``LN`` - Natural logarithm
- ``LOG`` - Logarithm with base
- ``LOG10`` - Base-10 logarithm
- ``POWER`` - Number raised to power
- ``SQRT`` - Square root
- ``SQRTPI`` - Square root of (number × π)

### Statistical Math
- ``SUMPRODUCT`` - Sum of products
- ``SUMSQ`` - Sum of squares
- ``SUMX2MY2`` - Sum of differences of squares
- ``SUMX2PY2`` - Sum of sums of squares
- ``SUMXMY2`` - Sum of squared differences
- ``SERIESSUM`` - Sum of power series

### Combinatorics & Factorials
- ``FACT`` - Factorial
- ``FACTDOUBLE`` - Double factorial
- ``COMBIN`` - Combinations
- ``COMBINA`` - Combinations with repetition (not yet implemented)
- ``PERMUT`` - Permutations
- ``PERMUTATIONA`` - Permutations with repetition (not yet implemented)
- ``MULTINOMIAL`` - Multinomial coefficient

### Number Theory
- ``GCD`` - Greatest common divisor
- ``LCM`` - Least common multiple
- ``MOD`` - Modulo/remainder
- ``QUOTIENT`` - Integer division

### Random Numbers
- ``RAND`` - Random number 0-1
- ``RANDBETWEEN`` - Random integer in range
- ``RANDARRAY`` - Array of random numbers

### Numeric Conversions
- ``ROMAN`` - Convert to Roman numerals
- ``ARABIC`` - Convert from Roman numerals
- ``BASE`` - Convert to base
- ``DECIMAL`` - Convert from base

### Aggregation Variants
- ``MINA`` - Minimum (evaluates text as 0)
- ``MAXA`` - Maximum (evaluates text as 0)

## Function Details

### ABS

Returns the absolute value of a number.

**Syntax:** `ABS(number)`

**Parameters:**
- `number`: The real number for which you want the absolute value

**Returns:** Number - The absolute value

**Examples:**
```swift
let result1 = evaluator.evaluate("=ABS(-5)")    // 5
let result2 = evaluator.evaluate("=ABS(3.14)")  // 3.14
let result3 = evaluator.evaluate("=ABS(0)")     // 0
```

**Excel Documentation:** [ABS function](https://support.microsoft.com/en-us/office/abs-function-3420200f-5628-4e8c-99da-c99d7c87713c)

**Implementation Status:** ✅ Full implementation

---

### ACOS

Returns the arccosine (inverse cosine) of a number in radians.

**Syntax:** `ACOS(number)`

**Parameters:**
- `number`: The cosine of the angle you want, must be between -1 and 1

**Returns:** Number - The arccosine in radians (0 to π)

**Examples:**
```swift
let result1 = evaluator.evaluate("=ACOS(1)")     // 0
let result2 = evaluator.evaluate("=ACOS(0)")     // π/2 (≈1.5708)
let result3 = evaluator.evaluate("=ACOS(-1)")    // π (≈3.1416)
```

**Excel Documentation:** [ACOS function](https://support.microsoft.com/en-us/office/acos-function-cb73173f-d089-4582-afa1-76e5524b5d5b)

**Implementation Status:** ✅ Full implementation

---

### ACOT

Returns the arccotangent (inverse cotangent) of a number in radians.

**Syntax:** `ACOT(number)`

**Parameters:**
- `number`: Any real number

**Returns:** Number - The arccotangent in radians

**Examples:**
```swift
let result1 = evaluator.evaluate("=ACOT(1)")     // π/4 (≈0.7854)
let result2 = evaluator.evaluate("=ACOT(0)")     // π/2 (≈1.5708)
```

**Excel Documentation:** [ACOT function](https://support.microsoft.com/en-us/office/acot-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905)

**Implementation Status:** ✅ Full implementation

---

### ACOTH

Returns the inverse hyperbolic cotangent of a number.

**Syntax:** `ACOTH(number)`

**Parameters:**
- `number`: Any real number with absolute value > 1

**Returns:** Number - The inverse hyperbolic cotangent

**Examples:**
```swift
let result = evaluator.evaluate("=ACOTH(2)")     // ≈0.5493
```

**Excel Documentation:** [ACOTH function](https://support.microsoft.com/en-us/office/acoth-function-cc49480f-f684-4171-9fc5-73e4e852300f)

**Implementation Status:** ✅ Full implementation

---

### ARABIC

Converts a Roman numeral to an Arabic numeral.

**Syntax:** `ARABIC(text)`

**Parameters:**
- `text`: A string containing a Roman numeral

**Returns:** Number - The Arabic numeral

**Examples:**
```swift
let result1 = evaluator.evaluate("=ARABIC(\"XIV\")")    // 14
let result2 = evaluator.evaluate("=ARABIC(\"MCMXC\")")  // 1990
let result3 = evaluator.evaluate("=ARABIC(\"IV\")")     // 4
```

**Excel Documentation:** [ARABIC function](https://support.microsoft.com/en-us/office/arabic-function-9a8da418-c17b-4ef9-a657-9370a30a674f)

**Implementation Status:** ✅ Full implementation

---

### ASIN

Returns the arcsine (inverse sine) of a number in radians.

**Syntax:** `ASIN(number)`

**Parameters:**
- `number`: The sine of the angle you want, must be between -1 and 1

**Returns:** Number - The arcsine in radians (-π/2 to π/2)

**Examples:**
```swift
let result1 = evaluator.evaluate("=ASIN(0.5)")   // π/6 (≈0.5236)
let result2 = evaluator.evaluate("=ASIN(1)")     // π/2 (≈1.5708)
let result3 = evaluator.evaluate("=ASIN(-1)")    // -π/2 (≈-1.5708)
```

**Excel Documentation:** [ASIN function](https://support.microsoft.com/en-us/office/asin-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347)

**Implementation Status:** ✅ Full implementation

---

### ASINH

Returns the inverse hyperbolic sine of a number.

**Syntax:** `ASINH(number)`

**Parameters:**
- `number`: Any real number

**Returns:** Number - The inverse hyperbolic sine

**Examples:**
```swift
let result1 = evaluator.evaluate("=ASINH(1)")    // ≈0.8814
let result2 = evaluator.evaluate("=ASINH(0)")    // 0
```

**Excel Documentation:** [ASINH function](https://support.microsoft.com/en-us/office/asinh-function-4e00475a-067a-43cf-926a-765b0249717c)

**Implementation Status:** ✅ Full implementation

---

### ATAN

Returns the arctangent (inverse tangent) of a number in radians.

**Syntax:** `ATAN(number)`

**Parameters:**
- `number`: The tangent of the angle you want

**Returns:** Number - The arctangent in radians (-π/2 to π/2)

**Examples:**
```swift
let result1 = evaluator.evaluate("=ATAN(1)")     // π/4 (≈0.7854)
let result2 = evaluator.evaluate("=ATAN(0)")     // 0
```

**Excel Documentation:** [ATAN function](https://support.microsoft.com/en-us/office/atan-function-50746fa8-630a-406b-81d0-4a2aed395543)

**Implementation Status:** ✅ Full implementation

---

### ATAN2

Returns the arctangent of x and y coordinates in radians.

**Syntax:** `ATAN2(x_num, y_num)`

**Parameters:**
- `x_num`: The x-coordinate
- `y_num`: The y-coordinate

**Returns:** Number - The arctangent in radians (-π to π)

**Examples:**
```swift
let result1 = evaluator.evaluate("=ATAN2(1, 1)")     // π/4 (≈0.7854)
let result2 = evaluator.evaluate("=ATAN2(1, 0)")     // 0
let result3 = evaluator.evaluate("=ATAN2(-1, 0)")    // π (≈3.1416)
```

**Excel Documentation:** [ATAN2 function](https://support.microsoft.com/en-us/office/atan2-function-c04592ab-b9e3-4908-b428-c96b3a565033)

**Implementation Status:** ✅ Full implementation

---

### ATANH

Returns the inverse hyperbolic tangent of a number.

**Syntax:** `ATANH(number)`

**Parameters:**
- `number`: Any real number with absolute value < 1

**Returns:** Number - The inverse hyperbolic tangent

**Examples:**
```swift
let result = evaluator.evaluate("=ATANH(0.5)")   // ≈0.5493
```

**Excel Documentation:** [ATANH function](https://support.microsoft.com/en-us/office/atanh-function-3cd65768-0de7-4f1d-b312-d01c8c930d90)

**Implementation Status:** ✅ Full implementation

---

### BASE

Converts a number to text representation in a given base.

**Syntax:** `BASE(number, radix, [min_length])`

**Parameters:**
- `number`: The number to convert (must be ≥ 0)
- `radix`: The base to convert to (2-36)
- `min_length` *(optional)*: Minimum length of returned string (pads with zeros)

**Returns:** Text - The number in the specified base

**Examples:**
```swift
let result1 = evaluator.evaluate("=BASE(15, 16)")      // "F"
let result2 = evaluator.evaluate("=BASE(7, 2)")        // "111"
let result3 = evaluator.evaluate("=BASE(7, 2, 8)")     // "00000111"
```

**Excel Documentation:** [BASE function](https://support.microsoft.com/en-us/office/base-function-2ef61411-aee9-4f29-a811-1c42456c6342)

**Implementation Status:** ✅ Full implementation

---

### CEILING

Rounds a number up to the nearest multiple of significance.

**Syntax:** `CEILING(number, [significance])`

**Parameters:**
- `number`: The value to round
- `significance` *(optional)*: The multiple to round to (default: 1)

**Returns:** Number - The rounded value

**Examples:**
```swift
let result1 = evaluator.evaluate("=CEILING(4.3)")      // 5
let result2 = evaluator.evaluate("=CEILING(4.3, 2)")   // 6
let result3 = evaluator.evaluate("=CEILING(-4.3, 2)")  // -4
```

**Excel Documentation:** [CEILING function](https://support.microsoft.com/en-us/office/ceiling-function-0a5cd7c8-0720-4f0a-bd2c-c943e510899f)

**Implementation Status:** ✅ Full implementation

**Aliases:** CEILING.MATH, CEILING.PRECISE, ISO.CEILING

---

### COMBIN

Returns the number of combinations for a given number of items.

**Syntax:** `COMBIN(number, number_chosen)`

**Parameters:**
- `number`: Total number of items
- `number_chosen`: Number of items in each combination

**Returns:** Number - The number of combinations (nCr)

**Examples:**
```swift
let result1 = evaluator.evaluate("=COMBIN(5, 2)")    // 10
let result2 = evaluator.evaluate("=COMBIN(10, 3)")   // 120
let result3 = evaluator.evaluate("=COMBIN(8, 8)")    // 1
```

**Excel Documentation:** [COMBIN function](https://support.microsoft.com/en-us/office/combin-function-12a3f276-0a21-423a-8de6-06990aaf638a)

**Implementation Status:** ✅ Full implementation

---

### COS

Returns the cosine of an angle in radians.

**Syntax:** `COS(number)`

**Parameters:**
- `number`: The angle in radians

**Returns:** Number - The cosine of the angle

**Examples:**
```swift
let result1 = evaluator.evaluate("=COS(0)")      // 1
let result2 = evaluator.evaluate("=COS(PI())")   // -1
let result3 = evaluator.evaluate("=COS(PI()/3)") // 0.5
```

**Excel Documentation:** [COS function](https://support.microsoft.com/en-us/office/cos-function-0fb808a5-95d6-4553-8148-22aebdce5f05)

**Implementation Status:** ✅ Full implementation

---

### COT

Returns the cotangent of an angle in radians.

**Syntax:** `COT(number)`

**Parameters:**
- `number`: The angle in radians

**Returns:** Number - The cotangent of the angle

**Examples:**
```swift
let result = evaluator.evaluate("=COT(PI()/4)")  // 1
```

**Excel Documentation:** [COT function](https://support.microsoft.com/en-us/office/cot-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a)

**Implementation Status:** ✅ Full implementation

---

### COTH

Returns the hyperbolic cotangent of a number.

**Syntax:** `COTH(number)`

**Parameters:**
- `number`: Any real number except 0

**Returns:** Number - The hyperbolic cotangent

**Examples:**
```swift
let result = evaluator.evaluate("=COTH(2)")      // ≈1.0373
```

**Excel Documentation:** [COTH function](https://support.microsoft.com/en-us/office/coth-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5)

**Implementation Status:** ✅ Full implementation

---

### CSC

Returns the cosecant of an angle in radians.

**Syntax:** `CSC(number)`

**Parameters:**
- `number`: The angle in radians

**Returns:** Number - The cosecant of the angle

**Examples:**
```swift
let result = evaluator.evaluate("=CSC(PI()/2)")  // 1
```

**Excel Documentation:** [CSC function](https://support.microsoft.com/en-us/office/csc-function-07379361-219a-4398-8675-07ddc4f135c1)

**Implementation Status:** ✅ Full implementation

---

### CSCH

Returns the hyperbolic cosecant of a number.

**Syntax:** `CSCH(number)`

**Parameters:**
- `number`: Any real number except 0

**Returns:** Number - The hyperbolic cosecant

**Examples:**
```swift
let result = evaluator.evaluate("=CSCH(1)")      // ≈0.8509
```

**Excel Documentation:** [CSCH function](https://support.microsoft.com/en-us/office/csch-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb)

**Implementation Status:** ✅ Full implementation

---

### DECIMAL

Converts a text representation of a number in a given base to decimal.

**Syntax:** `DECIMAL(text, radix)`

**Parameters:**
- `text`: The text representation of the number
- `radix`: The base of the number (2-36)

**Returns:** Number - The decimal value

**Examples:**
```swift
let result1 = evaluator.evaluate("=DECIMAL(\"FF\", 16)")   // 255
let result2 = evaluator.evaluate("=DECIMAL(\"111\", 2)")   // 7
let result3 = evaluator.evaluate("=DECIMAL(\"10\", 8)")    // 8
```

**Excel Documentation:** [DECIMAL function](https://support.microsoft.com/en-us/office/decimal-function-ee554665-6176-46ef-82de-0a283658da2e)

**Implementation Status:** ✅ Full implementation

---

### DEGREES

Converts radians to degrees.

**Syntax:** `DEGREES(angle)`

**Parameters:**
- `angle`: The angle in radians

**Returns:** Number - The angle in degrees

**Examples:**
```swift
let result1 = evaluator.evaluate("=DEGREES(PI())")     // 180
let result2 = evaluator.evaluate("=DEGREES(PI()/2)")   // 90
let result3 = evaluator.evaluate("=DEGREES(0)")        // 0
```

**Excel Documentation:** [DEGREES function](https://support.microsoft.com/en-us/office/degrees-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1)

**Implementation Status:** ✅ Full implementation

---

### EVEN

Rounds a number up to the nearest even integer.

**Syntax:** `EVEN(number)`

**Parameters:**
- `number`: The value to round

**Returns:** Number - The nearest even integer (away from zero)

**Examples:**
```swift
let result1 = evaluator.evaluate("=EVEN(3)")       // 4
let result2 = evaluator.evaluate("=EVEN(2)")       // 2
let result3 = evaluator.evaluate("=EVEN(-3)")      // -4
let result4 = evaluator.evaluate("=EVEN(1.5)")     // 2
```

**Excel Documentation:** [EVEN function](https://support.microsoft.com/en-us/office/even-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9)

**Implementation Status:** ✅ Full implementation

---

### EXP

Returns e raised to the power of a number.

**Syntax:** `EXP(number)`

**Parameters:**
- `number`: The exponent

**Returns:** Number - e^number

**Examples:**
```swift
let result1 = evaluator.evaluate("=EXP(1)")      // ≈2.7183 (e)
let result2 = evaluator.evaluate("=EXP(0)")      // 1
let result3 = evaluator.evaluate("=EXP(2)")      // ≈7.3891
```

**Excel Documentation:** [EXP function](https://support.microsoft.com/en-us/office/exp-function-c578f034-2c45-4c37-bc8c-329660a63abe)

**Implementation Status:** ✅ Full implementation

---

### FACT

Returns the factorial of a number.

**Syntax:** `FACT(number)`

**Parameters:**
- `number`: The non-negative number to calculate the factorial of

**Returns:** Number - number!

**Examples:**
```swift
let result1 = evaluator.evaluate("=FACT(5)")     // 120
let result2 = evaluator.evaluate("=FACT(0)")     // 1
let result3 = evaluator.evaluate("=FACT(10)")    // 3628800
```

**Excel Documentation:** [FACT function](https://support.microsoft.com/en-us/office/fact-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3)

**Implementation Status:** ✅ Full implementation (supports up to 170!)

---

### FACTDOUBLE

Returns the double factorial of a number.

**Syntax:** `FACTDOUBLE(number)`

**Parameters:**
- `number`: The non-negative number

**Returns:** Number - The double factorial (n!! = n × (n-2) × (n-4) × ...)

**Examples:**
```swift
let result1 = evaluator.evaluate("=FACTDOUBLE(6)")     // 48 (6×4×2)
let result2 = evaluator.evaluate("=FACTDOUBLE(7)")     // 105 (7×5×3×1)
let result3 = evaluator.evaluate("=FACTDOUBLE(1)")     // 1
```

**Excel Documentation:** [FACTDOUBLE function](https://support.microsoft.com/en-us/office/factdouble-function-e67697ac-d214-48eb-b7b7-cce2589ecac8)

**Implementation Status:** ✅ Full implementation

---

### FLOOR

Rounds a number down to the nearest multiple of significance.

**Syntax:** `FLOOR(number, [significance])`

**Parameters:**
- `number`: The value to round
- `significance` *(optional)*: The multiple to round to (default: 1)

**Returns:** Number - The rounded value

**Examples:**
```swift
let result1 = evaluator.evaluate("=FLOOR(4.7)")       // 4
let result2 = evaluator.evaluate("=FLOOR(4.3, 2)")    // 4
let result3 = evaluator.evaluate("=FLOOR(-4.3, 2)")   // -6
```

**Excel Documentation:** [FLOOR function](https://support.microsoft.com/en-us/office/floor-function-14bb497c-24f2-4e04-b327-b0b4de5a8886)

**Implementation Status:** ✅ Full implementation

**Aliases:** FLOOR.MATH, FLOOR.PRECISE

---

### GCD

Returns the greatest common divisor of two or more integers.

**Syntax:** `GCD(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: One or more integers

**Returns:** Number - The greatest common divisor

**Examples:**
```swift
let result1 = evaluator.evaluate("=GCD(12, 18)")        // 6
let result2 = evaluator.evaluate("=GCD(24, 36, 48)")    // 12
let result3 = evaluator.evaluate("=GCD(7, 13)")         // 1
```

**Excel Documentation:** [GCD function](https://support.microsoft.com/en-us/office/gcd-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a)

**Implementation Status:** ✅ Full implementation

---

### INT

Rounds a number down to the nearest integer.

**Syntax:** `INT(number)`

**Parameters:**
- `number`: The real number to round down

**Returns:** Number - The largest integer ≤ number

**Examples:**
```swift
let result1 = evaluator.evaluate("=INT(8.9)")      // 8
let result2 = evaluator.evaluate("=INT(-8.9)")     // -9
let result3 = evaluator.evaluate("=INT(5)")        // 5
```

**Excel Documentation:** [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef)

**Implementation Status:** ✅ Full implementation

---

### LCM

Returns the least common multiple of integers.

**Syntax:** `LCM(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: One or more integers

**Returns:** Number - The least common multiple

**Examples:**
```swift
let result1 = evaluator.evaluate("=LCM(4, 6)")         // 12
let result2 = evaluator.evaluate("=LCM(5, 10, 15)")    // 30
let result3 = evaluator.evaluate("=LCM(3, 7)")         // 21
```

**Excel Documentation:** [LCM function](https://support.microsoft.com/en-us/office/lcm-function-7152b67a-8bb5-4075-ae5c-06ede5563c94)

**Implementation Status:** ✅ Full implementation

---

### LN

Returns the natural logarithm (base e) of a number.

**Syntax:** `LN(number)`

**Parameters:**
- `number`: The positive real number

**Returns:** Number - The natural logarithm

**Examples:**
```swift
let result1 = evaluator.evaluate("=LN(EXP(1))")    // 1
let result2 = evaluator.evaluate("=LN(1)")         // 0
let result3 = evaluator.evaluate("=LN(10)")        // ≈2.3026
```

**Excel Documentation:** [LN function](https://support.microsoft.com/en-us/office/ln-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f)

**Implementation Status:** ✅ Full implementation

---

### LOG

Returns the logarithm of a number to a specified base.

**Syntax:** `LOG(number, [base])`

**Parameters:**
- `number`: The positive real number
- `base` *(optional)*: The base for the logarithm (default: 10)

**Returns:** Number - The logarithm

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOG(100)")         // 2 (log₁₀)
let result2 = evaluator.evaluate("=LOG(8, 2)")        // 3 (log₂)
let result3 = evaluator.evaluate("=LOG(16, 4)")       // 2
```

**Excel Documentation:** [LOG function](https://support.microsoft.com/en-us/office/log-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280)

**Implementation Status:** ✅ Full implementation

---

### LOG10

Returns the base-10 logarithm of a number.

**Syntax:** `LOG10(number)`

**Parameters:**
- `number`: The positive real number

**Returns:** Number - The base-10 logarithm

**Examples:**
```swift
let result1 = evaluator.evaluate("=LOG10(100)")    // 2
let result2 = evaluator.evaluate("=LOG10(1000)")   // 3
let result3 = evaluator.evaluate("=LOG10(1)")      // 0
```

**Excel Documentation:** [LOG10 function](https://support.microsoft.com/en-us/office/log10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211)

**Implementation Status:** ✅ Full implementation

---

### MAXA

Returns the largest value in a list, evaluating text and logical values.

**Syntax:** `MAXA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values to compare (text = 0, TRUE = 1, FALSE = 0)

**Returns:** Number - The maximum value

**Examples:**
```swift
let result = evaluator.evaluate("=MAXA(5, TRUE, \"hello\", 3)")  // 5
```

**Excel Documentation:** [MAXA function](https://support.microsoft.com/en-us/office/maxa-function-814bda1e-3840-4bff-9365-2f59ac2ee62d)

**Implementation Status:** ✅ Full implementation

---

### MINA

Returns the smallest value in a list, evaluating text and logical values.

**Syntax:** `MINA(value1, [value2], ...)`

**Parameters:**
- `value1, value2, ...`: Values to compare (text = 0, TRUE = 1, FALSE = 0)

**Returns:** Number - The minimum value

**Examples:**
```swift
let result = evaluator.evaluate("=MINA(5, TRUE, \"hello\", 3)")  // 0
```

**Excel Documentation:** [MINA function](https://support.microsoft.com/en-us/office/mina-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3)

**Implementation Status:** ✅ Full implementation

---

### MOD

Returns the remainder after division.

**Syntax:** `MOD(number, divisor)`

**Parameters:**
- `number`: The dividend
- `divisor`: The divisor

**Returns:** Number - The remainder (same sign as divisor)

**Examples:**
```swift
let result1 = evaluator.evaluate("=MOD(10, 3)")      // 1
let result2 = evaluator.evaluate("=MOD(-10, 3)")     // 2
let result3 = evaluator.evaluate("=MOD(10, -3)")     // -2
```

**Excel Documentation:** [MOD function](https://support.microsoft.com/en-us/office/mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3)

**Implementation Status:** ✅ Full implementation

---

### MROUND

Rounds a number to the nearest multiple.

**Syntax:** `MROUND(number, multiple)`

**Parameters:**
- `number`: The value to round
- `multiple`: The multiple to round to

**Returns:** Number - The rounded value

**Examples:**
```swift
let result1 = evaluator.evaluate("=MROUND(10, 3)")     // 9
let result2 = evaluator.evaluate("=MROUND(11, 3)")     // 12
let result3 = evaluator.evaluate("=MROUND(1.5, 0.1)")  // 1.5
```

**Excel Documentation:** [MROUND function](https://support.microsoft.com/en-us/office/mround-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427)

**Implementation Status:** ✅ Full implementation

---

### MULTINOMIAL

Returns the multinomial coefficient.

**Syntax:** `MULTINOMIAL(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Values for the multinomial

**Returns:** Number - (sum of values)! / (n1! × n2! × ... × nk!)

**Examples:**
```swift
let result1 = evaluator.evaluate("=MULTINOMIAL(2, 3)")       // 10
let result2 = evaluator.evaluate("=MULTINOMIAL(2, 3, 4)")    // 1260
```

**Excel Documentation:** [MULTINOMIAL function](https://support.microsoft.com/en-us/office/multinomial-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6)

**Implementation Status:** ✅ Full implementation

---

### ODD

Rounds a number up to the nearest odd integer.

**Syntax:** `ODD(number)`

**Parameters:**
- `number`: The value to round

**Returns:** Number - The nearest odd integer (away from zero)

**Examples:**
```swift
let result1 = evaluator.evaluate("=ODD(4)")        // 5
let result2 = evaluator.evaluate("=ODD(3)")        // 3
let result3 = evaluator.evaluate("=ODD(-4)")       // -5
let result4 = evaluator.evaluate("=ODD(1.5)")      // 3
```

**Excel Documentation:** [ODD function](https://support.microsoft.com/en-us/office/odd-function-deae64eb-e08a-4c88-8b40-6d0b42575c98)

**Implementation Status:** ✅ Full implementation

---

### PERMUT

Returns the number of permutations for a given number of objects.

**Syntax:** `PERMUT(number, number_chosen)`

**Parameters:**
- `number`: Total number of items
- `number_chosen`: Number of items in each permutation

**Returns:** Number - The number of permutations (nPr)

**Examples:**
```swift
let result1 = evaluator.evaluate("=PERMUT(5, 2)")    // 20
let result2 = evaluator.evaluate("=PERMUT(10, 3)")   // 720
let result3 = evaluator.evaluate("=PERMUT(4, 4)")    // 24
```

**Excel Documentation:** [PERMUT function](https://support.microsoft.com/en-us/office/permut-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3)

**Implementation Status:** ✅ Full implementation

---

### PI

Returns the value of π (pi).

**Syntax:** `PI()`

**Parameters:** None

**Returns:** Number - The value of π (≈3.14159265358979)

**Examples:**
```swift
let result1 = evaluator.evaluate("=PI()")              // ≈3.14159
let result2 = evaluator.evaluate("=2*PI()")            // ≈6.28319
let result3 = evaluator.evaluate("=SIN(PI()/2)")       // 1
```

**Excel Documentation:** [PI function](https://support.microsoft.com/en-us/office/pi-function-264199d0-a3ba-46b8-975a-c4a04608989b)

**Implementation Status:** ✅ Full implementation

---

### POWER

Returns the result of a number raised to a power.

**Syntax:** `POWER(number, power)`

**Parameters:**
- `number`: The base number
- `power`: The exponent

**Returns:** Number - number^power

**Examples:**
```swift
let result1 = evaluator.evaluate("=POWER(2, 3)")     // 8
let result2 = evaluator.evaluate("=POWER(5, 2)")     // 25
let result3 = evaluator.evaluate("=POWER(4, 0.5)")   // 2 (square root)
```

**Excel Documentation:** [POWER function](https://support.microsoft.com/en-us/office/power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a)

**Implementation Status:** ✅ Full implementation

---

### QUOTIENT

Returns the integer portion of a division.

**Syntax:** `QUOTIENT(numerator, denominator)`

**Parameters:**
- `numerator`: The dividend
- `denominator`: The divisor

**Returns:** Number - The integer part of the division

**Examples:**
```swift
let result1 = evaluator.evaluate("=QUOTIENT(10, 3)")    // 3
let result2 = evaluator.evaluate("=QUOTIENT(5, 2)")     // 2
let result3 = evaluator.evaluate("=QUOTIENT(-10, 3)")   // -3
```

**Excel Documentation:** [QUOTIENT function](https://support.microsoft.com/en-us/office/quotient-function-9f7bf099-2a18-4282-8fa4-65290cc99dee)

**Implementation Status:** ✅ Full implementation

---

### RADIANS

Converts degrees to radians.

**Syntax:** `RADIANS(angle)`

**Parameters:**
- `angle`: The angle in degrees

**Returns:** Number - The angle in radians

**Examples:**
```swift
let result1 = evaluator.evaluate("=RADIANS(180)")    // π (≈3.1416)
let result2 = evaluator.evaluate("=RADIANS(90)")     // π/2 (≈1.5708)
let result3 = evaluator.evaluate("=RADIANS(0)")      // 0
```

**Excel Documentation:** [RADIANS function](https://support.microsoft.com/en-us/office/radians-function-ac409508-3d48-45f5-ac02-1497c92de5bf)

**Implementation Status:** ✅ Full implementation

---

### RAND

Returns a random number between 0 and 1.

**Syntax:** `RAND()`

**Parameters:** None

**Returns:** Number - A random value in [0, 1)

**Examples:**
```swift
let result = evaluator.evaluate("=RAND()")           // e.g., 0.753
let scaled = evaluator.evaluate("=RAND() * 100")     // Random 0-100
```

**Excel Documentation:** [RAND function](https://support.microsoft.com/en-us/office/rand-function-4cbfa695-8869-4788-8d90-021ea9f5be73)

**Implementation Status:** ✅ Full implementation

---

### RANDARRAY

Returns an array of random numbers.

**Syntax:** `RANDARRAY([rows], [columns], [min], [max], [whole_number])`

**Parameters:**
- `rows` *(optional)*: Number of rows (default: 1)
- `columns` *(optional)*: Number of columns (default: 1)
- `min` *(optional)*: Minimum value (default: 0)
- `max` *(optional)*: Maximum value (default: 1)
- `whole_number` *(optional)*: TRUE for integers (default: FALSE)

**Returns:** Array - Array of random numbers

**Examples:**
```swift
let result1 = evaluator.evaluate("=RANDARRAY(2, 3)")         // 2×3 array
let result2 = evaluator.evaluate("=RANDARRAY(5, 1, 1, 100, TRUE)")  // 5 integers 1-100
```

**Excel Documentation:** [RANDARRAY function](https://support.microsoft.com/en-us/office/randarray-function-21261e55-3bec-4885-86a6-8b0a47fd4d33)

**Implementation Status:** ✅ Full implementation

---

### RANDBETWEEN

Returns a random integer between two values.

**Syntax:** `RANDBETWEEN(bottom, top)`

**Parameters:**
- `bottom`: The smallest integer to return
- `top`: The largest integer to return

**Returns:** Number - A random integer in [bottom, top]

**Examples:**
```swift
let result1 = evaluator.evaluate("=RANDBETWEEN(1, 10)")      // e.g., 7
let result2 = evaluator.evaluate("=RANDBETWEEN(-5, 5)")      // e.g., -2
```

**Excel Documentation:** [RANDBETWEEN function](https://support.microsoft.com/en-us/office/randbetween-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685)

**Implementation Status:** ✅ Full implementation

---

### ROMAN

Converts an Arabic numeral to Roman numerals.

**Syntax:** `ROMAN(number, [form])`

**Parameters:**
- `number`: The Arabic number to convert (0-3999)
- `form` *(optional)*: The type of Roman numeral (0-4, default: 0 for classic)

**Returns:** Text - The Roman numeral

**Examples:**
```swift
let result1 = evaluator.evaluate("=ROMAN(14)")       // "XIV"
let result2 = evaluator.evaluate("=ROMAN(1990)")     // "MCMXC"
let result3 = evaluator.evaluate("=ROMAN(499)")      // "CDXCIX"
```

**Excel Documentation:** [ROMAN function](https://support.microsoft.com/en-us/office/roman-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5)

**Implementation Status:** ✅ Full implementation

---

### ROUND

Rounds a number to a specified number of digits.

**Syntax:** `ROUND(number, [num_digits])`

**Parameters:**
- `number`: The number to round
- `num_digits` *(optional)*: Number of decimal places (default: 0)

**Returns:** Number - The rounded number

**Examples:**
```swift
let result1 = evaluator.evaluate("=ROUND(2.15, 1)")    // 2.2
let result2 = evaluator.evaluate("=ROUND(2.149, 1)")   // 2.1
let result3 = evaluator.evaluate("=ROUND(21.5, -1)")   // 20
```

**Excel Documentation:** [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c)

**Implementation Status:** ✅ Full implementation

---

### ROUNDDOWN

Rounds a number down, toward zero.

**Syntax:** `ROUNDDOWN(number, num_digits)`

**Parameters:**
- `number`: The number to round down
- `num_digits`: Number of decimal places

**Returns:** Number - The rounded number

**Examples:**
```swift
let result1 = evaluator.evaluate("=ROUNDDOWN(3.7, 0)")     // 3
let result2 = evaluator.evaluate("=ROUNDDOWN(-3.7, 0)")    // -3
let result3 = evaluator.evaluate("=ROUNDDOWN(3.14159, 2)") // 3.14
```

**Excel Documentation:** [ROUNDDOWN function](https://support.microsoft.com/en-us/office/rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53)

**Implementation Status:** ✅ Full implementation

---

### ROUNDUP

Rounds a number up, away from zero.

**Syntax:** `ROUNDUP(number, num_digits)`

**Parameters:**
- `number`: The number to round up
- `num_digits`: Number of decimal places

**Returns:** Number - The rounded number

**Examples:**
```swift
let result1 = evaluator.evaluate("=ROUNDUP(3.2, 0)")     // 4
let result2 = evaluator.evaluate("=ROUNDUP(-3.2, 0)")    // -4
let result3 = evaluator.evaluate("=ROUNDUP(3.14159, 2)") // 3.15
```

**Excel Documentation:** [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7)

**Implementation Status:** ✅ Full implementation

---

### SEC

Returns the secant of an angle.

**Syntax:** `SEC(number)`

**Parameters:**
- `number`: The angle in radians

**Returns:** Number - The secant (1/cos)

**Examples:**
```swift
let result = evaluator.evaluate("=SEC(0)")       // 1
```

**Excel Documentation:** [SEC function](https://support.microsoft.com/en-us/office/sec-function-ff224717-9c87-4170-9b58-d069ced6d5f7)

**Implementation Status:** ✅ Full implementation

---

### SECH

Returns the hyperbolic secant of a number.

**Syntax:** `SECH(number)`

**Parameters:**
- `number`: Any real number

**Returns:** Number - The hyperbolic secant (1/cosh)

**Examples:**
```swift
let result = evaluator.evaluate("=SECH(0)")      // 1
```

**Excel Documentation:** [SECH function](https://support.microsoft.com/en-us/office/sech-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f)

**Implementation Status:** ✅ Full implementation

---

### SERIESSUM

Returns the sum of a power series.

**Syntax:** `SERIESSUM(x, n, m, coefficients)`

**Parameters:**
- `x`: The input value to the power series
- `n`: The initial power to raise x
- `m`: The increment to increase n for each term
- `coefficients`: Array of coefficients

**Returns:** Number - Sum of series Σ(coeffᵢ × x^(n + m×i))

**Examples:**
```swift
let result = evaluator.evaluate("=SERIESSUM(1, 0, 1, {1,2,3})")  // 6
```

**Excel Documentation:** [SERIESSUM function](https://support.microsoft.com/en-us/office/seriessum-function-a3ab25b5-1093-4f5b-b084-96c49087f637)

**Implementation Status:** ✅ Full implementation

---

### SIGN

Returns the sign of a number.

**Syntax:** `SIGN(number)`

**Parameters:**
- `number`: Any real number

**Returns:** Number - 1 if positive, -1 if negative, 0 if zero

**Examples:**
```swift
let result1 = evaluator.evaluate("=SIGN(10)")      // 1
let result2 = evaluator.evaluate("=SIGN(-5)")      // -1
let result3 = evaluator.evaluate("=SIGN(0)")       // 0
```

**Excel Documentation:** [SIGN function](https://support.microsoft.com/en-us/office/sign-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8)

**Implementation Status:** ✅ Full implementation

---

### SIN

Returns the sine of an angle in radians.

**Syntax:** `SIN(number)`

**Parameters:**
- `number`: The angle in radians

**Returns:** Number - The sine of the angle

**Examples:**
```swift
let result1 = evaluator.evaluate("=SIN(0)")        // 0
let result2 = evaluator.evaluate("=SIN(PI()/2)")   // 1
let result3 = evaluator.evaluate("=SIN(PI()/6)")   // 0.5
```

**Excel Documentation:** [SIN function](https://support.microsoft.com/en-us/office/sin-function-cf0e3432-8b9e-483c-bc55-a76651c95602)

**Implementation Status:** ✅ Full implementation

---

### SQRT

Returns the square root of a number.

**Syntax:** `SQRT(number)`

**Parameters:**
- `number`: The number to get the square root of (must be ≥ 0)

**Returns:** Number - The square root

**Examples:**
```swift
let result1 = evaluator.evaluate("=SQRT(16)")      // 4
let result2 = evaluator.evaluate("=SQRT(2)")       // ≈1.4142
let result3 = evaluator.evaluate("=SQRT(0)")       // 0
```

**Excel Documentation:** [SQRT function](https://support.microsoft.com/en-us/office/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf)

**Implementation Status:** ✅ Full implementation

---

### SQRTPI

Returns the square root of (number × π).

**Syntax:** `SQRTPI(number)`

**Parameters:**
- `number`: The number to multiply by π before taking the square root

**Returns:** Number - √(number × π)

**Examples:**
```swift
let result1 = evaluator.evaluate("=SQRTPI(1)")     // ≈1.7725 (√π)
let result2 = evaluator.evaluate("=SQRTPI(2)")     // ≈2.5066
```

**Excel Documentation:** [SQRTPI function](https://support.microsoft.com/en-us/office/sqrtpi-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4)

**Implementation Status:** ✅ Full implementation

---

### SUMPRODUCT

Returns the sum of the products of corresponding array components.

**Syntax:** `SUMPRODUCT(array1, [array2], ...)`

**Parameters:**
- `array1, array2, ...`: Arrays whose components you want to multiply and sum

**Returns:** Number - The sum of products

**Examples:**
```swift
let result1 = evaluator.evaluate("=SUMPRODUCT(A1:A3, B1:B3)")  // Σ(A×B)
let result2 = evaluator.evaluate("=SUMPRODUCT({1,2,3}, {4,5,6})")  // 32
```

**Excel Documentation:** [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e)

**Implementation Status:** ✅ Full implementation

---

### SUMSQ

Returns the sum of the squares of the arguments.

**Syntax:** `SUMSQ(number1, [number2], ...)`

**Parameters:**
- `number1, number2, ...`: Numbers to square and sum

**Returns:** Number - Σ(number²)

**Examples:**
```swift
let result1 = evaluator.evaluate("=SUMSQ(3, 4)")           // 25 (3²+4²)
let result2 = evaluator.evaluate("=SUMSQ(2, 3, 4)")        // 29
```

**Excel Documentation:** [SUMSQ function](https://support.microsoft.com/en-us/office/sumsq-function-e3313c02-51cc-4963-aae6-31442d9ec307)

**Implementation Status:** ✅ Full implementation

---

### SUMX2MY2

Returns the sum of the difference of squares.

**Syntax:** `SUMX2MY2(array_x, array_y)`

**Parameters:**
- `array_x`: First array or range
- `array_y`: Second array or range

**Returns:** Number - Σ(x² - y²)

**Examples:**
```swift
let result = evaluator.evaluate("=SUMX2MY2({2,3,9}, {6,5,11})")  // -55
```

**Excel Documentation:** [SUMX2MY2 function](https://support.microsoft.com/en-us/office/sumx2my2-function-9e599cc5-5399-48e9-a5e0-e37812dfa3e9)

**Implementation Status:** ✅ Full implementation

---

### SUMX2PY2

Returns the sum of the sum of squares.

**Syntax:** `SUMX2PY2(array_x, array_y)`

**Parameters:**
- `array_x`: First array or range
- `array_y`: Second array or range

**Returns:** Number - Σ(x² + y²)

**Examples:**
```swift
let result = evaluator.evaluate("=SUMX2PY2({2,3,9}, {6,5,11})")  // 247
```

**Excel Documentation:** [SUMX2PY2 function](https://support.microsoft.com/en-us/office/sumx2py2-function-826b60b4-0aa2-4e5e-81d2-be704d3d786f)

**Implementation Status:** ✅ Full implementation

---

### SUMXMY2

Returns the sum of squares of differences.

**Syntax:** `SUMXMY2(array_x, array_y)`

**Parameters:**
- `array_x`: First array or range
- `array_y`: Second array or range

**Returns:** Number - Σ((x - y)²)

**Examples:**
```swift
let result = evaluator.evaluate("=SUMXMY2({2,3,9}, {6,5,11})")  // 24
```

**Excel Documentation:** [SUMXMY2 function](https://support.microsoft.com/en-us/office/sumxmy2-function-9d144ac1-4d79-43de-b524-e2ecdf8dc0f9)

**Implementation Status:** ✅ Full implementation

---

### TAN

Returns the tangent of an angle in radians.

**Syntax:** `TAN(number)`

**Parameters:**
- `number`: The angle in radians

**Returns:** Number - The tangent of the angle

**Examples:**
```swift
let result1 = evaluator.evaluate("=TAN(0)")        // 0
let result2 = evaluator.evaluate("=TAN(PI()/4)")   // 1
```

**Excel Documentation:** [TAN function](https://support.microsoft.com/en-us/office/tan-function-08851a40-179f-4052-b789-d7f699447401)

**Implementation Status:** ✅ Full implementation

---

### TRUNC

Truncates a number to a specified number of decimal places.

**Syntax:** `TRUNC(number, [num_digits])`

**Parameters:**
- `number`: The number to truncate
- `num_digits` *(optional)*: Precision (default: 0)

**Returns:** Number - The truncated number

**Examples:**
```swift
let result1 = evaluator.evaluate("=TRUNC(8.9)")       // 8
let result2 = evaluator.evaluate("=TRUNC(-8.9)")      // -8
let result3 = evaluator.evaluate("=TRUNC(3.14159, 2)")  // 3.14
```

**Excel Documentation:** [TRUNC function](https://support.microsoft.com/en-us/office/trunc-function-8b86a64c-3127-43db-ba14-aa5ceb292721)

**Implementation Status:** ✅ Full implementation

---

## Functions Not Yet Implemented

The following mathematical functions are recognized but not yet implemented in Cuneiform:

### ACOSH
Inverse hyperbolic cosine (returns #CALC! error)

### COMBINA
Combinations with repetition (not yet implemented)

### COSH
Hyperbolic cosine (not yet implemented)

### PERMUTATIONA
Permutations with repetition (not yet implemented)

### PRODUCT
Multiplies all numbers (not yet implemented)

### SINH
Hyperbolic sine (not yet implemented)

### TANH
Hyperbolic tangent (not yet implemented)

## See Also

- ``StatisticalFunctions`` - Statistical analysis functions
- ``FinancialFunctions`` - Financial calculation functions
- ``EngineeringFunctions`` - Engineering and conversion functions
- ``LogicalFunctions`` - Logical operations and conditional functions
- ``FormulaEvaluator`` - The core formula evaluation engine
