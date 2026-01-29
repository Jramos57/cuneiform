# Engineering Functions

Comprehensive engineering functions for unit conversion, number base conversion, bitwise operations, Bessel functions, error functions, and complex number calculations in Cuneiform.

## Overview

Cuneiform provides a complete set of engineering functions compatible with Excel formulas. These functions support unit conversions, number base conversions (binary, octal, decimal, hexadecimal), bitwise operations, Bessel functions, error functions, and comprehensive complex number arithmetic.

All engineering functions are implemented in the `FormulaEvaluator` and can be used in cell formulas throughout your spreadsheet application.

## Quick Reference

### Unit Conversion
- ``CONVERT`` - Convert between measurement units

### Comparison Functions
- ``DELTA`` - Test if two values are equal
- ``GESTEP`` - Test if number is greater than or equal to threshold

### Number Base Conversion
- ``DEC2BIN`` - Convert decimal to binary
- ``DEC2OCT`` - Convert decimal to octal
- ``DEC2HEX`` - Convert decimal to hexadecimal
- ``BIN2DEC`` - Convert binary to decimal
- ``BIN2OCT`` - Convert binary to octal
- ``BIN2HEX`` - Convert binary to hexadecimal
- ``OCT2DEC`` - Convert octal to decimal
- ``OCT2BIN`` - Convert octal to binary
- ``OCT2HEX`` - Convert octal to hexadecimal
- ``HEX2DEC`` - Convert hexadecimal to decimal
- ``HEX2BIN`` - Convert hexadecimal to binary
- ``HEX2OCT`` - Convert hexadecimal to octal

### Bitwise Operations
- ``BITAND`` - Bitwise AND
- ``BITOR`` - Bitwise OR
- ``BITXOR`` - Bitwise XOR
- ``BITLSHIFT`` - Bitwise left shift
- ``BITRSHIFT`` - Bitwise right shift

### Bessel Functions
- ``BESSELI`` - Modified Bessel function In(x)
- ``BESSELJ`` - Bessel function Jn(x)
- ``BESSELK`` - Modified Bessel function Kn(x)
- ``BESSELY`` - Bessel function Yn(x)

### Error Functions
- ``ERF`` - Error function
- ``ERFC`` - Complementary error function

### Complex Number Operations
- ``COMPLEX`` - Create complex number from real and imaginary parts
- ``IMREAL`` - Extract real part
- ``IMAGINARY`` - Extract imaginary part
- ``IMABS`` - Absolute value (modulus)
- ``IMARGUMENT`` - Argument (angle)
- ``IMCONJUGATE`` - Complex conjugate

### Complex Arithmetic
- ``IMADD`` - Addition
- ``IMSUM`` - Sum of multiple complex numbers
- ``IMSUB`` - Subtraction
- ``IMPRODUCT`` - Product of multiple complex numbers
- ``IMDIV`` - Division

### Complex Mathematical Functions
- ``IMSQRT`` - Square root
- ``IMPOWER`` - Raise to power
- ``IMEXP`` - Exponential
- ``IMLN`` - Natural logarithm
- ``IMLOG10`` - Base-10 logarithm
- ``IMLOG2`` - Base-2 logarithm

### Complex Trigonometric Functions
- ``IMSIN`` - Sine
- ``IMCOS`` - Cosine
- ``IMTAN`` - Tangent
- ``IMSEC`` - Secant
- ``IMCSC`` - Cosecant

### Complex Hyperbolic Functions
- ``IMSINH`` - Hyperbolic sine
- ``IMCOSH`` - Hyperbolic cosine
- ``IMTANH`` - Hyperbolic tangent
- ``IMSECH`` - Hyperbolic secant
- ``IMCSCH`` - Hyperbolic cosecant
- ``IMASINH`` - Inverse hyperbolic sine
- ``IMACOSH`` - Inverse hyperbolic cosine
- ``IMATANH`` - Inverse hyperbolic tangent

## Function Details

### BIN2DEC

Converts a binary number to decimal.

**Syntax:** `BIN2DEC(number)`

**Parameters:**
- `number`: The binary number (string) you want to convert (max 10 characters)

**Returns:** Number - The decimal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=BIN2DEC(\"1010\")")      // 10
let result2 = evaluator.evaluate("=BIN2DEC(\"11111111\")")  // 255
let result3 = evaluator.evaluate("=BIN2DEC(\"1\")")         // 1
```

**Excel Documentation:** [BIN2DEC function](https://support.microsoft.com/en-us/office/bin2dec-function-63905b57-b3a0-453d-99f4-647bb519cd6c)

**Implementation Status:** ✅ Full implementation

---

### BIN2HEX

Converts a binary number to hexadecimal.

**Syntax:** `BIN2HEX(number, [places])`

**Parameters:**
- `number`: The binary number (string) you want to convert (max 10 characters)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The hexadecimal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=BIN2HEX(\"1010\")")       // "A"
let result2 = evaluator.evaluate("=BIN2HEX(\"1010\", 4)")    // "000A"
let result3 = evaluator.evaluate("=BIN2HEX(\"11111111\")")   // "FF"
```

**Excel Documentation:** [BIN2HEX function](https://support.microsoft.com/en-us/office/bin2hex-function-0375e507-f5e5-4077-9af8-28d84f9f41cc)

**Implementation Status:** ✅ Full implementation

---

### BIN2OCT

Converts a binary number to octal.

**Syntax:** `BIN2OCT(number, [places])`

**Parameters:**
- `number`: The binary number (string) you want to convert (max 10 characters)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The octal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=BIN2OCT(\"1010\")")      // "12"
let result2 = evaluator.evaluate("=BIN2OCT(\"1010\", 4)")   // "0012"
let result3 = evaluator.evaluate("=BIN2OCT(\"11111111\")")  // "377"
```

**Excel Documentation:** [BIN2OCT function](https://support.microsoft.com/en-us/office/bin2oct-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd)

**Implementation Status:** ✅ Full implementation

---

### BITAND

Returns a bitwise AND of two numbers.

**Syntax:** `BITAND(number1, number2)`

**Parameters:**
- `number1`: Must be in decimal form and greater than or equal to 0
- `number2`: Must be in decimal form and greater than or equal to 0

**Returns:** Number - The bitwise AND result

**Examples:**
```swift
let result1 = evaluator.evaluate("=BITAND(5, 3)")    // 1 (0101 AND 0011 = 0001)
let result2 = evaluator.evaluate("=BITAND(12, 10)")  // 8 (1100 AND 1010 = 1000)
let result3 = evaluator.evaluate("=BITAND(15, 7)")   // 7 (1111 AND 0111 = 0111)
```

**Excel Documentation:** [BITAND function](https://support.microsoft.com/en-us/office/bitand-function-8a2be3d7-91c3-4b48-9517-64548008563a)

**Implementation Status:** ✅ Full implementation

---

### BITLSHIFT

Returns a number shifted left by the specified number of bits.

**Syntax:** `BITLSHIFT(number, shift_amount)`

**Parameters:**
- `number`: The number to be shifted (must be an integer >= 0)
- `shift_amount`: The number of bits to shift left (must be an integer)

**Returns:** Number - The shifted value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BITLSHIFT(5, 2)")   // 20 (0101 << 2 = 10100)
let result2 = evaluator.evaluate("=BITLSHIFT(3, 4)")   // 48 (0011 << 4 = 110000)
let result3 = evaluator.evaluate("=BITLSHIFT(1, 8)")   // 256
```

**Excel Documentation:** [BITLSHIFT function](https://support.microsoft.com/en-us/office/bitlshift-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c)

**Implementation Status:** ✅ Full implementation

---

### BITOR

Returns a bitwise OR of two numbers.

**Syntax:** `BITOR(number1, number2)`

**Parameters:**
- `number1`: Must be in decimal form and greater than or equal to 0
- `number2`: Must be in decimal form and greater than or equal to 0

**Returns:** Number - The bitwise OR result

**Examples:**
```swift
let result1 = evaluator.evaluate("=BITOR(5, 3)")     // 7 (0101 OR 0011 = 0111)
let result2 = evaluator.evaluate("=BITOR(12, 10)")   // 14 (1100 OR 1010 = 1110)
let result3 = evaluator.evaluate("=BITOR(8, 4)")     // 12 (1000 OR 0100 = 1100)
```

**Excel Documentation:** [BITOR function](https://support.microsoft.com/en-us/office/bitor-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2)

**Implementation Status:** ✅ Full implementation

---

### BITRSHIFT

Returns a number shifted right by the specified number of bits.

**Syntax:** `BITRSHIFT(number, shift_amount)`

**Parameters:**
- `number`: The number to be shifted (must be an integer >= 0)
- `shift_amount`: The number of bits to shift right (must be an integer)

**Returns:** Number - The shifted value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BITRSHIFT(20, 2)")   // 5 (10100 >> 2 = 0101)
let result2 = evaluator.evaluate("=BITRSHIFT(48, 4)")   // 3 (110000 >> 4 = 0011)
let result3 = evaluator.evaluate("=BITRSHIFT(256, 8)")  // 1
```

**Excel Documentation:** [BITRSHIFT function](https://support.microsoft.com/en-us/office/bitrshift-function-274d6996-f42c-4743-abdb-4ff95351222c)

**Implementation Status:** ✅ Full implementation

---

### BITXOR

Returns a bitwise XOR of two numbers.

**Syntax:** `BITXOR(number1, number2)`

**Parameters:**
- `number1`: Must be in decimal form and greater than or equal to 0
- `number2`: Must be in decimal form and greater than or equal to 0

**Returns:** Number - The bitwise XOR result

**Examples:**
```swift
let result1 = evaluator.evaluate("=BITXOR(5, 3)")    // 6 (0101 XOR 0011 = 0110)
let result2 = evaluator.evaluate("=BITXOR(12, 10)")  // 6 (1100 XOR 1010 = 0110)
let result3 = evaluator.evaluate("=BITXOR(15, 7)")   // 8 (1111 XOR 0111 = 1000)
```

**Excel Documentation:** [BITXOR function](https://support.microsoft.com/en-us/office/bitxor-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4)

**Implementation Status:** ✅ Full implementation

---

### BESSELI

Returns the modified Bessel function In(x).

**Syntax:** `BESSELI(x, n)`

**Parameters:**
- `x`: The value at which to evaluate the function
- `n`: The order of the Bessel function (must be >= 0)

**Returns:** Number - The modified Bessel function value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BESSELI(1.5, 1)")   // ≈0.9817
let result2 = evaluator.evaluate("=BESSELI(2, 2)")     // ≈0.6889
```

**Excel Documentation:** [BESSELI function](https://support.microsoft.com/en-us/office/besseli-function-8d33855c-9a8d-444b-98e0-852267b1c0df)

**Implementation Status:** ✅ Full implementation

---

### BESSELJ

Returns the Bessel function Jn(x).

**Syntax:** `BESSELJ(x, n)`

**Parameters:**
- `x`: The value at which to evaluate the function
- `n`: The order of the Bessel function (must be >= 0)

**Returns:** Number - The Bessel function value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BESSELJ(1.5, 1)")   // ≈0.5579
let result2 = evaluator.evaluate("=BESSELJ(2, 2)")     // ≈0.3528
```

**Excel Documentation:** [BESSELJ function](https://support.microsoft.com/en-us/office/besselj-function-839cb181-48de-408b-9d80-bd02982d94f7)

**Implementation Status:** ✅ Full implementation

---

### BESSELK

Returns the modified Bessel function Kn(x).

**Syntax:** `BESSELK(x, n)`

**Parameters:**
- `x`: The value at which to evaluate the function (must be > 0)
- `n`: The order of the Bessel function (must be >= 0)

**Returns:** Number - The modified Bessel function value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BESSELK(1.5, 1)")   // ≈0.2773
let result2 = evaluator.evaluate("=BESSELK(2, 2)")     // ≈0.2537
```

**Excel Documentation:** [BESSELK function](https://support.microsoft.com/en-us/office/besselk-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70)

**Implementation Status:** ✅ Full implementation

---

### BESSELY

Returns the Bessel function Yn(x).

**Syntax:** `BESSELY(x, n)`

**Parameters:**
- `x`: The value at which to evaluate the function (must be > 0)
- `n`: The order of the Bessel function (must be >= 0)

**Returns:** Number - The Bessel function value

**Examples:**
```swift
let result1 = evaluator.evaluate("=BESSELY(1.5, 1)")   // ≈-0.4123
let result2 = evaluator.evaluate("=BESSELY(2, 2)")     // ≈-0.6174
```

**Excel Documentation:** [BESSELY function](https://support.microsoft.com/en-us/office/bessely-function-f3a356b3-da89-42c3-8974-2da54d6353a2)

**Implementation Status:** ✅ Full implementation

---

### COMPLEX

Converts real and imaginary coefficients into a complex number.

**Syntax:** `COMPLEX(real_num, i_num, [suffix])`

**Parameters:**
- `real_num`: The real coefficient of the complex number
- `i_num`: The imaginary coefficient of the complex number
- `suffix`: (Optional) The suffix for the imaginary component ("i" or "j", default is "i")

**Returns:** Text - The complex number in the form "x+yi" or "x+yj"

**Examples:**
```swift
let result1 = evaluator.evaluate("=COMPLEX(3, 4)")        // "3+4i"
let result2 = evaluator.evaluate("=COMPLEX(3, -4)")       // "3-4i"
let result3 = evaluator.evaluate("=COMPLEX(0, 1)")        // "i"
let result4 = evaluator.evaluate("=COMPLEX(3, 4, \"j\")") // "3+4j"
```

**Excel Documentation:** [COMPLEX function](https://support.microsoft.com/en-us/office/complex-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128)

**Implementation Status:** ✅ Full implementation

---

### CONVERT

Converts a number from one measurement system to another.

**Syntax:** `CONVERT(number, from_unit, to_unit)`

**Parameters:**
- `number`: The value to convert
- `from_unit`: The unit to convert from
- `to_unit`: The unit to convert to

**Supported Units:**
- **Distance**: "m" (meter), "mi" (mile), "ft" (feet), "in" (inch), "yd" (yard), "ang" (angstrom), "ly" (light-year), "Nmi" (nautical mile), "pica" (pica), "survey_mi" (US survey mile)
- **Weight**: "g" (gram), "kg" (kilogram), "lbm" (pound mass), "u" (atomic mass unit), "ozm" (ounce mass), "ton" (ton), "uk_ton" (UK ton), "cwt" (hundredweight), "stone" (stone), "uk_cwt" (UK hundredweight), "grain" (grain)
- **Time**: "sec" (second), "min" (minute), "hr" (hour), "day" (day), "yr" (year)
- **Temperature**: "C" (Celsius), "F" (Fahrenheit), "K" (Kelvin), "Rank" (Rankine), "Reau" (Réaumur)
- **Pressure**: "Pa" (pascal), "atm" (atmosphere), "mmHg" (millimeter of mercury), "psi" (pounds per square inch), "Torr" (torr)
- **Force**: "N" (newton), "dyn" (dyne), "lbf" (pound force), "pond" (pond)
- **Energy**: "J" (joule), "e" (erg), "cal" (calorie), "eV" (electron volt), "HPh" (horsepower-hour), "Wh" (watt-hour), "flb" (foot-pound), "BTU" (British thermal unit)
- **Power**: "W" (watt), "HP" (horsepower), "PS" (Pferdestärke)
- **Magnetism**: "T" (tesla), "ga" (gauss)
- **Liquid**: "l" (liter), "tsp" (teaspoon), "tbs" (tablespoon), "oz" (fluid ounce), "cup" (cup), "pt" (pint), "qt" (quart), "gal" (gallon), "uk_pt" (UK pint), "uk_qt" (UK quart), "uk_gal" (UK gallon)
- **Bits and Bytes**: "bit" (bit), "byte" (byte)

**Prefixes**: "Y", "Z", "E", "P", "T", "G", "M", "k", "h", "da", "e", "d", "c", "m", "u", "n", "p", "f", "a", "z", "y"

**Returns:** Number - The converted value

**Examples:**
```swift
let result1 = evaluator.evaluate("=CONVERT(1, \"m\", \"ft\")")     // 3.28084
let result2 = evaluator.evaluate("=CONVERT(100, \"C\", \"F\")")    // 212
let result3 = evaluator.evaluate("=CONVERT(1, \"kg\", \"lbm\")")   // 2.20462
let result4 = evaluator.evaluate("=CONVERT(1, \"km\", \"mi\")")    // 0.621371
```

**Excel Documentation:** [CONVERT function](https://support.microsoft.com/en-us/office/convert-function-d785bef1-808e-4aac-bdcd-666c810f9af2)

**Implementation Status:** ✅ Full implementation

---

### DEC2BIN

Converts a decimal number to binary.

**Syntax:** `DEC2BIN(number, [places])`

**Parameters:**
- `number`: The decimal integer you want to convert (must be >= -512 and < 512)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The binary equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=DEC2BIN(10)")        // "1010"
let result2 = evaluator.evaluate("=DEC2BIN(10, 8)")     // "00001010"
let result3 = evaluator.evaluate("=DEC2BIN(255)")       // "11111111"
```

**Excel Documentation:** [DEC2BIN function](https://support.microsoft.com/en-us/office/dec2bin-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838)

**Implementation Status:** ✅ Full implementation

---

### DEC2HEX

Converts a decimal number to hexadecimal.

**Syntax:** `DEC2HEX(number, [places])`

**Parameters:**
- `number`: The decimal integer you want to convert
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The hexadecimal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=DEC2HEX(100)")       // "64"
let result2 = evaluator.evaluate("=DEC2HEX(100, 4)")    // "0064"
let result3 = evaluator.evaluate("=DEC2HEX(255)")       // "FF"
```

**Excel Documentation:** [DEC2HEX function](https://support.microsoft.com/en-us/office/dec2hex-function-6344ee8b-b6b5-4c6a-a672-f64666704619)

**Implementation Status:** ✅ Full implementation

---

### DEC2OCT

Converts a decimal number to octal.

**Syntax:** `DEC2OCT(number, [places])`

**Parameters:**
- `number`: The decimal integer you want to convert
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The octal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=DEC2OCT(8)")         // "10"
let result2 = evaluator.evaluate("=DEC2OCT(8, 4)")      // "0010"
let result3 = evaluator.evaluate("=DEC2OCT(64)")        // "100"
```

**Excel Documentation:** [DEC2OCT function](https://support.microsoft.com/en-us/office/dec2oct-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f)

**Implementation Status:** ✅ Full implementation

---

### DELTA

Tests whether two values are equal. Returns 1 if equal, 0 otherwise.

**Syntax:** `DELTA(number1, [number2])`

**Parameters:**
- `number1`: The first number
- `number2`: (Optional) The second number (defaults to 0)

**Returns:** Number - 1 if equal, 0 if not equal

**Examples:**
```swift
let result1 = evaluator.evaluate("=DELTA(5, 5)")   // 1
let result2 = evaluator.evaluate("=DELTA(5, 3)")   // 0
let result3 = evaluator.evaluate("=DELTA(0)")      // 1
```

**Excel Documentation:** [DELTA function](https://support.microsoft.com/en-us/office/delta-function-2f763672-c959-4e07-ac33-fe03220ba432)

**Implementation Status:** ✅ Full implementation

---

### ERF

Returns the error function integrated between two limits.

**Syntax:** `ERF(lower_limit, [upper_limit])`

**Parameters:**
- `lower_limit`: The lower bound for integrating ERF
- `upper_limit`: (Optional) The upper bound for integrating ERF (if omitted, ERF integrates from 0 to lower_limit)

**Returns:** Number - The error function value

**Examples:**
```swift
let result1 = evaluator.evaluate("=ERF(1)")        // ≈0.8427
let result2 = evaluator.evaluate("=ERF(0, 1)")     // ≈0.8427
let result3 = evaluator.evaluate("=ERF(0.5)")      // ≈0.5205
```

**Excel Documentation:** [ERF function](https://support.microsoft.com/en-us/office/erf-function-c53c7e7b-5482-4b6c-883e-56df3c9af349)

**Implementation Status:** ✅ Full implementation

---

### ERFC

Returns the complementary error function integrated between x and infinity.

**Syntax:** `ERFC(x)`

**Parameters:**
- `x`: The lower bound for integrating ERFC

**Returns:** Number - The complementary error function value (1 - ERF(x))

**Examples:**
```swift
let result1 = evaluator.evaluate("=ERFC(1)")       // ≈0.1573
let result2 = evaluator.evaluate("=ERFC(0)")       // 1
let result3 = evaluator.evaluate("=ERFC(0.5)")     // ≈0.4795
```

**Excel Documentation:** [ERFC function](https://support.microsoft.com/en-us/office/erfc-function-736e0318-70ba-4e8b-8d08-461fe68b71b3)

**Implementation Status:** ✅ Full implementation

---

### GESTEP

Tests whether a number is greater than or equal to a threshold value. Returns 1 if true, 0 otherwise.

**Syntax:** `GESTEP(number, [step])`

**Parameters:**
- `number`: The value to test against step
- `step`: (Optional) The threshold value (defaults to 0)

**Returns:** Number - 1 if number >= step, 0 otherwise

**Examples:**
```swift
let result1 = evaluator.evaluate("=GESTEP(5, 3)")   // 1
let result2 = evaluator.evaluate("=GESTEP(3, 5)")   // 0
let result3 = evaluator.evaluate("=GESTEP(5, 5)")   // 1
let result4 = evaluator.evaluate("=GESTEP(5)")      // 1
```

**Excel Documentation:** [GESTEP function](https://support.microsoft.com/en-us/office/gestep-function-f37e7d2a-41da-4129-be95-640883fca9df)

**Implementation Status:** ✅ Full implementation

---

### HEX2BIN

Converts a hexadecimal number to binary.

**Syntax:** `HEX2BIN(number, [places])`

**Parameters:**
- `number`: The hexadecimal number (string) you want to convert (max 10 characters)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The binary equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=HEX2BIN(\"A\")")        // "1010"
let result2 = evaluator.evaluate("=HEX2BIN(\"A\", 8)")     // "00001010"
let result3 = evaluator.evaluate("=HEX2BIN(\"FF\")")       // "11111111"
```

**Excel Documentation:** [HEX2BIN function](https://support.microsoft.com/en-us/office/hex2bin-function-a13aafaa-5737-4920-8424-643e581828c1)

**Implementation Status:** ✅ Full implementation

---

### HEX2DEC

Converts a hexadecimal number to decimal.

**Syntax:** `HEX2DEC(number)`

**Parameters:**
- `number`: The hexadecimal number (string) you want to convert (max 10 characters)

**Returns:** Number - The decimal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=HEX2DEC(\"A\")")        // 10
let result2 = evaluator.evaluate("=HEX2DEC(\"FF\")")       // 255
let result3 = evaluator.evaluate("=HEX2DEC(\"64\")")       // 100
```

**Excel Documentation:** [HEX2DEC function](https://support.microsoft.com/en-us/office/hex2dec-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e)

**Implementation Status:** ✅ Full implementation

---

### HEX2OCT

Converts a hexadecimal number to octal.

**Syntax:** `HEX2OCT(number, [places])`

**Parameters:**
- `number`: The hexadecimal number (string) you want to convert (max 10 characters)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The octal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=HEX2OCT(\"A\")")        // "12"
let result2 = evaluator.evaluate("=HEX2OCT(\"A\", 4)")     // "0012"
let result3 = evaluator.evaluate("=HEX2OCT(\"FF\")")       // "377"
```

**Excel Documentation:** [HEX2OCT function](https://support.microsoft.com/en-us/office/hex2oct-function-54d52808-5d19-4bd0-8a63-1096a5d11912)

**Implementation Status:** ✅ Full implementation

---

### IMABS

Returns the absolute value (modulus) of a complex number.

**Syntax:** `IMABS(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Number - The absolute value |z| = √(x² + y²)

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMABS(\"3+4i\")")    // 5
let result2 = evaluator.evaluate("=IMABS(\"5-12i\")")   // 13
let result3 = evaluator.evaluate("=IMABS(\"1+i\")")     // ≈1.4142
```

**Excel Documentation:** [IMABS function](https://support.microsoft.com/en-us/office/imabs-function-b31e73c6-d90c-4062-90bc-8eb351d765a1)

**Implementation Status:** ✅ Full implementation

---

### IMACOSH

Returns the inverse hyperbolic cosine of a complex number.

**Syntax:** `IMACOSH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The inverse hyperbolic cosine as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMACOSH(\"2+3i\")")
```

**Excel Documentation:** [IMACOSH function](https://support.microsoft.com/en-us/office/imacosh-function-9b17595f-2cce-4bce-9420-1508f1a05531)

**Implementation Status:** ✅ Full implementation

---

### IMADD

Returns the sum of two or more complex numbers.

**Syntax:** `IMADD(inumber1, inumber2, ...)`

**Parameters:**
- `inumber1, inumber2, ...`: Complex numbers in the form "x+yi" or "x+yj"

**Returns:** Text - The sum as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMADD(\"3+4i\", \"5+2i\")")       // "8+6i"
let result2 = evaluator.evaluate("=IMADD(\"1+i\", \"2-i\", \"3\")")  // "6+0i"
```

**Excel Documentation:** [IMADD function](https://support.microsoft.com/en-us/office/imadd-function-d1a57e8d-3f49-4e0a-9c41-1b6cd0b3f6bf)

**Implementation Status:** ✅ Full implementation

---

### IMAGINARY

Returns the imaginary coefficient of a complex number.

**Syntax:** `IMAGINARY(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Number - The imaginary coefficient

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMAGINARY(\"3+4i\")")   // 4
let result2 = evaluator.evaluate("=IMAGINARY(\"5-2i\")")   // -2
let result3 = evaluator.evaluate("=IMAGINARY(\"7\")")      // 0
```

**Excel Documentation:** [IMAGINARY function](https://support.microsoft.com/en-us/office/imaginary-function-dd5952fd-473d-44d9-95a1-9a17b23e428a)

**Implementation Status:** ✅ Full implementation

---

### IMARGUMENT

Returns the argument (angle) θ of a complex number in radians.

**Syntax:** `IMARGUMENT(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Number - The argument in radians (range: -π to π)

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMARGUMENT(\"3+4i\")")   // ≈0.9273 (≈53.13°)
let result2 = evaluator.evaluate("=IMARGUMENT(\"1+i\")")    // ≈0.7854 (π/4)
let result3 = evaluator.evaluate("=IMARGUMENT(\"-1\")")     // ≈3.1416 (π)
```

**Excel Documentation:** [IMARGUMENT function](https://support.microsoft.com/en-us/office/imargument-function-eed89c90-fc9f-44a9-9622-9075b9befd17)

**Implementation Status:** ✅ Full implementation

---

### IMASINH

Returns the inverse hyperbolic sine of a complex number.

**Syntax:** `IMASINH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The inverse hyperbolic sine as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMASINH(\"2+3i\")")
```

**Excel Documentation:** [IMASINH function](https://support.microsoft.com/en-us/office/imasinh-function-5c4c5f37-8a87-4155-9b08-5e09f2d5f0f0)

**Implementation Status:** ✅ Full implementation

---

### IMATANH

Returns the inverse hyperbolic tangent of a complex number.

**Syntax:** `IMATANH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The inverse hyperbolic tangent as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMATANH(\"0.5+0.5i\")")
```

**Excel Documentation:** [IMATANH function](https://support.microsoft.com/en-us/office/imatanh-function-fb4a8f91-7b0b-4b1f-9ce6-f8a6ef6e9d96)

**Implementation Status:** ✅ Full implementation

---

### IMCONJUGATE

Returns the complex conjugate of a complex number.

**Syntax:** `IMCONJUGATE(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The complex conjugate (x-yi)

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMCONJUGATE(\"3+4i\")")   // "3-4i"
let result2 = evaluator.evaluate("=IMCONJUGATE(\"5-2i\")")   // "5+2i"
let result3 = evaluator.evaluate("=IMCONJUGATE(\"7\")")      // "7"
```

**Excel Documentation:** [IMCONJUGATE function](https://support.microsoft.com/en-us/office/imconjugate-function-124983e4-b0e6-471f-80de-01e95cf7fb6f)

**Implementation Status:** ✅ Full implementation

---

### IMCOS

Returns the cosine of a complex number.

**Syntax:** `IMCOS(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The cosine as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMCOS(\"1+i\")")
```

**Excel Documentation:** [IMCOS function](https://support.microsoft.com/en-us/office/imcos-function-dad75277-f592-4a6b-ad6c-be93a808a53c)

**Implementation Status:** ✅ Full implementation

---

### IMCOSH

Returns the hyperbolic cosine of a complex number.

**Syntax:** `IMCOSH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The hyperbolic cosine as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMCOSH(\"1+i\")")
```

**Excel Documentation:** [IMCOSH function](https://support.microsoft.com/en-us/office/imcosh-function-053e4ddb-4122-458b-be9a-457c405e90ff)

**Implementation Status:** ✅ Full implementation

---

### IMCSC

Returns the cosecant of a complex number.

**Syntax:** `IMCSC(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The cosecant as a complex number (1/sin(z))

**Examples:**
```swift
let result = evaluator.evaluate("=IMCSC(\"1+i\")")
```

**Excel Documentation:** [IMCSC function](https://support.microsoft.com/en-us/office/imcsc-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323)

**Implementation Status:** ✅ Full implementation

---

### IMCSCH

Returns the hyperbolic cosecant of a complex number.

**Syntax:** `IMCSCH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The hyperbolic cosecant as a complex number (1/sinh(z))

**Examples:**
```swift
let result = evaluator.evaluate("=IMCSCH(\"1+i\")")
```

**Excel Documentation:** [IMCSCH function](https://support.microsoft.com/en-us/office/imcsch-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9)

**Implementation Status:** ✅ Full implementation

---

### IMDIV

Returns the quotient of two complex numbers.

**Syntax:** `IMDIV(inumber1, inumber2)`

**Parameters:**
- `inumber1`: The complex numerator
- `inumber2`: The complex denominator

**Returns:** Text - The quotient as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMDIV(\"8+6i\", \"2+i\")")    // "4+i"
let result2 = evaluator.evaluate("=IMDIV(\"1+i\", \"1-i\")")     // "i"
```

**Excel Documentation:** [IMDIV function](https://support.microsoft.com/en-us/office/imdiv-function-a505aff7-af8a-4451-8142-77ec3d74d83f)

**Implementation Status:** ✅ Full implementation

---

### IMEXP

Returns the exponential of a complex number (e^z).

**Syntax:** `IMEXP(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The exponential as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMEXP(\"i\")")       // ≈"0.5403+0.8415i"
let result2 = evaluator.evaluate("=IMEXP(\"1+i\")")     // ≈"1.4686+2.2874i"
```

**Excel Documentation:** [IMEXP function](https://support.microsoft.com/en-us/office/imexp-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f)

**Implementation Status:** ✅ Full implementation

---

### IMLN

Returns the natural logarithm of a complex number.

**Syntax:** `IMLN(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The natural logarithm as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMLN(\"i\")")        // ≈"1.5708i"
let result2 = evaluator.evaluate("=IMLN(\"3+4i\")")     // ≈"1.6094+0.9273i"
```

**Excel Documentation:** [IMLN function](https://support.microsoft.com/en-us/office/imln-function-32b98bcf-8b81-437c-a636-6fb3aad509d8)

**Implementation Status:** ✅ Full implementation

---

### IMLOG10

Returns the base-10 logarithm of a complex number.

**Syntax:** `IMLOG10(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The base-10 logarithm as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMLOG10(\"3+4i\")")
```

**Excel Documentation:** [IMLOG10 function](https://support.microsoft.com/en-us/office/imlog10-function-58200fca-e2a2-4271-8a98-ccd4360213a5)

**Implementation Status:** ✅ Full implementation

---

### IMLOG2

Returns the base-2 logarithm of a complex number.

**Syntax:** `IMLOG2(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The base-2 logarithm as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMLOG2(\"3+4i\")")
```

**Excel Documentation:** [IMLOG2 function](https://support.microsoft.com/en-us/office/imlog2-function-152e13b4-bc79-486c-a243-e6a676878c51)

**Implementation Status:** ✅ Full implementation

---

### IMPOWER

Returns a complex number raised to an integer power.

**Syntax:** `IMPOWER(inumber, number)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"
- `number`: The power to which you want to raise the complex number

**Returns:** Text - The result as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMPOWER(\"2+3i\", 2)")   // "-5+12i"
let result2 = evaluator.evaluate("=IMPOWER(\"i\", 2)")      // "-1"
let result3 = evaluator.evaluate("=IMPOWER(\"1+i\", 3)")    // "-2+2i"
```

**Excel Documentation:** [IMPOWER function](https://support.microsoft.com/en-us/office/impower-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39)

**Implementation Status:** ✅ Full implementation

---

### IMPRODUCT

Returns the product of two or more complex numbers.

**Syntax:** `IMPRODUCT(inumber1, inumber2, ...)`

**Parameters:**
- `inumber1, inumber2, ...`: Complex numbers in the form "x+yi" or "x+yj"

**Returns:** Text - The product as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMPRODUCT(\"3+4i\", \"5+2i\")")   // "7+26i"
let result2 = evaluator.evaluate("=IMPRODUCT(\"i\", \"i\")")         // "-1"
```

**Excel Documentation:** [IMPRODUCT function](https://support.microsoft.com/en-us/office/improduct-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba)

**Implementation Status:** ✅ Full implementation

---

### IMREAL

Returns the real coefficient of a complex number.

**Syntax:** `IMREAL(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Number - The real coefficient

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMREAL(\"3+4i\")")   // 3
let result2 = evaluator.evaluate("=IMREAL(\"5-2i\")")   // 5
let result3 = evaluator.evaluate("=IMREAL(\"i\")")      // 0
```

**Excel Documentation:** [IMREAL function](https://support.microsoft.com/en-us/office/imreal-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366)

**Implementation Status:** ✅ Full implementation

---

### IMSEC

Returns the secant of a complex number.

**Syntax:** `IMSEC(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The secant as a complex number (1/cos(z))

**Examples:**
```swift
let result = evaluator.evaluate("=IMSEC(\"1+i\")")
```

**Excel Documentation:** [IMSEC function](https://support.microsoft.com/en-us/office/imsec-function-6df11132-4411-4df4-a3dc-1f17372459e0)

**Implementation Status:** ✅ Full implementation

---

### IMSECH

Returns the hyperbolic secant of a complex number.

**Syntax:** `IMSECH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The hyperbolic secant as a complex number (1/cosh(z))

**Examples:**
```swift
let result = evaluator.evaluate("=IMSECH(\"1+i\")")
```

**Excel Documentation:** [IMSECH function](https://support.microsoft.com/en-us/office/imsech-function-f250304f-788b-4505-954e-eb01fa50903b)

**Implementation Status:** ✅ Full implementation

---

### IMSIN

Returns the sine of a complex number.

**Syntax:** `IMSIN(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The sine as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMSIN(\"1+i\")")
```

**Excel Documentation:** [IMSIN function](https://support.microsoft.com/en-us/office/imsin-function-1ab02a39-a721-48de-82ef-f52bf37859f6)

**Implementation Status:** ✅ Full implementation

---

### IMSINH

Returns the hyperbolic sine of a complex number.

**Syntax:** `IMSINH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The hyperbolic sine as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMSINH(\"1+i\")")
```

**Excel Documentation:** [IMSINH function](https://support.microsoft.com/en-us/office/imsinh-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d)

**Implementation Status:** ✅ Full implementation

---

### IMSQRT

Returns the square root of a complex number.

**Syntax:** `IMSQRT(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The square root as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMSQRT(\"-1\")")         // "i"
let result2 = evaluator.evaluate("=IMSQRT(\"3+4i\")")       // "2+i"
let result3 = evaluator.evaluate("=IMSQRT(\"-8+6i\")")      // ≈"1+3i"
```

**Excel Documentation:** [IMSQRT function](https://support.microsoft.com/en-us/office/imsqrt-function-e1753f80-ba11-4664-a10e-e17368396b70)

**Implementation Status:** ✅ Full implementation

---

### IMSUB

Returns the difference of two complex numbers.

**Syntax:** `IMSUB(inumber1, inumber2)`

**Parameters:**
- `inumber1`: The complex number from which to subtract inumber2
- `inumber2`: The complex number to subtract from inumber1

**Returns:** Text - The difference as a complex number

**Examples:**
```swift
let result1 = evaluator.evaluate("=IMSUB(\"8+6i\", \"3+2i\")")   // "5+4i"
let result2 = evaluator.evaluate("=IMSUB(\"1+i\", \"i\")")       // "1"
```

**Excel Documentation:** [IMSUB function](https://support.microsoft.com/en-us/office/imsub-function-2e404b4d-4935-4e85-9f52-cb08b9a45054)

**Implementation Status:** ✅ Full implementation

---

### IMSUM

Returns the sum of multiple complex numbers.

**Syntax:** `IMSUM(inumber1, inumber2, ...)`

**Parameters:**
- `inumber1, inumber2, ...`: Complex numbers in the form "x+yi" or "x+yj"

**Returns:** Text - The sum as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMSUM(\"3+4i\", \"5+2i\", \"1-i\")")  // "9+5i"
```

**Excel Documentation:** [IMSUM function](https://support.microsoft.com/en-us/office/imsum-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f)

**Implementation Status:** ✅ Full implementation

---

### IMTAN

Returns the tangent of a complex number.

**Syntax:** `IMTAN(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The tangent as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMTAN(\"1+i\")")
```

**Excel Documentation:** [IMTAN function](https://support.microsoft.com/en-us/office/imtan-function-8478f45d-610a-43cf-8544-9fc0b553a132)

**Implementation Status:** ✅ Full implementation

---

### IMTANH

Returns the hyperbolic tangent of a complex number.

**Syntax:** `IMTANH(inumber)`

**Parameters:**
- `inumber`: A complex number in the form "x+yi" or "x+yj"

**Returns:** Text - The hyperbolic tangent as a complex number

**Examples:**
```swift
let result = evaluator.evaluate("=IMTANH(\"1+i\")")
```

**Excel Documentation:** [IMTANH function](https://support.microsoft.com/en-us/office/imtanh-function-1e648c9f-9b40-4d97-a51c-d26e8f8ed6f3)

**Implementation Status:** ✅ Full implementation

---

### OCT2BIN

Converts an octal number to binary.

**Syntax:** `OCT2BIN(number, [places])`

**Parameters:**
- `number`: The octal number (string) you want to convert (max 10 characters)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The binary equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=OCT2BIN(\"12\")")       // "1010"
let result2 = evaluator.evaluate("=OCT2BIN(\"12\", 8)")    // "00001010"
let result3 = evaluator.evaluate("=OCT2BIN(\"377\")")      // "11111111"
```

**Excel Documentation:** [OCT2BIN function](https://support.microsoft.com/en-us/office/oct2bin-function-55383471-3c56-4d27-9522-1a8ec646c589)

**Implementation Status:** ✅ Full implementation

---

### OCT2DEC

Converts an octal number to decimal.

**Syntax:** `OCT2DEC(number)`

**Parameters:**
- `number`: The octal number (string) you want to convert (max 10 characters)

**Returns:** Number - The decimal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=OCT2DEC(\"12\")")       // 10
let result2 = evaluator.evaluate("=OCT2DEC(\"377\")")      // 255
let result3 = evaluator.evaluate("=OCT2DEC(\"100\")")      // 64
```

**Excel Documentation:** [OCT2DEC function](https://support.microsoft.com/en-us/office/oct2dec-function-87606014-cb98-44b2-8dbb-e48f8ced1554)

**Implementation Status:** ✅ Full implementation

---

### OCT2HEX

Converts an octal number to hexadecimal.

**Syntax:** `OCT2HEX(number, [places])`

**Parameters:**
- `number`: The octal number (string) you want to convert (max 10 characters)
- `places`: (Optional) The number of characters to use (pads with leading zeros)

**Returns:** Text - The hexadecimal equivalent

**Examples:**
```swift
let result1 = evaluator.evaluate("=OCT2HEX(\"12\")")       // "A"
let result2 = evaluator.evaluate("=OCT2HEX(\"12\", 4)")    // "000A"
let result3 = evaluator.evaluate("=OCT2HEX(\"377\")")      // "FF"
```

**Excel Documentation:** [OCT2HEX function](https://support.microsoft.com/en-us/office/oct2hex-function-912175b4-d497-41b4-a029-221f051b858f)

**Implementation Status:** ✅ Full implementation

---

## See Also

- <doc:Mathematical> - Mathematical and trigonometric functions
- <doc:Statistical> - Statistical analysis functions
- <doc:Financial> - Financial calculation functions
- ``FormulaEvaluator`` - Programmatic formula evaluation
- <doc:WritingWorkbooks> - Creating formulas in workbooks
