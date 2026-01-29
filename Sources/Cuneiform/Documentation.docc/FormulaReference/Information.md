# Information Functions

Type checking, data inspection, and cell information functions.

## Overview

Information functions provide detailed information about values, cells, and the workbook environment. These functions are essential for validating data types, inspecting cell properties, checking for errors, and building robust formulas that adapt to different data conditions.

Cuneiform implements 28 information functions compatible with Excel, enabling sophisticated data validation, conditional logic, and dynamic formula construction based on cell characteristics.

### Quick Reference

| Function | Description | Status |
|----------|-------------|--------|
| **ISBLANK** | Tests if a value is blank | ✅ Full |
| **ISERR** | Tests if a value is any error except #N/A | ✅ Full |
| **ISERROR** | Tests if a value is any error | ✅ Full |
| **ISEVEN** | Tests if a number is even | ✅ Full |
| **ISFORMULA** | Tests if a cell contains a formula | 🔄 Stub |
| **ISLOGICAL** | Tests if a value is a logical value | ✅ Full |
| **ISNA** | Tests if a value is #N/A error | ✅ Full |
| **ISNONTEXT** | Tests if a value is not text | ✅ Full |
| **ISNUMBER** | Tests if a value is a number | ✅ Full |
| **ISODD** | Tests if a number is odd | ✅ Full |
| **ISOMITTED** | Tests if a value is omitted (LAMBDA) | ⚠️ Partial |
| **ISREF** | Tests if a value is a reference | ✅ Full |
| **ISTEXT** | Tests if a value is text | ✅ Full |
| **N** | Converts value to number | ✅ Full |
| **NA** | Returns #N/A error | ✅ Full |
| **TYPE** | Returns type code for a value | ✅ Full |
| **CELL** | Returns information about a cell | ⚠️ Partial |
| **INFO** | Returns information about environment | ⚠️ Partial |
| **SHEET** | Returns sheet number | ⚠️ Partial |
| **SHEETS** | Returns number of sheets | ⚠️ Partial |
| **ROW** | Returns row number of reference | ✅ Full |
| **COLUMN** | Returns column number of reference | ✅ Full |
| **ROWS** | Returns number of rows in reference | ✅ Full |
| **COLUMNS** | Returns number of columns in reference | ✅ Full |
| **ADDRESS** | Creates cell address from row/column | ✅ Full |
| **AREAS** | Returns number of areas in reference | ⚠️ Partial |
| **FORMULATEXT** | Returns formula as text | 🔄 Stub |
| **ERROR.TYPE** | Returns error type number | ✅ Full |

## Type Checking Functions

### ISBLANK

Tests if a value or cell is blank.

**Syntax:** `ISBLANK(value)`

**Parameters:**
- `value`: The value or cell reference to test

**Returns:** Boolean TRUE if blank, FALSE otherwise

**Examples:**
```swift
// Test if cell is empty
let isEmpty = try evaluator.evaluate("=ISBLANK(A1)")

// Check for blank cells in validation
let isValid = try evaluator.evaluate("=IF(ISBLANK(A1), \"Required\", \"OK\")")

// Count blank cells with COUNTIF
let blanks = try evaluator.evaluate("=SUMPRODUCT(--ISBLANK(A1:A10))")

// Conditional formatting based on blank
let formatted = try evaluator.evaluate("=NOT(ISBLANK(B1))")
```

**Notes:**
- Returns TRUE for empty cells only
- Empty string ("") is considered blank
- Cells with spaces are not blank
- Different from testing `A1=""` which treats 0 as blank

**Excel Documentation:** [ISBLANK function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISERROR

Tests if a value is any error value.

**Syntax:** `ISERROR(value)`

**Parameters:**
- `value`: The value to test for error

**Returns:** Boolean TRUE if any error type, FALSE otherwise

**Examples:**
```swift
// Test if calculation results in error
let hasError = try evaluator.evaluate("=ISERROR(A1/B1)")

// Combine with IF to handle errors
let safe = try evaluator.evaluate("=IF(ISERROR(VLOOKUP(A1, Table, 2)), \"Not Found\", VLOOKUP(A1, Table, 2))")

// Check array for any errors
let checkRange = try evaluator.evaluate("=SUMPRODUCT(--ISERROR(A1:A10))")

// Validate before calculation
let result = try evaluator.evaluate("=IF(ISERROR(A1), 0, A1*2)")
```

**Notes:**
- Returns TRUE for all error types: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
- Use ISERR to exclude #N/A errors
- Use ISNA to test only for #N/A errors
- Consider IFERROR function for simpler error handling

**Excel Documentation:** [ISERROR function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISERR

Tests if a value is any error except #N/A.

**Syntax:** `ISERR(value)`

**Parameters:**
- `value`: The value to test for error

**Returns:** Boolean TRUE if error (excluding #N/A), FALSE otherwise

**Examples:**
```swift
// Test for errors except #N/A
let hasError = try evaluator.evaluate("=ISERR(A1/B1)")

// Allow #N/A but catch other errors
let lookup = try evaluator.evaluate("=IF(ISERR(VLOOKUP(A1, Table, 2)), \"Error\", VLOOKUP(A1, Table, 2))")

// Distinguish between #N/A and other errors
let status = try evaluator.evaluate("=IF(ISNA(A1), \"Not Found\", IF(ISERR(A1), \"Error\", A1))")

// Validate calculation errors only
let check = try evaluator.evaluate("=NOT(ISERR(A1*B1))")
```

**Notes:**
- Returns TRUE for: #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
- Returns FALSE for: #N/A error
- Useful when #N/A is expected (lookups) but other errors indicate problems
- More specific than ISERROR

**Excel Documentation:** [ISERR function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISNA

Tests if a value is the #N/A error.

**Syntax:** `ISNA(value)`

**Parameters:**
- `value`: The value to test for #N/A error

**Returns:** Boolean TRUE if #N/A, FALSE otherwise

**Examples:**
```swift
// Test if lookup returned #N/A
let notFound = try evaluator.evaluate("=ISNA(VLOOKUP(A1, Table, 2, FALSE))")

// Handle #N/A differently from other errors
let result = try evaluator.evaluate("=IF(ISNA(A1), \"Not in list\", IF(ISERROR(A1), \"Error\", A1))")

// Count #N/A errors in range
let naCount = try evaluator.evaluate("=SUMPRODUCT(--ISNA(A1:A10))")

// Suppress only #N/A errors
let display = try evaluator.evaluate("=IF(ISNA(MATCH(A1, B:B, 0)), \"\", MATCH(A1, B:B, 0))")
```

**Notes:**
- Returns TRUE only for #N/A error
- Returns FALSE for all other values including other errors
- Commonly used with lookup functions (VLOOKUP, MATCH, XLOOKUP)
- Consider IFNA function for simpler #N/A handling

**Excel Documentation:** [ISNA function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISREF

Tests if a value is a reference.

**Syntax:** `ISREF(value)`

**Parameters:**
- `value`: The value to test

**Returns:** Boolean TRUE if reference, FALSE otherwise

**Examples:**
```swift
// Test if argument is a cell reference
let isRef = try evaluator.evaluate("=ISREF(A1)")  // TRUE

// Test if named range is a reference
let isRange = try evaluator.evaluate("=ISREF(MyRange)")  // TRUE

// Constants are not references
let notRef = try evaluator.evaluate("=ISREF(100)")  // FALSE

// Check formula result
let checkResult = try evaluator.evaluate("=ISREF(SUM(A1:A10))")  // FALSE
```

**Notes:**
- Returns TRUE for cell references (A1, $B$2)
- Returns TRUE for range references (A1:B10)
- Returns TRUE for named ranges that refer to cells
- Returns FALSE for values, even if from a cell reference
- Tests the argument itself, not its evaluated value

**Excel Documentation:** [ISREF function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISTEXT

Tests if a value is text.

**Syntax:** `ISTEXT(value)`

**Parameters:**
- `value`: The value to test

**Returns:** Boolean TRUE if text, FALSE otherwise

**Examples:**
```swift
// Test if cell contains text
let hasText = try evaluator.evaluate("=ISTEXT(A1)")

// Validate text input
let isValid = try evaluator.evaluate("=IF(ISTEXT(A1), \"Valid\", \"Enter text\")")

// Count text cells
let textCount = try evaluator.evaluate("=SUMPRODUCT(--ISTEXT(A1:A10))")

// Filter text values
let filtered = try evaluator.evaluate("=IF(ISTEXT(A1), A1, \"\")")
```

**Notes:**
- Returns TRUE only for text values
- Returns FALSE for numbers, dates, booleans, errors, and blank cells
- Empty string ("") returns TRUE
- Use ISNONTEXT for opposite test

**Excel Documentation:** [ISTEXT function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISNUMBER

Tests if a value is a number.

**Syntax:** `ISNUMBER(value)`

**Parameters:**
- `value`: The value to test

**Returns:** Boolean TRUE if number, FALSE otherwise

**Examples:**
```swift
// Test if cell contains number
let hasNumber = try evaluator.evaluate("=ISNUMBER(A1)")

// Validate numeric input
let isValid = try evaluator.evaluate("=IF(ISNUMBER(A1), A1*2, \"Enter number\")")

// Count numeric cells
let numberCount = try evaluator.evaluate("=SUMPRODUCT(--ISNUMBER(A1:A10))")

// Check if SEARCH found text (returns number)
let found = try evaluator.evaluate("=ISNUMBER(SEARCH(\"excel\", A1))")
```

**Notes:**
- Returns TRUE for all numeric values including dates/times
- Returns FALSE for text, booleans, errors, and blank cells
- Dates are numbers so ISNUMBER returns TRUE for dates
- Commonly used to test SEARCH/FIND results

**Excel Documentation:** [ISNUMBER function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISLOGICAL

Tests if a value is a logical value (TRUE or FALSE).

**Syntax:** `ISLOGICAL(value)`

**Parameters:**
- `value`: The value to test

**Returns:** Boolean TRUE if logical value, FALSE otherwise

**Examples:**
```swift
// Test if cell contains boolean
let isBoolean = try evaluator.evaluate("=ISLOGICAL(A1)")

// Test formula result
let checkResult = try evaluator.evaluate("=ISLOGICAL(A1>100)")  // TRUE

// Validate boolean input
let isValid = try evaluator.evaluate("=IF(ISLOGICAL(A1), \"Valid\", \"Enter TRUE or FALSE\")")

// Count logical values
let boolCount = try evaluator.evaluate("=SUMPRODUCT(--ISLOGICAL(A1:A10))")
```

**Notes:**
- Returns TRUE only for TRUE and FALSE values
- Returns FALSE for 1/0, "TRUE"/"FALSE" text, and all other values
- Formulas that return TRUE/FALSE will show TRUE for ISLOGICAL

**Excel Documentation:** [ISLOGICAL function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISNONTEXT

Tests if a value is not text.

**Syntax:** `ISNONTEXT(value)`

**Parameters:**
- `value`: The value to test

**Returns:** Boolean TRUE if not text, FALSE if text

**Examples:**
```swift
// Test if cell is not text
let notText = try evaluator.evaluate("=ISNONTEXT(A1)")

// Accept numbers and dates only
let isValid = try evaluator.evaluate("=IF(ISNONTEXT(A1), A1, \"Must be numeric\")")

// Count non-text cells (including blanks)
let nonTextCount = try evaluator.evaluate("=SUMPRODUCT(--ISNONTEXT(A1:A10))")

// Filter numeric and blank values
let filtered = try evaluator.evaluate("=IF(ISNONTEXT(A1), A1, 0)")
```

**Notes:**
- Returns TRUE for numbers, booleans, errors, and blank cells
- Returns FALSE only for text values (including empty string)
- Opposite of ISTEXT
- Blank cells return TRUE

**Excel Documentation:** [ISNONTEXT function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665)

**Implementation Status:** ✅ Full implementation

---

### ISEVEN

Tests if a number is even.

**Syntax:** `ISEVEN(number)`

**Parameters:**
- `number`: The numeric value to test

**Returns:** Boolean TRUE if even, FALSE if odd

**Examples:**
```swift
// Test if number is even
let even = try evaluator.evaluate("=ISEVEN(4)")  // TRUE

// Conditional formatting for even rows
let rowColor = try evaluator.evaluate("=ISEVEN(ROW())")

// Alternate row processing
let process = try evaluator.evaluate("=IF(ISEVEN(ROW()), A1*2, A1)")

// Validate even input
let isValid = try evaluator.evaluate("=IF(ISEVEN(A1), A1, \"Must be even\")")
```

**Notes:**
- Returns TRUE for even numbers (divisible by 2)
- Returns FALSE for odd numbers
- Returns #VALUE! error for non-numeric values
- Truncates decimals before testing: ISEVEN(3.9) returns FALSE

**Excel Documentation:** [ISEVEN function](https://support.microsoft.com/en-us/office/iseven-function-aa15929a-d77b-4fbb-92f4-2f479af55356)

**Implementation Status:** ✅ Full implementation

---

### ISODD

Tests if a number is odd.

**Syntax:** `ISODD(number)`

**Parameters:**
- `number`: The numeric value to test

**Returns:** Boolean TRUE if odd, FALSE if even

**Examples:**
```swift
// Test if number is odd
let odd = try evaluator.evaluate("=ISODD(5)")  // TRUE

// Conditional formatting for odd rows
let rowColor = try evaluator.evaluate("=ISODD(ROW())")

// Alternate row processing
let process = try evaluator.evaluate("=IF(ISODD(ROW()), A1, A1*2)")

// Validate odd input
let isValid = try evaluator.evaluate("=IF(ISODD(A1), A1, \"Must be odd\")")
```

**Notes:**
- Returns TRUE for odd numbers (not divisible by 2)
- Returns FALSE for even numbers
- Returns #VALUE! error for non-numeric values
- Truncates decimals before testing: ISODD(4.9) returns TRUE

**Excel Documentation:** [ISODD function](https://support.microsoft.com/en-us/office/isodd-function-1208a56d-4f10-4f44-a5fc-648cafd6c07a)

**Implementation Status:** ✅ Full implementation

---

### ISFORMULA

Tests if a cell contains a formula.

**Syntax:** `ISFORMULA(reference)`

**Parameters:**
- `reference`: Cell reference to test

**Returns:** Boolean TRUE if cell contains formula, FALSE otherwise

**Examples:**
```swift
// Test if cell has formula
let hasFormula = try evaluator.evaluate("=ISFORMULA(A1)")

// Find formula cells
let isFormulaCell = try evaluator.evaluate("=IF(ISFORMULA(A1), \"Formula\", \"Value\")")

// Count formulas in range
let formulaCount = try evaluator.evaluate("=SUMPRODUCT(--ISFORMULA(A1:A10))")

// Conditional formatting for formulas
let highlight = try evaluator.evaluate("=ISFORMULA(A1)")
```

**Notes:**
- Returns TRUE if cell contains a formula
- Returns FALSE for values, text, or blank cells
- Can be used for auditing and validation
- Implementation is simplified

**Excel Documentation:** [ISFORMULA function](https://support.microsoft.com/en-us/office/isformula-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5)

**Implementation Status:** 🔄 Stub (always returns FALSE in current implementation)

---

### ISOMITTED

Tests if a value in a LAMBDA function is omitted.

**Syntax:** `ISOMITTED(argument)`

**Parameters:**
- `argument`: The argument to test

**Returns:** Boolean TRUE if argument is omitted, FALSE otherwise

**Examples:**
```swift
// Used within LAMBDA for optional parameters
let lambda = try evaluator.evaluate("=LAMBDA(x, y, IF(ISOMITTED(y), x*2, x*y))")

// Provide default value for omitted parameter
let withDefault = try evaluator.evaluate("=IF(ISOMITTED(A1), 10, A1)")

// Check multiple optional parameters
let multiParam = try evaluator.evaluate("=LAMBDA(a, b, c, IF(ISOMITTED(c), a+b, a+b+c))")
```

**Notes:**
- Primarily used with LAMBDA functions
- Allows creating functions with optional parameters
- Returns FALSE for any provided argument (including 0, blank, or FALSE)
- Simplified implementation as full LAMBDA support is limited

**Excel Documentation:** [ISOMITTED function](https://support.microsoft.com/en-us/office/isomitted-function-831c59a5-c2f9-4df7-ba8a-e84cd0c94a21)

**Implementation Status:** ⚠️ Partial (limited LAMBDA support)

---

## Type Information Functions

### TYPE

Returns a number indicating the data type of a value.

**Syntax:** `TYPE(value)`

**Parameters:**
- `value`: The value to determine the type of

**Returns:** Number indicating type:
- 1 = Number
- 2 = Text
- 4 = Boolean (TRUE/FALSE)
- 16 = Error value
- 64 = Array

**Examples:**
```swift
// Determine value type
let type = try evaluator.evaluate("=TYPE(A1)")

// Branch based on type
let result = try evaluator.evaluate("""
    =CHOOSE(TYPE(A1), 
        "Number", 
        "Text", 
        , 
        "Boolean", 
        REPT(",", 11), 
        "Error")
    """)

// Validate specific type
let isNumber = try evaluator.evaluate("=TYPE(A1)=1")

// Handle different types differently
let process = try evaluator.evaluate("=IF(TYPE(A1)=2, UPPER(A1), A1*2)")
```

**Notes:**
- Returns numeric codes, not text descriptions
- Type 3 is obsolete (Lotus 1-2-3 compatibility)
- Array type (64) returned for range/array values
- Useful for building type-flexible formulas

**Excel Documentation:** [TYPE function](https://support.microsoft.com/en-us/office/type-function-45b4e688-4bc3-48b3-a105-ffa892995899)

**Implementation Status:** ✅ Full implementation

---

### N

Converts a value to a number.

**Syntax:** `N(value)`

**Parameters:**
- `value`: The value to convert to number

**Returns:** Number value based on input:
- Numbers return as-is
- TRUE returns 1, FALSE returns 0
- Dates return as serial numbers
- Text and blank return 0
- Errors return the error

**Examples:**
```swift
// Convert boolean to number
let num = try evaluator.evaluate("=N(TRUE)")  // 1

// Convert date to serial number
let serial = try evaluator.evaluate("=N(\"1/1/2024\")")

// Sum boolean values
let trueCount = try evaluator.evaluate("=SUM(N(A1>100), N(B1>100), N(C1>100))")

// Force numeric conversion
let converted = try evaluator.evaluate("=N(A1) + N(B1)")
```

**Notes:**
- Mainly for compatibility with other spreadsheet programs
- Numbers pass through unchanged
- Booleans convert to 1 (TRUE) or 0 (FALSE)
- Text returns 0 (not text-to-number conversion)
- Errors pass through as errors
- Limited practical use in modern Excel

**Excel Documentation:** [N function](https://support.microsoft.com/en-us/office/n-function-a624cad1-3635-4208-b54a-29733d1278c9)

**Implementation Status:** ✅ Full implementation

---

### NA

Returns the #N/A error value.

**Syntax:** `NA()`

**Parameters:** None

**Returns:** #N/A error value

**Examples:**
```swift
// Explicitly return #N/A
let notAvailable = try evaluator.evaluate("=NA()")

// Mark incomplete data
let status = try evaluator.evaluate("=IF(A1=\"\", NA(), A1*2)")

// Use in charts to hide points
let chartValue = try evaluator.evaluate("=IF(A1<0, NA(), A1)")

// Placeholder for missing data
let data = try evaluator.evaluate("=IF(ISBLANK(A1), NA(), VLOOKUP(A1, Table, 2))")
```

**Notes:**
- #N/A means "Not Available" or "Not Applicable"
- Charts skip #N/A values (doesn't plot points)
- Useful for hiding incomplete data in visualizations
- #N/A propagates through formulas
- Can be tested with ISNA function
- Better than empty cells for indicating missing data

**Excel Documentation:** [NA function](https://support.microsoft.com/en-us/office/na-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c)

**Implementation Status:** ✅ Full implementation

---

### ERROR.TYPE

Returns a number corresponding to an error type.

**Syntax:** `ERROR.TYPE(error_val)`

**Parameters:**
- `error_val`: The error value to identify

**Returns:** Number indicating error type:
- 1 = #NULL!
- 2 = #DIV/0!
- 3 = #VALUE!
- 4 = #REF!
- 5 = #NAME?
- 6 = #NUM!
- 7 = #N/A
- #N/A if value is not an error

**Examples:**
```swift
// Identify error type
let errorNum = try evaluator.evaluate("=ERROR.TYPE(A1)")

// Handle different errors differently
let message = try evaluator.evaluate("""
    =CHOOSE(ERROR.TYPE(A1), 
        "Null intersection", 
        "Division by zero", 
        "Wrong type", 
        "Invalid reference", 
        "Unknown name", 
        "Invalid number", 
        "Not available")
    """)

// Test for specific error
let isDivZero = try evaluator.evaluate("=ERROR.TYPE(A1/B1)=2")

// Custom error handling
let handle = try evaluator.evaluate("=IF(ISERROR(A1), \"Error \" & ERROR.TYPE(A1), A1)")
```

**Notes:**
- Returns #N/A if value is not an error
- Useful for creating custom error messages
- Can identify specific errors for targeted handling
- Combine with CHOOSE for error descriptions
- Returns numeric codes, not error text

**Excel Documentation:** [ERROR.TYPE function](https://support.microsoft.com/en-us/office/error-type-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa)

**Implementation Status:** ✅ Full implementation

---

## Cell Information Functions

### CELL

Returns information about the formatting, location, or contents of a cell.

**Syntax:** `CELL(info_type, [reference])`

**Parameters:**
- `info_type`: Text string specifying what information to return
- `reference` *(optional)*: Cell reference (defaults to last changed cell)

**Info Types:**
- "address" - Cell address as text
- "col" - Column number
- "row" - Row number
- "type" - Cell content type ("b"=blank, "l"=label/text, "v"=value)
- "width" - Column width
- "contents" - Cell value
- "format" - Number format code
- "color" - 1 if formatted for negative values, 0 otherwise
- "parentheses" - 1 if format uses parentheses for positive, 0 otherwise
- "prefix" - Text alignment prefix (' " ^ \)

**Returns:** Requested information about the cell

**Examples:**
```swift
// Get cell address
let address = try evaluator.evaluate("=CELL(\"address\", A1)")  // "$A$1"

// Get column number
let col = try evaluator.evaluate("=CELL(\"col\", E5)")  // 5

// Get row number
let row = try evaluator.evaluate("=CELL(\"row\", E5)")  // 5

// Get cell type
let type = try evaluator.evaluate("=CELL(\"type\", A1)")  // "v" for value

// Get column width
let width = try evaluator.evaluate("=CELL(\"width\", A1)")
```

**Notes:**
- Simplified implementation returns basic information
- Some info types may not be fully supported
- Format-related info types return default values
- Volatile function - recalculates with every change
- Case-insensitive info_type strings

**Excel Documentation:** [CELL function](https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf)

**Implementation Status:** ⚠️ Partial (basic info types supported)

---

### ROW

Returns the row number of a reference.

**Syntax:** `ROW([reference])`

**Parameters:**
- `reference` *(optional)*: Cell or range reference (defaults to current cell)

**Returns:** Row number as integer

**Examples:**
```swift
// Get row number of specific cell
let rowNum = try evaluator.evaluate("=ROW(C5)")  // 5

// Get current row (when used in cell)
let currentRow = try evaluator.evaluate("=ROW()")  // Returns row of formula cell

// Create sequence of row numbers
let sequence = try evaluator.evaluate("=ROW(A1:A10)")  // Array: 1,2,3...10

// Conditional formatting by row
let shade = try evaluator.evaluate("=MOD(ROW(), 2)=0")  // Even rows

// Reference offset rows
let offset = try evaluator.evaluate("=INDEX(A:A, ROW()+5)")
```

**Notes:**
- Without argument, returns row of cell containing formula
- With range, returns array of all row numbers
- Commonly used for creating sequences
- Useful for row-based conditional formatting
- Returns only the first row number if reference spans multiple rows

**Excel Documentation:** [ROW function](https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d)

**Implementation Status:** ✅ Full implementation

---

### COLUMN

Returns the column number of a reference.

**Syntax:** `COLUMN([reference])`

**Parameters:**
- `reference` *(optional)*: Cell or range reference (defaults to current cell)

**Returns:** Column number as integer (A=1, B=2, etc.)

**Examples:**
```swift
// Get column number of specific cell
let colNum = try evaluator.evaluate("=COLUMN(D1)")  // 4

// Get current column
let currentCol = try evaluator.evaluate("=COLUMN()")  // Returns column of formula cell

// Create sequence of column numbers
let sequence = try evaluator.evaluate("=COLUMN(A1:J1)")  // Array: 1,2,3...10

// Convert column to letter
let letter = try evaluator.evaluate("=ADDRESS(1, COLUMN(E1), 4)")

// Reference offset columns
let offset = try evaluator.evaluate("=INDEX(1:1, COLUMN()+3)")
```

**Notes:**
- Without argument, returns column of cell containing formula
- With range, returns array of all column numbers
- A=1, B=2, C=3, ..., Z=26, AA=27, etc.
- Useful for creating column sequences
- Returns only the first column number if reference spans multiple columns

**Excel Documentation:** [COLUMN function](https://support.microsoft.com/en-us/office/column-function-44e8c754-711c-4df3-9da4-47a55042554b)

**Implementation Status:** ✅ Full implementation

---

### ROWS

Returns the number of rows in a reference or array.

**Syntax:** `ROWS(array)`

**Parameters:**
- `array`: Range, array, or reference

**Returns:** Number of rows as integer

**Examples:**
```swift
// Count rows in range
let rowCount = try evaluator.evaluate("=ROWS(A1:A10)")  // 10

// Count rows in 2D range
let rows2D = try evaluator.evaluate("=ROWS(A1:C10)")  // 10

// Dynamic range size
let size = try evaluator.evaluate("=ROWS(A1:INDEX(A:A, COUNTA(A:A)))")

// Create dynamic formulas
let avg = try evaluator.evaluate("=SUMPRODUCT(A1:A10)/ROWS(A1:A10)")

// Validate range size
let isValid = try evaluator.evaluate("=ROWS(A1:A10)=10")
```

**Notes:**
- Returns number of rows in the range
- Works with single cells (returns 1)
- Works with multi-dimensional arrays
- Useful for dynamic formula calculations
- Combine with COLUMNS for total cell count

**Excel Documentation:** [ROWS function](https://support.microsoft.com/en-us/office/rows-function-b592593e-3fc2-47f2-bec1-bda493811d0c)

**Implementation Status:** ✅ Full implementation

---

### COLUMNS

Returns the number of columns in a reference or array.

**Syntax:** `COLUMNS(array)`

**Parameters:**
- `array`: Range, array, or reference

**Returns:** Number of columns as integer

**Examples:**
```swift
// Count columns in range
let colCount = try evaluator.evaluate("=COLUMNS(A1:J1)")  // 10

// Count columns in 2D range
let cols2D = try evaluator.evaluate("=COLUMNS(A1:C10)")  // 3

// Dynamic column count
let dynamicCols = try evaluator.evaluate("=COLUMNS(A1:INDEX(1:1, COUNTA(1:1)))")

// Calculate total cells
let totalCells = try evaluator.evaluate("=ROWS(A1:C10)*COLUMNS(A1:C10)")

// Validate table structure
let hasThreeCols = try evaluator.evaluate("=COLUMNS(A1:Z1)=3")
```

**Notes:**
- Returns number of columns in the range
- Works with single cells (returns 1)
- Works with multi-dimensional arrays
- A to Z = 26 columns, A to AA = 27 columns
- Combine with ROWS for range dimensions

**Excel Documentation:** [COLUMNS function](https://support.microsoft.com/en-us/office/columns-function-4e8e7b4e-e603-43e8-b177-956088fa48ca)

**Implementation Status:** ✅ Full implementation

---

### ADDRESS

Creates a cell address text string from row and column numbers.

**Syntax:** `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])`

**Parameters:**
- `row_num`: Row number
- `column_num`: Column number
- `abs_num` *(optional)*: Reference type (1=absolute, 2=abs row/rel col, 3=rel row/abs col, 4=relative)
- `a1` *(optional)*: TRUE for A1 style, FALSE for R1C1 style
- `sheet_text` *(optional)*: Sheet name to include

**Returns:** Cell address as text string

**Examples:**
```swift
// Create basic address
let addr = try evaluator.evaluate("=ADDRESS(5, 3)")  // "$C$5"

// Create relative address
let relAddr = try evaluator.evaluate("=ADDRESS(5, 3, 4)")  // "C5"

// Mixed reference - absolute row
let mixedAddr = try evaluator.evaluate("=ADDRESS(5, 3, 2)")  // "C$5"

// Mixed reference - absolute column
let mixedAddr2 = try evaluator.evaluate("=ADDRESS(5, 3, 3)")  // "$C5"

// Use with ROW and COLUMN for dynamic references
let dynamic = try evaluator.evaluate("=ADDRESS(ROW(), COLUMN()+1)")
```

**Notes:**
- Default is absolute reference ($A$1)
- Row and column numbers must be positive integers
- Can create addresses outside current worksheet range
- Combine with INDIRECT to reference the created address
- Sheet name parameter not fully implemented

**Excel Documentation:** [ADDRESS function](https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89)

**Implementation Status:** ✅ Full implementation (basic parameters)

---

### AREAS

Returns the number of areas in a reference.

**Syntax:** `AREAS(reference)`

**Parameters:**
- `reference`: Reference to a cell or range

**Returns:** Number of areas in the reference

**Examples:**
```swift
// Single area
let singleArea = try evaluator.evaluate("=AREAS(A1:B10)")  // 1

// Multiple areas (using comma)
let multiArea = try evaluator.evaluate("=AREAS((A1:B10,D1:E10))")  // 2

// Named range areas
let namedAreas = try evaluator.evaluate("=AREAS(MyRange)")

// Validate reference structure
let isSingleArea = try evaluator.evaluate("=AREAS(A1:B10)=1")
```

**Notes:**
- An "area" is a contiguous range of cells
- Multiple areas are separated by commas in Excel
- Simplified implementation assumes single area
- Useful for validating reference structure
- Works with named ranges

**Excel Documentation:** [AREAS function](https://support.microsoft.com/en-us/office/areas-function-8392ba32-7a41-43b3-96b0-3695d2ec6152)

**Implementation Status:** ⚠️ Partial (always returns 1)

---

### FORMULATEXT

Returns the formula in a cell as text.

**Syntax:** `FORMULATEXT(reference)`

**Parameters:**
- `reference`: Cell reference

**Returns:** Formula as text string, or #N/A if cell doesn't contain formula

**Examples:**
```swift
// Get formula from cell
let formula = try evaluator.evaluate("=FORMULATEXT(A1)")

// Check if cells have same formula
let sameFormula = try evaluator.evaluate("=FORMULATEXT(A1)=FORMULATEXT(A2)")

// Document formulas
let doc = try evaluator.evaluate("=IF(ISFORMULA(A1), FORMULATEXT(A1), \"No formula\")")

// Extract formula for analysis
let analysis = try evaluator.evaluate("=SEARCH(\"SUM\", FORMULATEXT(A1))")
```

**Notes:**
- Returns #N/A if cell doesn't contain a formula
- Returns formula exactly as entered
- Includes the leading equals sign
- Useful for formula auditing and documentation
- Simplified implementation

**Excel Documentation:** [FORMULATEXT function](https://support.microsoft.com/en-us/office/formulatext-function-0a786771-54fd-4ae2-96ee-09cda35439c8)

**Implementation Status:** 🔄 Stub (returns #N/A - requires formula tracking)

---

## Workbook Information Functions

### INFO

Returns information about the current operating environment.

**Syntax:** `INFO(type_text)`

**Parameters:**
- `type_text`: Text specifying what information to return

**Info Types:**
- "directory" - Path of current directory
- "numfile" - Number of worksheets in open workbooks
- "origin" - Reference to first visible cell
- "osversion" - Operating system version
- "recalc" - Recalculation mode
- "release" - Excel version
- "system" - Operating system ("mac" or "pcdos")

**Returns:** Requested environment information

**Examples:**
```swift
// Get operating system
let os = try evaluator.evaluate("=INFO(\"system\")")  // "mac"

// Get OS version
let version = try evaluator.evaluate("=INFO(\"osversion\")")

// Get Excel version
let excelVer = try evaluator.evaluate("=INFO(\"release\")")

// Get recalculation mode
let recalc = try evaluator.evaluate("=INFO(\"recalc\")")  // "Automatic"

// Get number of files
let fileCount = try evaluator.evaluate("=INFO(\"numfile\")")
```

**Notes:**
- Returns system and environment information
- Simplified implementation with default values
- Some info types may return placeholder values
- Rarely used in modern formulas
- Useful for system-specific formula behavior

**Excel Documentation:** [INFO function](https://support.microsoft.com/en-us/office/info-function-725f259a-0e4b-49b3-8b52-58815c69acae)

**Implementation Status:** ⚠️ Partial (returns default values)

---

### SHEET

Returns the sheet number of a referenced sheet.

**Syntax:** `SHEET([value])`

**Parameters:**
- `value` *(optional)*: Sheet name or reference (defaults to current sheet)

**Returns:** Sheet number (position in workbook)

**Examples:**
```swift
// Get current sheet number
let sheetNum = try evaluator.evaluate("=SHEET()")  // 1

// Get specific sheet number
let sheet2 = try evaluator.evaluate("=SHEET(Sheet2!A1)")

// Get sheet by name
let namedSheet = try evaluator.evaluate("=SHEET(\"Data\")")

// Conditional by sheet
let conditional = try evaluator.evaluate("=IF(SHEET()=1, \"First\", \"Other\")")
```

**Notes:**
- Without argument, returns current sheet number
- First sheet = 1, second = 2, etc.
- Simplified implementation assumes single sheet
- Returns #REF! if sheet doesn't exist
- Useful for sheet-specific formulas

**Excel Documentation:** [SHEET function](https://support.microsoft.com/en-us/office/sheet-function-44718b6f-8b87-47a1-a9d6-b701c06cff24)

**Implementation Status:** ⚠️ Partial (always returns 1)

---

### SHEETS

Returns the number of sheets in a reference.

**Syntax:** `SHEETS([reference])`

**Parameters:**
- `reference` *(optional)*: Reference to a sheet or range (defaults to all sheets)

**Returns:** Number of sheets

**Examples:**
```swift
// Get total sheets in workbook
let totalSheets = try evaluator.evaluate("=SHEETS()")  // 1

// Get sheets referenced
let sheetsInRef = try evaluator.evaluate("=SHEETS(Sheet1:Sheet3)")

// Validate workbook structure
let hasMultipleSheets = try evaluator.evaluate("=SHEETS()>1")

// Calculate across sheets
let allSheets = try evaluator.evaluate("=SHEETS()*100")
```

**Notes:**
- Without argument, returns total sheets in workbook
- With reference, counts sheets in that range
- Simplified implementation assumes single sheet
- Useful for workbook structure validation
- Hidden sheets are included in count

**Excel Documentation:** [SHEETS function](https://support.microsoft.com/en-us/office/sheets-function-770515eb-e1e8-45ce-8066-b557e5e4b80b)

**Implementation Status:** ⚠️ Partial (always returns 1)

---

## Working with Information Functions

### Type Validation

Use IS* functions to validate data types before processing:

```swift
// Comprehensive validation
let validate = try evaluator.evaluate("""
    =IF(ISBLANK(A1), "Empty",
        IF(ISERROR(A1), "Error",
            IF(ISTEXT(A1), "Text: " & A1,
                IF(ISNUMBER(A1), "Number: " & A1,
                    "Other"))))
    """)

// Type-specific processing
let process = try evaluator.evaluate("""
    =IF(ISNUMBER(A1), A1*2,
        IF(ISTEXT(A1), UPPER(A1),
            A1))
    """)

// Safe calculation
let safeDivide = try evaluator.evaluate("""
    =IF(OR(NOT(ISNUMBER(A1)), NOT(ISNUMBER(B1)), B1=0),
        "Invalid",
        A1/B1)
    """)
```

### Error Detection and Handling

Combine error checking functions for robust formulas:

```swift
// Specific error handling
let handleError = try evaluator.evaluate("""
    =IF(ISNA(A1), "Not found",
        IF(ISERR(A1), "Calculation error",
            A1))
    """)

// Error type identification
let errorMsg = try evaluator.evaluate("""
    =IF(ISERROR(A1),
        "Error type: " & ERROR.TYPE(A1),
        A1)
    """)

// Conditional on error presence
let conditional = try evaluator.evaluate("""
    =IF(ISERROR(VLOOKUP(A1, Table, 2, 0)),
        VLOOKUP(A1, AlternateTable, 2, 0),
        VLOOKUP(A1, Table, 2, 0))
    """)
```

### Dynamic References

Build dynamic cell references using cell information:

```swift
// Create dynamic address
let dynamicRef = try evaluator.evaluate("""
    =INDIRECT(ADDRESS(ROW(), COLUMN()+1))
    """)

// Offset reference
let offset = try evaluator.evaluate("""
    =INDEX(A:A, ROW()+5)
    """)

// Dynamic range sizing
let dynamicRange = try evaluator.evaluate("""
    =AVERAGE(INDIRECT("A1:A" & ROWS(A:A)))
    """)

// Column-based calculations
let sumColumn = try evaluator.evaluate("""
    =SUM(INDIRECT(ADDRESS(2, COLUMN()) & ":" & ADDRESS(100, COLUMN())))
    """)
```

### Conditional Formatting Formulas

Information functions are perfect for conditional formatting:

```swift
// Highlight even rows
let evenRows = try evaluator.evaluate("=ISEVEN(ROW())")

// Highlight errors
let highlightErrors = try evaluator.evaluate("=ISERROR(A1)")

// Highlight specific error types
let highlightNA = try evaluator.evaluate("=ISNA(A1)")

// Highlight text vs numbers
let highlightType = try evaluator.evaluate("=ISTEXT(A1)")

// Alternate column shading
let alternateCols = try evaluator.evaluate("=MOD(COLUMN(), 2)=0")
```

### Data Quality Checks

Validate data quality and completeness:

```swift
// Check for incomplete data
let isComplete = try evaluator.evaluate("""
    =AND(
        NOT(ISBLANK(A1)),
        ISNUMBER(B1),
        ISTEXT(C1),
        NOT(ISERROR(D1))
    )
    """)

// Count blank cells
let blankCount = try evaluator.evaluate("=SUMPRODUCT(--ISBLANK(A1:A100))")

// Count error cells
let errorCount = try evaluator.evaluate("=SUMPRODUCT(--ISERROR(A1:A100))")

// Identify data type distribution
let typeDistribution = try evaluator.evaluate("""
    =SUMPRODUCT(--ISNUMBER(A1:A100)) & " numbers, " &
     SUMPRODUCT(--ISTEXT(A1:A100)) & " text"
    """)
```

### Range Analysis

Analyze ranges and references:

```swift
// Calculate total cells
let totalCells = try evaluator.evaluate("=ROWS(A1:E10)*COLUMNS(A1:E10)")

// Verify square range
let isSquare = try evaluator.evaluate("=ROWS(A1:Z10)=COLUMNS(A1:Z10)")

// Get range dimensions
let dimensions = try evaluator.evaluate("""
    =ROWS(A1:E10) & " x " & COLUMNS(A1:E10)
    """)

// Dynamic average per column
let avgPerCol = try evaluator.evaluate("=SUM(A1:Z100)/COLUMNS(A1:Z100)")
```

## Common Patterns

### Type-Safe Calculations

Always validate types before calculations:

```swift
// Safe arithmetic
=IF(AND(ISNUMBER(A1), ISNUMBER(B1)), A1+B1, "Invalid input")

// Safe text operations
=IF(ISTEXT(A1), LEN(A1), "Not text")

// Safe logical operations
=IF(ISLOGICAL(A1), IF(A1, "Yes", "No"), "Not boolean")
```

### Multi-Level Error Handling

Handle different error types appropriately:

```swift
// Three-tier error handling
=IF(ISNA(A1), "Not found",
    IF(ISERR(A1), "Error: " & ERROR.TYPE(A1),
        A1))

// Suppress #N/A, show other errors
=IF(ISNA(A1), "", IF(ISERROR(A1), "ERROR", A1))

// Custom messages by error type
=IF(ERROR.TYPE(A1)=2, "Check denominator",
    IF(ERROR.TYPE(A1)=7, "Value not found",
        A1))
```

### Dynamic Cell Selection

Use ROW, COLUMN, and ADDRESS for dynamic references:

```swift
// Reference cell to the right
=INDIRECT(ADDRESS(ROW(), COLUMN()+1))

// Reference same column, different row
=INDIRECT(ADDRESS(5, COLUMN()))

// Create diagonal reference
=INDIRECT(ADDRESS(ROW(), ROW()))

// Reference by offset
=INDEX(A:Z, ROW(), COLUMN())
```

### Data Validation

Build comprehensive validation rules:

```swift
// Must be number in range
=AND(ISNUMBER(A1), A1>=0, A1<=100)

// Must be non-blank text
=AND(ISTEXT(A1), NOT(ISBLANK(A1)))

// Must be boolean
=ISLOGICAL(A1)

// Must not be error
=NOT(ISERROR(A1))
```

## Performance Notes

- **IS* functions**: Very fast, minimal overhead
- **TYPE function**: Fast type checking, useful for branching logic
- **CELL/INFO**: May be slower due to environment queries
- **ROW/COLUMN**: Fast, frequently used in array formulas
- **FORMULATEXT**: Requires formula tracking, may have overhead
- **Volatile functions**: CELL, INFO recalculate frequently
- **Array usage**: IS* functions work efficiently on ranges with SUMPRODUCT

## See Also

- <doc:Logical> - Logical functions for conditional operations
- <doc:Lookup> - Lookup and reference functions
- <doc:Text> - Text manipulation functions
- <doc:Mathematical> - Mathematical operations
- <doc:FormulaReference> - Complete formula reference
- ``FormulaEvaluator`` - Evaluate information formulas programmatically
