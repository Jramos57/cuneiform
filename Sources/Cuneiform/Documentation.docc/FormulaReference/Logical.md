# Logical Functions

Boolean logic operations and conditional evaluation functions.

## Overview

Logical functions perform boolean operations, conditional evaluation, and error handling. These functions are essential for building decision trees, validating data, and handling errors gracefully in your formulas.

Cuneiform implements 11 logical functions compatible with Excel, enabling sophisticated conditional logic and error handling in spreadsheet calculations.

### Quick Reference

| Function | Description | Status |
|----------|-------------|--------|
| **AND** | Returns TRUE if all arguments are TRUE | ✅ Full |
| **FALSE** | Returns the boolean value FALSE | 🔄 Stub |
| **IF** | Tests a condition and returns different values | ✅ Full |
| **IFERROR** | Returns a value if expression is an error | ✅ Full |
| **IFNA** | Returns a value if expression is #N/A | ✅ Full |
| **IFS** | Tests multiple conditions | ✅ Full |
| **NOT** | Reverses the logic of its argument | ✅ Full |
| **OR** | Returns TRUE if any argument is TRUE | ✅ Full |
| **SWITCH** | Evaluates an expression against values | ✅ Full |
| **TRUE** | Returns the boolean value TRUE | 🔄 Stub |
| **XOR** | Returns TRUE for odd number of TRUE arguments | ✅ Full |

## Conditional Functions

### IF

Tests a logical condition and returns one value for TRUE and another for FALSE.

**Syntax:** `IF(logical_test, value_if_true, [value_if_false])`

**Parameters:**
- `logical_test`: The condition to test
- `value_if_true`: Value returned if condition is TRUE
- `value_if_false` *(optional)*: Value returned if condition is FALSE (defaults to FALSE)

**Returns:** The value corresponding to the result of the logical test

**Examples:**
```swift
// Basic IF statement
let result = try evaluator.evaluate("=IF(A1>100, \"High\", \"Low\")")

// Nested IF for multiple conditions
let grade = try evaluator.evaluate("=IF(A1>=90, \"A\", IF(A1>=80, \"B\", IF(A1>=70, \"C\", \"F\")))")

// IF with calculations
let bonus = try evaluator.evaluate("=IF(B1>10000, B1*0.1, B1*0.05)")

// IF without else clause (returns FALSE)
let flag = try evaluator.evaluate("=IF(A1>0, \"Positive\")")
```

**Excel Documentation:** [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)

**Implementation Status:** ✅ Full implementation

---

### IFS

Evaluates multiple conditions and returns a value corresponding to the first TRUE condition.

**Syntax:** `IFS(condition1, value1, [condition2, value2], ...)`

**Parameters:**
- `condition1`: First logical test
- `value1`: Value returned if condition1 is TRUE
- `condition2...conditionN` *(optional)*: Additional logical tests
- `value2...valueN` *(optional)*: Corresponding values for additional conditions

**Returns:** The value corresponding to the first TRUE condition, or #N/A if no conditions are TRUE

**Examples:**
```swift
// Multiple conditions without nesting
let result = try evaluator.evaluate("=IFS(A1>=90, \"A\", A1>=80, \"B\", A1>=70, \"C\", A1>=60, \"D\", TRUE, \"F\")")

// Pricing tiers
let price = try evaluator.evaluate("=IFS(B1>=100, B1*0.8, B1>=50, B1*0.9, B1>=10, B1*0.95, TRUE, B1)")

// Status based on multiple criteria
let status = try evaluator.evaluate("=IFS(C1=\"Complete\", \"Done\", C1=\"In Progress\", \"Working\", C1=\"\", \"Not Started\")")
```

**Notes:**
- Conditions are evaluated in order from left to right
- Returns #N/A error if no condition evaluates to TRUE
- Use `TRUE` as the final condition for a default value
- More efficient than nested IF statements

**Excel Documentation:** [IFS function](https://support.microsoft.com/en-us/office/ifs-function-36329a26-37b2-467c-972b-4a39bd951d45)

**Implementation Status:** ✅ Full implementation

---

### SWITCH

Evaluates an expression against a list of values and returns the result corresponding to the first matching value.

**Syntax:** `SWITCH(expression, value1, result1, [value2, result2], ..., [default])`

**Parameters:**
- `expression`: Expression to evaluate and compare
- `value1`: First value to match against expression
- `result1`: Result returned if value1 matches expression
- `value2...valueN` *(optional)*: Additional values to match
- `result2...resultN` *(optional)*: Corresponding results
- `default` *(optional)*: Value returned if no match found (must be last argument)

**Returns:** The result corresponding to the first matching value, default value, or #N/A if no match

**Examples:**
```swift
// Department name lookup
let dept = try evaluator.evaluate("=SWITCH(A1, 1, \"Sales\", 2, \"Marketing\", 3, \"Engineering\", \"Unknown\")")

// Convert codes to descriptions
let status = try evaluator.evaluate("=SWITCH(B1, \"A\", \"Active\", \"I\", \"Inactive\", \"P\", \"Pending\", \"Invalid\")")

// Numeric range mapping
let tier = try evaluator.evaluate("=SWITCH(C1, 1, \"Bronze\", 2, \"Silver\", 3, \"Gold\", 4, \"Platinum\")")

// Without default (returns #N/A if no match)
let color = try evaluator.evaluate("=SWITCH(D1, \"R\", \"Red\", \"G\", \"Green\", \"B\", \"Blue\")")
```

**Notes:**
- More readable than nested IF statements for matching discrete values
- Expression is evaluated once, not for each comparison
- If the final argument count is even, it's treated as a default value
- Returns #N/A if no match and no default provided

**Excel Documentation:** [SWITCH function](https://support.microsoft.com/en-us/office/switch-function-47ab33c0-28ce-4530-8a45-d532ec4aa25e)

**Implementation Status:** ✅ Full implementation

---

## Error Handling Functions

### IFERROR

Returns a value you specify if a formula evaluates to an error; otherwise returns the result of the formula.

**Syntax:** `IFERROR(value, value_if_error)`

**Parameters:**
- `value`: The expression to check for an error
- `value_if_error`: Value returned if the expression evaluates to any error

**Returns:** The original value if no error, or the alternate value if an error occurs

**Examples:**
```swift
// Handle division by zero
let result = try evaluator.evaluate("=IFERROR(A1/B1, 0)")

// VLOOKUP with error handling
let lookup = try evaluator.evaluate("=IFERROR(VLOOKUP(A1, B1:C10, 2, FALSE), \"Not Found\")")

// Nested calculations with error protection
let calc = try evaluator.evaluate("=IFERROR(SQRT(A1) + SQRT(B1), \"Invalid\")")

// Chain multiple operations
let complex = try evaluator.evaluate("=IFERROR(INDEX(A1:A10, MATCH(B1, C1:C10, 0)), \"No match\")")
```

**Notes:**
- Catches all error types: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
- Useful for preventing error propagation in complex formulas
- More efficient than testing for errors with multiple IF statements
- If you only need to catch #N/A errors, use IFNA instead

**Excel Documentation:** [IFERROR function](https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611)

**Implementation Status:** ✅ Full implementation

---

### IFNA

Returns a value you specify if the expression evaluates to #N/A; otherwise returns the result of the expression.

**Syntax:** `IFNA(value, value_if_na)`

**Parameters:**
- `value`: The expression to check for #N/A error
- `value_if_na`: Value returned if the expression evaluates to #N/A

**Returns:** The original value if not #N/A, or the alternate value if #N/A

**Examples:**
```swift
// Handle VLOOKUP not found
let result = try evaluator.evaluate("=IFNA(VLOOKUP(A1, B1:C10, 2, FALSE), \"Not in list\")")

// MATCH with default value
let position = try evaluator.evaluate("=IFNA(MATCH(A1, B1:B100, 0), 0)")

// XLOOKUP with custom message
let lookup = try evaluator.evaluate("=IFNA(XLOOKUP(A1, B1:B10, C1:C10), \"No match found\")")

// Keep other errors visible
let check = try evaluator.evaluate("=IFNA(A1/B1, \"N/A\")")  // Shows #DIV/0! if B1 is zero
```

**Notes:**
- Only catches #N/A errors, unlike IFERROR which catches all errors
- More specific than IFERROR - allows other error types to propagate
- Commonly used with lookup functions (VLOOKUP, XLOOKUP, MATCH)
- Introduced in Excel 2013

**Excel Documentation:** [IFNA function](https://support.microsoft.com/en-us/office/ifna-function-6626c961-a569-42fc-a49d-79b4951fd461)

**Implementation Status:** ✅ Full implementation

---

## Boolean Logic Functions

### AND

Returns TRUE if all arguments evaluate to TRUE.

**Syntax:** `AND(logical1, [logical2], ...)`

**Parameters:**
- `logical1`: First condition to test
- `logical2...logicalN` *(optional)*: Additional conditions to test

**Returns:** Boolean TRUE if all arguments are TRUE, FALSE otherwise

**Examples:**
```swift
// Simple AND
let result = try evaluator.evaluate("=AND(A1>0, B1>0)")  // TRUE if both positive

// Multiple conditions
let valid = try evaluator.evaluate("=AND(A1>=0, A1<=100, B1<>\"\")")

// With IF for complex logic
let approved = try evaluator.evaluate("=IF(AND(A1>1000, B1=\"Approved\", C1<30), \"Yes\", \"No\")")

// Array arguments
let allPositive = try evaluator.evaluate("=AND(A1:A10>0)")  // TRUE if all cells are positive
```

**Notes:**
- Returns FALSE if any argument is FALSE
- Empty cells and text are ignored when in arrays
- Numbers: 0 is FALSE, non-zero is TRUE
- Can evaluate arrays and ranges

**Excel Documentation:** [AND function](https://support.microsoft.com/en-us/office/and-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9)

**Implementation Status:** ✅ Full implementation

---

### OR

Returns TRUE if any argument evaluates to TRUE.

**Syntax:** `OR(logical1, [logical2], ...)`

**Parameters:**
- `logical1`: First condition to test
- `logical2...logicalN` *(optional)*: Additional conditions to test

**Returns:** Boolean TRUE if any argument is TRUE, FALSE if all are FALSE

**Examples:**
```swift
// Simple OR
let result = try evaluator.evaluate("=OR(A1>100, B1>100)")  // TRUE if either exceeds 100

// Multiple alternatives
let flagged = try evaluator.evaluate("=OR(A1=\"Error\", A1=\"Warning\", A1=\"Critical\")")

// With IF for validation
let valid = try evaluator.evaluate("=IF(OR(A1=\"\", A1=0), \"Empty or Zero\", \"Has Value\")")

// Array arguments
let anyNegative = try evaluator.evaluate("=OR(A1:A10<0)")  // TRUE if any cell is negative
```

**Notes:**
- Returns TRUE if at least one argument is TRUE
- Returns FALSE only if all arguments are FALSE
- Empty cells and text are ignored when in arrays
- Often combined with IF for decision logic

**Excel Documentation:** [OR function](https://support.microsoft.com/en-us/office/or-function-7d17ad14-8700-4281-b308-00b131e22af0)

**Implementation Status:** ✅ Full implementation

---

### NOT

Reverses the logic of its argument.

**Syntax:** `NOT(logical)`

**Parameters:**
- `logical`: A value or expression that can be evaluated to TRUE or FALSE

**Returns:** Boolean TRUE if argument is FALSE, FALSE if argument is TRUE

**Examples:**
```swift
// Reverse a boolean
let result = try evaluator.evaluate("=NOT(A1)")  // FALSE if A1 is TRUE

// Negate a comparison
let notEqual = try evaluator.evaluate("=NOT(A1=B1)")  // Same as A1<>B1

// Complex logic inversion
let excluded = try evaluator.evaluate("=NOT(AND(A1>0, A1<100))")  // TRUE if outside range

// With IF for readable logic
let check = try evaluator.evaluate("=IF(NOT(A1=\"\"), \"Has value\", \"Empty\")")
```

**Notes:**
- Converts TRUE to FALSE and FALSE to TRUE
- Numbers: 0 becomes TRUE, non-zero becomes FALSE
- Useful for inverting complex logical expressions
- Can make formulas more readable than using inequality operators

**Excel Documentation:** [NOT function](https://support.microsoft.com/en-us/office/not-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77)

**Implementation Status:** ✅ Full implementation

---

### XOR

Returns a logical exclusive OR of all arguments.

**Syntax:** `XOR(logical1, [logical2], ...)`

**Parameters:**
- `logical1`: First logical value to test
- `logical2...logicalN` *(optional)*: Additional logical values (1 to 254)

**Returns:** Boolean TRUE if an odd number of arguments are TRUE, FALSE otherwise

**Examples:**
```swift
// Simple XOR - TRUE only if exactly one is TRUE
let result = try evaluator.evaluate("=XOR(A1>100, B1>100)")

// Multiple values - TRUE if odd count are TRUE
let toggle = try evaluator.evaluate("=XOR(A1, B1, C1)")  // TRUE if 1 or 3 are TRUE

// Toggle logic with three conditions
let state = try evaluator.evaluate("=XOR(A1=\"On\", B1=\"On\", C1=\"On\")")

// Detect single outlier in set
let outlier = try evaluator.evaluate("=XOR(A1>0, B1>0, C1>0, D1>0)")
```

**Notes:**
- Returns TRUE when an odd number of arguments are TRUE
- Returns FALSE when an even number (including zero) are TRUE
- Different from OR (which returns TRUE if any are TRUE)
- Useful for toggle logic and detecting single exceptions
- Numbers: 0 is FALSE, non-zero is TRUE

**Excel Documentation:** [XOR function](https://support.microsoft.com/en-us/office/xor-function-1548d4c2-5e47-4f77-9a92-0533bba14f37)

**Implementation Status:** ✅ Full implementation

---

## Boolean Constants

### TRUE

Returns the logical value TRUE.

**Syntax:** `TRUE()`

**Parameters:** None

**Returns:** Boolean value TRUE

**Examples:**
```swift
// Explicit TRUE value
let value = try evaluator.evaluate("=TRUE()")  // Returns TRUE

// Use in comparisons
let result = try evaluator.evaluate("=IF(A1=TRUE(), \"Yes\", \"No\")")

// Default condition in IFS
let grade = try evaluator.evaluate("=IFS(A1>=90, \"A\", A1>=80, \"B\", TRUE(), \"F\")")
```

**Notes:**
- Usually unnecessary as you can just type TRUE without parentheses
- Primarily for compatibility and explicit boolean values
- More commonly used without the function: `=IF(TRUE, "yes", "no")`

**Excel Documentation:** [TRUE function](https://support.microsoft.com/en-us/office/true-function-7652c6e3-8987-48d0-97cd-ef223246b3fb)

**Implementation Status:** 🔄 Stub (not yet implemented as function, but TRUE constant is supported)

---

### FALSE

Returns the logical value FALSE.

**Syntax:** `FALSE()`

**Parameters:** None

**Returns:** Boolean value FALSE

**Examples:**
```swift
// Explicit FALSE value
let value = try evaluator.evaluate("=FALSE()")  // Returns FALSE

// Use in logic
let result = try evaluator.evaluate("=IF(A1=FALSE(), \"No\", \"Yes\")")

// Initialize boolean cell
let init = try evaluator.evaluate("=FALSE()")
```

**Notes:**
- Usually unnecessary as you can just type FALSE without parentheses
- Primarily for compatibility and explicit boolean values
- More commonly used without the function: `=IF(FALSE, "yes", "no")`

**Excel Documentation:** [FALSE function](https://support.microsoft.com/en-us/office/false-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904)

**Implementation Status:** 🔄 Stub (not yet implemented as function, but FALSE constant is supported)

---

## Working with Logical Functions

### Combining Logical Functions

Build complex decision trees by nesting logical functions:

```swift
// Nested IF with AND
let approved = try evaluator.evaluate("""
    =IF(AND(A1>1000, B1="Approved", C1<TODAY()), "Processed", "Pending")
    """)

// Multiple conditions with OR and AND
let eligible = try evaluator.evaluate("""
    =IF(OR(AND(Age>=18, Citizen=TRUE), Status="Exempt"), "Eligible", "Not Eligible")
    """)

// IFS with complex conditions
let tier = try evaluator.evaluate("""
    =IFS(
        AND(Sales>10000, Rating>=4.5), "Platinum",
        AND(Sales>5000, Rating>=4.0), "Gold",
        Sales>1000, "Silver",
        TRUE, "Bronze"
    )
    """)
```

### Error Handling Patterns

Use error handling functions to create robust formulas:

```swift
// Nested error handling
let safe = try evaluator.evaluate("""
    =IFERROR(
        IFNA(VLOOKUP(A1, Table, 2, FALSE), "Not Found"),
        "Error in lookup"
    )
    """)

// Chain of lookups with fallback
let result = try evaluator.evaluate("""
    =IFNA(XLOOKUP(A1, Primary), 
        IFNA(XLOOKUP(A1, Secondary), 
            "Not in either table"))
    """)

// Preserve specific errors
let calc = try evaluator.evaluate("""
    =IFNA(A1/B1, "N/A")
    """)  // Shows #DIV/0! but replaces #N/A
```

### Validation Logic

Combine logical functions for data validation:

```swift
// Multi-criteria validation
let isValid = try evaluator.evaluate("""
    =AND(
        NOT(A1=""),
        A1>=0,
        A1<=100,
        ISNUMBER(A1)
    )
    """)

// Exclusive conditions check
let category = try evaluator.evaluate("""
    =IF(XOR(IsRetail, IsWholesale), 
        "Valid", 
        "Error: Must be exactly one category")
    """)

// Complex business rules
let status = try evaluator.evaluate("""
    =IFS(
        AND(Amount>10000, Priority="High"), "Escalate",
        OR(Days>30, Status="Overdue"), "Follow Up",
        Status="Complete", "Archive",
        TRUE, "Active"
    )
    """)
```

## Common Patterns

### Replacing Nested IFs

**Before (nested IF):**
```swift
=IF(A1>=90, "A", IF(A1>=80, "B", IF(A1>=70, "C", IF(A1>=60, "D", "F"))))
```

**After (using IFS):**
```swift
=IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", A1>=60, "D", TRUE, "F")
```

### Safe Lookups

Always wrap lookups in error handlers:

```swift
// Basic lookup with error handling
=IFERROR(VLOOKUP(A1, Table, 2, FALSE), "Not Found")

// Only handle #N/A, show other errors
=IFNA(XLOOKUP(A1, Range1, Range2), "No match")

// Multiple fallback options
=IFNA(XLOOKUP(A1, Primary), IFNA(XLOOKUP(A1, Secondary), "Not Found"))
```

### Boolean Flags

Use logical functions to create boolean flags:

```swift
// Eligibility flag
=AND(Age>=18, Income>30000, CreditScore>=650)

// Any warning condition
=OR(Status="Error", Days>30, Amount<0)

// Exactly one condition true
=XOR(IsNew, IsModified, IsDeleted)
```

## Performance Notes

- **Short-circuit evaluation**: AND returns FALSE on first FALSE argument, OR returns TRUE on first TRUE argument
- **IFS vs nested IF**: IFS is more readable and performs similarly
- **SWITCH vs IF chains**: SWITCH is more efficient for comparing a single value against multiple options
- **IFERROR cost**: Adds minimal overhead; use liberally for robustness
- **Array arguments**: AND/OR can evaluate entire ranges efficiently

## See Also

- <doc:Information> - Type checking functions (ISERROR, ISNA, ISLOGICAL)
- <doc:Lookup> - Lookup functions commonly used with logical operations
- <doc:Mathematical> - Comparison operators used in logical tests
- <doc:FormulaReference> - Complete formula reference
- ``FormulaEvaluator`` - Evaluate logical formulas programmatically
