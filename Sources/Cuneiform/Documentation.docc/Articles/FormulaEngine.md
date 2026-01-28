# Formula Engine

Deep dive into Cuneiform's comprehensive formula parsing and evaluation system.

## Overview

Cuneiform includes a powerful formula engine with **467 Excel-compatible functions** across 12 categories. The engine parses formula strings into an Abstract Syntax Tree (AST), evaluates them recursively, and provides Excel-compatible results including error propagation.

### Key Features

- **467 Functions**: 97% full implementations (453 functions), 3% stubs (14 functions)
- **Full Excel Compatibility**: Matches Excel's function behavior and error handling
- **Type System**: Supports numbers, strings, booleans, errors, and arrays
- **Error Propagation**: Complete #REF!, #VALUE!, #DIV/0!, #N/A, #NAME?, #NUM!, #NULL! support
- **Range Operations**: Handles cell ranges like `A1:C10` in function arguments
- **Array Formulas**: Supports array constants and array results
- **Cross-Sheet References**: Parse and evaluate references like `Sheet2!A1`

## Architecture

The formula engine consists of two main components:

### 1. FormulaParser

Converts formula strings into Abstract Syntax Trees:

```swift
let parser = FormulaParser()
let ast = try parser.parse("=SUM(A1:A10) / COUNT(A1:A10)")

// Resulting AST:
// BinaryOp(.divide)
//   ├─ FunctionCall("SUM", args: [Range(A1:A10)])
//   └─ FunctionCall("COUNT", args: [Range(A1:A10)])
```

Located at: `Sources/Cuneiform/SpreadsheetML/FormulaParser.swift:1`

### 2. FormulaEvaluator

Evaluates AST nodes and executes functions:

```swift
var evaluator = FormulaEvaluator()
evaluator.setCell("A1", value: .number(10))
evaluator.setCell("A2", value: .number(20))
evaluator.setCell("A3", value: .number(30))

let result = evaluator.evaluate("=AVERAGE(A1:A3)")  // 20.0
```

Located at: `Sources/Cuneiform/SpreadsheetML/FormulaEvaluator.swift:1` (12,357 lines)

## Formula Expression Types

The ``FormulaExpression`` enum represents all possible AST nodes:

```swift
public enum FormulaExpression: Sendable {
    case number(Double)
    case string(String)
    case cellRef(CellReference)
    case range(CellReference, CellReference)
    case binaryOp(BinaryOperator, FormulaExpression, FormulaExpression)
    case functionCall(String, [FormulaExpression])
    case error(String)
}
```

### Examples

```swift
// Number literal
.number(42.0)

// String literal
.string("Hello")

// Cell reference
.cellRef(CellReference(column: "A", row: 1))

// Range
.range(
    CellReference(column: "A", row: 1),
    CellReference(column: "C", row: 10)
)

// Binary operation
.binaryOp(.add, .number(10), .number(5))

// Function call
.functionCall("SUM", [.range(ref1, ref2)])
```

## Formula Values

The ``FormulaValue`` enum represents evaluation results:

```swift
public enum FormulaValue: Sendable, Equatable {
    case number(Double)
    case string(String)
    case boolean(Bool)
    case error(String)
    case array([[FormulaValue]])
}
```

### Type Coercion

FormulaValue provides automatic type coercion:

```swift
let value: FormulaValue = .number(42)

// Convert to different types
value.asDouble   // Optional(42.0)
value.asString   // "42.0"
value.asBoolean  // Optional(true) - non-zero is true

let boolValue: FormulaValue = .boolean(true)
boolValue.asDouble  // Optional(1.0)
```

## Binary Operators

Supported operators with Excel-compatible behavior:

### Arithmetic Operators

```swift
// Addition
evaluate("=10 + 5")  // 15.0

// Subtraction
evaluate("=10 - 5")  // 5.0

// Multiplication
evaluate("=10 * 5")  // 50.0

// Division
evaluate("=10 / 5")  // 2.0
evaluate("=10 / 0")  // #DIV/0!

// Power
evaluate("=2 ^ 3")   // 8.0
```

### Comparison Operators

```swift
evaluate("=10 > 5")   // TRUE
evaluate("=10 < 5")   // FALSE
evaluate("=10 >= 10") // TRUE
evaluate("=10 <= 5")  // FALSE
evaluate("=10 = 10")  // TRUE (equals)
evaluate("=10 <> 5")  // TRUE (not equals)
```

### String Concatenation

```swift
evaluate("=\"Hello\" & \" World\"")  // "Hello World"
evaluate("=\"Value: \" & 42")        // "Value: 42"
```

## Function Categories

### Mathematical Functions (67 functions)

Basic arithmetic, rounding, trigonometry, matrix operations.

**Examples:**
```swift
evaluate("=SUM(1, 2, 3)")           // 6.0
evaluate("=AVERAGE(10, 20, 30)")    // 20.0
evaluate("=ROUND(3.14159, 2)")      // 3.14
evaluate("=ABS(-42)")               // 42.0
evaluate("=MOD(17, 5)")             // 2.0
evaluate("=POWER(2, 8)")            // 256.0
evaluate("=SQRT(16)")               // 4.0
evaluate("=SIN(PI()/2)")            // 1.0
evaluate("=PRODUCT(2, 3, 4)")       // 24.0
evaluate("=SUMPRODUCT({1;2}, {3;4})")  // 11.0
```

See <doc:FormulaReference/Mathematical> for all 67 functions.

### Statistical Functions (134 functions)

Averages, variance, distributions, regression, forecasting.

**Examples:**
```swift
evaluate("=STDEV.S(1, 2, 3, 4, 5)")     // 1.58...
evaluate("=VAR.P(1, 2, 3, 4, 5)")       // 2.0
evaluate("=MEDIAN(1, 2, 3, 4, 5)")      // 3.0
evaluate("=PERCENTILE.INC({1;2;3}, 0.5)")  // 2.0
evaluate("=QUARTILE.INC({1;2;3;4}, 1)") // 1.75
evaluate("=NORM.DIST(0, 0, 1, TRUE)")   // 0.5
evaluate("=T.DIST(2, 5, TRUE)")         // CDF value
evaluate("=CORREL({1;2;3}, {2;4;6})")   // 1.0 (perfect correlation)
```

See <doc:FormulaReference/Statistical> for all 134 functions.

### Text Functions (36 functions)

String manipulation, search, formatting.

**Examples:**
```swift
evaluate("=CONCATENATE(\"Hello\", \" \", \"World\")")  // "Hello World"
evaluate("=LEFT(\"Hello\", 3)")         // "Hel"
evaluate("=RIGHT(\"World\", 3)")        // "rld"
evaluate("=MID(\"Excel\", 2, 3)")       // "xce"
evaluate("=LEN(\"Hello\")")             // 5.0
evaluate("=UPPER(\"hello\")")           // "HELLO"
evaluate("=LOWER(\"WORLD\")")           // "world"
evaluate("=FIND(\"l\", \"Hello\")")     // 3.0
evaluate("=SUBSTITUTE(\"Hi\", \"i\", \"ello\")")  // "Hello"
evaluate("=TRIM(\"  space  \")")        // "space"
```

See <doc:FormulaReference/Text> for all 36 functions.

### Date & Time Functions (24 functions)

Date arithmetic, extraction, working days.

**Examples:**
```swift
evaluate("=DATE(2024, 1, 15)")          // Serial number for Jan 15, 2024
evaluate("=YEAR(DATE(2024, 1, 15))")    // 2024
evaluate("=MONTH(DATE(2024, 1, 15))")   // 1
evaluate("=DAY(DATE(2024, 1, 15))")     // 15
evaluate("=TODAY()")                    // Current date serial
evaluate("=NOW()")                      // Current datetime serial
evaluate("=WEEKDAY(DATE(2024, 1, 15))") // Day of week
evaluate("=NETWORKDAYS(start, end)")    // Working days between dates
```

See <doc:FormulaReference/DateTime> for all 24 functions.

### Financial Functions (52 functions)

Present value, future value, depreciation, rates.

**Examples:**
```swift
evaluate("=PV(0.05, 10, -100, 0, 0)")   // Present value
evaluate("=FV(0.05, 10, -100, 0, 0)")   // Future value
evaluate("=PMT(0.05, 10, 1000, 0, 0)")  // Payment per period
evaluate("=RATE(10, -100, 1000, 0, 0)") // Interest rate
evaluate("=NPER(0.05, -100, 1000)")     // Number of periods
evaluate("=NPV(0.1, 100, 200, 300)")    // Net present value
evaluate("=IRR({-100, 30, 40, 50})")    // Internal rate of return
evaluate("=SLN(1000, 100, 10)")         // Straight-line depreciation
```

See <doc:FormulaReference/Financial> for all 52 functions.

### Logical Functions (16 functions)

Conditional logic, boolean operations.

**Examples:**
```swift
evaluate("=IF(10 > 5, \"Yes\", \"No\")")  // "Yes"
evaluate("=AND(TRUE, TRUE)")              // TRUE
evaluate("=OR(FALSE, TRUE)")              // TRUE
evaluate("=NOT(FALSE)")                   // TRUE
evaluate("=XOR(TRUE, FALSE)")             // TRUE
evaluate("=IFERROR(1/0, \"Error\")")      // "Error"
evaluate("=IFNA(NA(), \"Not Available\")") // "Not Available"
```

See <doc:FormulaReference/Logical> for all 16 functions.

### Lookup & Reference Functions (21 functions)

Table lookups, cell references, index/match.

**Examples:**
```swift
// VLOOKUP: Vertical lookup in a table
evaluate("=VLOOKUP(\"Apple\", A1:B10, 2, FALSE)")

// HLOOKUP: Horizontal lookup
evaluate("=HLOOKUP(\"Q1\", A1:D5, 2, FALSE)")

// INDEX: Get value at row/column
evaluate("=INDEX(A1:C3, 2, 3)")

// MATCH: Find position of value
evaluate("=MATCH(\"Apple\", A1:A10, 0)")

// OFFSET: Dynamic reference
evaluate("=OFFSET(A1, 1, 1, 3, 2)")

// CHOOSE: Select from list
evaluate("=CHOOSE(2, \"A\", \"B\", \"C\")")  // "B"
```

See <doc:FormulaReference/Lookup> for all 21 functions.

### Engineering Functions (56 functions)

Complex numbers, base conversions, Bessel functions.

**Examples:**
```swift
// Complex number operations
evaluate("=COMPLEX(3, 4)")              // "3+4i"
evaluate("=IMABS(\"3+4i\")")            // 5.0
evaluate("=IMREAL(\"3+4i\")")           // 3.0
evaluate("=IMAGINARY(\"3+4i\")")        // 4.0
evaluate("=IMSUM(\"3+4i\", \"1+2i\")")  // "4+6i"

// Base conversions
evaluate("=DEC2HEX(255)")               // "FF"
evaluate("=HEX2DEC(\"FF\")")            // 255
evaluate("=BIN2DEC(\"1010\")")          // 10

// Bitwise operations
evaluate("=BITAND(5, 3)")               // 1
evaluate("=BITOR(5, 3)")                // 7
evaluate("=BITXOR(5, 3)")               // 6
```

See <doc:FormulaReference/Engineering> for all 56 functions.

### Database Functions (10 functions)

Query operations on structured data ranges.

**Examples:**
```swift
// DSUM: Sum values matching criteria
evaluate("=DSUM(database, \"Sales\", criteria)")

// DAVERAGE: Average values matching criteria
evaluate("=DAVERAGE(database, \"Price\", criteria)")

// DCOUNT: Count matching records
evaluate("=DCOUNT(database, \"Name\", criteria)")

// DMAX/DMIN: Max/min matching values
evaluate("=DMAX(database, \"Score\", criteria)")
evaluate("=DMIN(database, \"Score\", criteria)")
```

See <doc:FormulaReference/Database> for all 10 functions.

### Information Functions (23 functions)

Type checking, error detection, environment info.

**Examples:**
```swift
evaluate("=ISBLANK(A1)")                // TRUE/FALSE
evaluate("=ISNUMBER(42)")               // TRUE
evaluate("=ISTEXT(\"Hello\")")          // TRUE
evaluate("=ISERROR(1/0)")               // TRUE
evaluate("=ISNA(NA())")                 // TRUE
evaluate("=TYPE(42)")                   // 1 (number)
evaluate("=TYPE(\"text\")")             // 2 (text)
evaluate("=N(TRUE)")                    // 1
evaluate("=CELL(\"address\", A1)")      // "$A$1"
```

See <doc:FormulaReference/Information> for all 23 functions.

### Compatibility Functions (14 functions)

Alternative names for newer Excel functions.

**Examples:**
```swift
evaluate("=FORECAST.LINEAR(x, known_y, known_x)")
evaluate("=PERCENTILE.INC(array, k)")
evaluate("=PERCENTRANK.INC(array, x)")
evaluate("=QUARTILE.INC(array, quart)")
evaluate("=RANK.EQ(number, ref, order)")
```

See <doc:FormulaReference/Compatibility> for all 14 functions.

### Web & Service Functions (14 functions - Stubs)

Web services and external data (not implemented, return #N/A):

```swift
evaluate("=WEBSERVICE(url)")            // #N/A (stub)
evaluate("=FILTERXML(xml, xpath)")      // #N/A (stub)
evaluate("=ENCODEURL(text)")            // #N/A (stub)
```

Also includes **LAMBDA functions** (stubs):
- LAMBDA, LET, MAP, REDUCE, SCAN, etc.

See <doc:FormulaReference/WebService> for all 14 stub functions.

## Error Handling

The formula engine propagates errors exactly like Excel:

### Error Types

```swift
public enum FormulaError {
    case ref      // #REF! - Invalid cell reference
    case value    // #VALUE! - Wrong value type
    case div0     // #DIV/0! - Division by zero
    case na       // #N/A - Value not available
    case name     // #NAME? - Unrecognized function name
    case num      // #NUM! - Invalid numeric value
    case null     // #NULL! - Invalid range intersection
}
```

### Error Propagation

Errors propagate through calculations:

```swift
// Division by zero
evaluate("=10 / 0")           // #DIV/0!

// Error propagates
evaluate("=SUM(10, 1/0, 5)")  // #DIV/0!

// IFERROR catches errors
evaluate("=IFERROR(1/0, 0)")  // 0

// Error in nested function
evaluate("=SQRT(-1)")         // #NUM!
evaluate("=ABS(SQRT(-1))")    // #NUM! (propagated)
```

### Function-Specific Errors

```swift
// Invalid reference
evaluate("=VLOOKUP(\"X\", A1:B10, 2, FALSE)")  // #N/A if not found

// Invalid argument
evaluate("=LOG(-1)")          // #NUM!

// Type mismatch
evaluate("=SUM(\"text\")")    // #VALUE!

// Missing function
evaluate("=NOTAFUNCTION()")   // #NAME?
```

## Range Operations

Ranges are central to Excel formulas:

### Range Creation

```swift
// Parse range
let parser = FormulaParser()
let ast = try parser.parse("=SUM(A1:C10)")

// AST contains:
// FunctionCall("SUM", args: [
//     .range(
//         CellReference(column: "A", row: 1),
//         CellReference(column: "C", row: 10)
//     )
// ])
```

### Range Evaluation

Ranges evaluate to 2D arrays:

```swift
// Range becomes array
let range = evaluate("A1:B2")
// [[.number(1), .number(2)],
//  [.number(3), .number(4)]]

// Functions process arrays
evaluate("=SUM(A1:B2)")      // 10 (sum of all values)
evaluate("=AVERAGE(A1:B2)")  // 2.5 (average of all values)
evaluate("=COUNT(A1:B2)")    // 4 (count of numeric values)
```

### Cross-Sheet References

```swift
// Reference cells in another sheet
evaluate("=Sheet2!A1")
evaluate("=SUM(Sheet2!A1:A10)")
evaluate("=Sheet2!A1 + Sheet3!B5")
```

## Array Formulas

Support for array constants and array results:

### Array Constants

```swift
// Vertical array
evaluate("=SUM({1;2;3})")    // 6

// Horizontal array
evaluate("=SUM({1,2,3})")    // 6

// 2D array
evaluate("=SUM({1,2;3,4})")  // 10
// [{1, 2},
//  {3, 4}]
```

### Array Results

```swift
// Function returns array
let result = evaluate("=FREQUENCY({1;2;3;4}, {2})")
// Result is array of frequencies

// Array in nested function
evaluate("=SUM(FREQUENCY({1;2;3;4}, {2}))")
```

## Performance Characteristics

### Function Complexity

| Function Type | Complexity | Notes |
|--------------|-----------|-------|
| Math | O(1) | Simple arithmetic |
| Aggregation | O(n) | SUM, AVERAGE, COUNT |
| Statistical | O(n log n) | Sorting for MEDIAN, PERCENTILE |
| Lookup | O(n) or O(log n) | VLOOKUP: linear or binary |
| Text | O(n) | String operations |
| Array | O(n × m) | 2D array processing |

### Optimization Tips

```swift
// ✅ Good: Reuse evaluator
var evaluator = FormulaEvaluator()
for formula in formulas {
    let result = evaluator.evaluate(formula)
}

// ❌ Bad: Create new evaluator each time
for formula in formulas {
    var evaluator = FormulaEvaluator()  // Unnecessary allocation
    let result = evaluator.evaluate(formula)
}
```

## Extending the Formula Engine

### Custom Functions

You can add custom functions to the evaluator:

```swift
extension FormulaEvaluator {
    // Add custom function
    mutating func registerCustom(_ name: String, impl: @escaping ([FormulaValue]) -> FormulaValue) {
        // Implementation would add to function registry
    }
}

// Usage
evaluator.registerCustom("MYFUNCTION") { args in
    guard case .number(let x) = args[0] else {
        return .error("VALUE")
    }
    return .number(x * 2)  // Example: double the input
}

evaluate("=MYFUNCTION(21)")  // 42
```

### Custom Cell Resolvers

Provide custom cell resolution logic:

```swift
let evaluator = FormulaEvaluator { cellRef in
    // Custom logic to resolve cell values
    // Could fetch from database, web API, etc.
    if cellRef.column == "A" && cellRef.row == 1 {
        return .number(42)
    }
    return nil
}
```

## Implementation Status

### Coverage Statistics

- **Total Functions**: 467
- **Full Implementations**: 453 (97%)
- **Stubs**: 14 (3%)

### Stub Functions (Return #N/A)

**LAMBDA Functions** (7):
- LAMBDA, LET, MAP, REDUCE, SCAN, MAKEARRAY, BYCOL, BYROW

**Cube/OLAP Functions** (6):
- CUBEVALUE, CUBEMEMBER, CUBESET, CUBERANKEDMEMBER, CUBEKPIMEMBER, CUBESETCOUNT

**Web Functions** (1):
- WEBSERVICE

All other 453 functions have complete, tested implementations.

## Testing

The formula engine has extensive test coverage:

```bash
# Run formula tests
swift test --filter FormulaEvaluatorTests

# Test specific function
swift test --filter testSUM

# Performance benchmarks
swift test --filter FormulaPerformanceTests
```

Located at: `Tests/CuneiformTests/FormulaEvaluatorTests.swift:1`

### Test Categories

- **Basic Operations**: Binary operators, precedence
- **Function Tests**: All 467 functions tested
- **Error Propagation**: Error handling in all contexts
- **Range Operations**: Cell ranges and arrays
- **Edge Cases**: Empty cells, missing references, type coercion
- **Performance**: Benchmarks for common operations

## Examples

### Data Analysis

```swift
var evaluator = FormulaEvaluator()

// Set up data
for i in 1...10 {
    evaluator.setCell("A\(i)", value: .number(Double(i * 10)))
}

// Compute statistics
let sum = evaluator.evaluate("=SUM(A1:A10)")       // 550
let avg = evaluator.evaluate("=AVERAGE(A1:A10)")   // 55
let min = evaluator.evaluate("=MIN(A1:A10)")       // 10
let max = evaluator.evaluate("=MAX(A1:A10)")       // 100
let std = evaluator.evaluate("=STDEV.S(A1:A10)")   // ~30.28
```

### Financial Calculations

```swift
// Loan payment calculation
let payment = evaluator.evaluate("=PMT(0.05/12, 30*12, -200000)")
// Monthly payment for $200k loan at 5% APR over 30 years

// Investment return
let fv = evaluator.evaluate("=FV(0.08, 10, -1000, 0, 0)")
// Future value of $1000/year investment at 8% for 10 years
```

### Text Processing

```swift
evaluator.setCell("A1", value: .text("  John Doe  "))

let cleaned = evaluator.evaluate("=TRIM(A1)")              // "John Doe"
let upper = evaluator.evaluate("=UPPER(TRIM(A1))")         // "JOHN DOE"
let first = evaluator.evaluate("=LEFT(TRIM(A1), 4)")       // "John"
let length = evaluator.evaluate("=LEN(TRIM(A1))")          // 8
```

## See Also

- <doc:FormulaReference> - Complete function reference with examples
- <doc:Architecture> - How the formula engine fits into Cuneiform
- ``FormulaEvaluator`` - API reference
- ``FormulaParser`` - Parser API reference
- ``FormulaExpression`` - AST node types
- ``FormulaValue`` - Result value types
