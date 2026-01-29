# Formula Reference

Complete reference for all 467 Excel formula functions supported by Cuneiform.

## Overview

Cuneiform's formula engine supports 467 Excel functions across 12 categories, providing 97% full implementations. Use these functions in formulas when creating workbooks or evaluating expressions programmatically.

### Formula Categories

Browse functions by category:

- **<doc:Mathematical>** - Mathematical and trigonometric functions (80 functions)
- **<doc:Statistical>** - Statistical analysis and distributions (100 functions)
- **<doc:Text>** - Text manipulation and formatting (40 functions)
- **<doc:DateTime>** - Date and time operations (25 functions)
- **<doc:Financial>** - Financial calculations and analysis (55 functions)
- **<doc:Logical>** - Logical operations and conditions (10 functions)
- **<doc:Lookup>** - Lookup, reference, and dynamic arrays (35 functions)
- **<doc:Engineering>** - Engineering and complex number operations (60 functions)
- **<doc:Database>** - Database-style queries on ranges (10 functions)
- **<doc:Information>** - Information and type checking (30 functions)
- **<doc:Compatibility>** - Compatibility with older Excel versions (15 functions)
- **<doc:WebService>** - Web and external data functions (7 functions)

### Using Formulas

Create formulas when writing workbooks:

```swift
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Calculations")

writer.modifySheet(at: sheetIndex) { sheet in
    // Basic arithmetic
    sheet.writeFormula("A1 + A2", to: "A3")
    
    // Built-in functions
    sheet.writeFormula("SUM(A1:A10)", to: "A11")
    sheet.writeFormula("AVERAGE(A1:A10)", to: "A12")
    
    // Nested functions
    sheet.writeFormula("IF(A1>100, SUM(B1:B10), 0)", to: "C1")
    
    // Cross-sheet references
    sheet.writeFormula("SUM(Sheet1!A1:A10)", to: "D1")
}
```

Evaluate formulas programmatically:

```swift
let evaluator = FormulaEvaluator(workbook: workbook)

// Simple evaluation
let result = try evaluator.evaluate("=SUM(1, 2, 3)")  // 6.0

// Reference worksheet cells
let total = try evaluator.evaluate("=SUM(A1:A10)")

// Complex expressions
let analysis = try evaluator.evaluate("=IF(AVERAGE(A1:A10) > 50, \"High\", \"Low\")")
```

### Implementation Status

**Implementation Coverage**:
- ✅ 452 functions with full implementations (97%)
- ⚠️ 10 functions with partial implementations (2%)
- 🔄 5 functions are stubs for future implementation (1%)

**Full Implementation** means the function behaves identically to Excel in all tested scenarios.

**Partial Implementation** means the function works for common cases but may not support all Excel edge cases or optional parameters.

**Stub Implementation** means the function is recognized but returns a placeholder value or limited functionality.

### Formula Syntax

Cuneiform supports Excel's formula syntax:

**Operators**:
- Arithmetic: `+`, `-`, `*`, `/`, `^`, `%`
- Comparison: `=`, `<>`, `<`, `>`, `<=`, `>=`
- Text: `&` (concatenation)
- Reference: `:` (range), `,` (union), ` ` (intersection)

**References**:
- Cell: `A1`, `$A$1` (absolute), `A$1` (mixed)
- Range: `A1:B10`
- Sheet: `Sheet1!A1`, `'Sheet Name'!A1:B10`
- Named: `SalesData`, `TaxRate`

**Data Types**:
- Numbers: `42`, `3.14`, `1.5E+10`
- Text: `"Hello"`, `"Don't"`
- Boolean: `TRUE`, `FALSE`
- Error: `#REF!`, `#VALUE!`, `#N/A`, `#DIV/0!`, `#NUM!`, `#NAME?`, `#NULL!`
- Arrays: `{1, 2, 3}`, `{1; 2; 3}` (rows)

### Error Handling

Formula errors are propagated through calculations:

| Error | Meaning | Common Causes |
|-------|---------|---------------|
| `#REF!` | Invalid reference | Deleted cell, invalid range |
| `#VALUE!` | Wrong type | Text where number expected |
| `#N/A` | Value not available | VLOOKUP not found |
| `#DIV/0!` | Division by zero | Divide by zero or empty cell |
| `#NUM!` | Invalid numeric value | Argument out of range |
| `#NAME?` | Unrecognized name | Unknown function or name |
| `#NULL!` | Null intersection | Range intersection is empty |

Example error handling:

```swift
let evaluator = FormulaEvaluator(workbook: workbook)

do {
    let result = try evaluator.evaluate("=1/0")
    // Result would be .error("DIV/0")
} catch {
    print("Evaluation error: \(error)")
}
```

### See Also

- <doc:FormulaEngine> - Architecture and evaluation details
- ``FormulaEvaluator`` - Programmatic formula evaluation
- ``FormulaParser`` - Formula parsing and AST generation
- <doc:WritingWorkbooks> - Creating formulas in workbooks
