# Database Functions

Excel-compatible database functions for querying and analyzing structured data ranges.

## Overview

Cuneiform provides a complete set of database functions that perform SQL-like operations on structured Excel ranges. These functions allow you to query, filter, aggregate, and analyze data organized in a database-style format with column headers.

Database functions operate on three key components:
1. **Database** - A range of cells organized as a table with column headers in the first row
2. **Field** - The column to perform calculations on (specified by name or column number)
3. **Criteria** - A range that defines filtering conditions (with headers matching database columns)

All database functions follow the pattern: `D<OPERATION>(database, field, criteria)` where the operation mirrors standard Excel functions like SUM, AVERAGE, COUNT, etc.

## Quick Reference

### Aggregation Functions
- ``DSUM`` - Sum values matching criteria
- ``DAVERAGE`` - Average values matching criteria
- ``DCOUNT`` - Count cells with numbers matching criteria
- ``DCOUNTA`` - Count non-empty cells matching criteria
- ``DMAX`` - Find maximum value matching criteria
- ``DMIN`` - Find minimum value matching criteria
- ``DPRODUCT`` - Multiply values matching criteria

### Retrieval Functions
- ``DGET`` - Extract single value matching criteria

### Statistical Functions
- ``DSTDEV`` - Standard deviation of values matching criteria (not yet implemented)
- ``DVAR`` - Variance of values matching criteria (not yet implemented)

## Database Structure

Database functions require data organized in a specific format:

**Example Database Range (A1:D6)**:
```
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10
Pear      9         8      8
Cherry    8         9      6
```

**Example Criteria Range (F1:G2)**:
```
Tree      Height
Apple     >12
```

This criteria would match rows where Tree = "Apple" AND Height > 12.

## Function Details

### DSUM

Calculates the sum of values in a database column that match specified criteria.

**Syntax:** `DSUM(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to sum - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The sum of matching values

**Examples:**
```swift
// Sum yield for all apple trees over 12 feet tall
let result1 = evaluator.evaluate("=DSUM(A1:D6, \"Yield\", F1:G2)")  // 24

// Same using column number instead of name
let result2 = evaluator.evaluate("=DSUM(A1:D6, 4, F1:G2)")  // 24

// Multiple criteria: Apple trees with age > 15
let result3 = evaluator.evaluate("=DSUM(A1:D6, \"Yield\", F1:G3)")  // 14
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:G2):
Tree      Height
Apple     >12
```

**Excel Documentation:** [DSUM function](https://support.microsoft.com/en-us/office/dsum-function-53181285-0c4b-4f5a-aaa3-529a322be41b)

**Implementation Status:** ✅ Full implementation

---

### DAVERAGE

Calculates the average of values in a database column that match specified criteria.

**Syntax:** `DAVERAGE(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to average - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The average of matching values, or #DIV/0! if no matches

**Examples:**
```swift
// Average age of apple trees
let result1 = evaluator.evaluate("=DAVERAGE(A1:D6, \"Age\", F1:F2)")  // 17.5

// Average height of all cherry trees
let result2 = evaluator.evaluate("=DAVERAGE(A1:D6, \"Height\", F1:F2)")  // 10.5

// Average yield for trees taller than 12 feet
let result3 = evaluator.evaluate("=DAVERAGE(A1:D6, \"Yield\", F1:G2)")  // 11
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9

Criteria (F1:F2):
Tree
Apple
```

**Excel Documentation:** [DAVERAGE function](https://support.microsoft.com/en-us/office/daverage-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee)

**Implementation Status:** ✅ Full implementation

---

### DCOUNT

Counts the cells containing numbers in a database column that match specified criteria.

**Syntax:** `DCOUNT(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to count - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The count of cells with numeric values

**Examples:**
```swift
// Count number of apple tree records with numeric height
let result1 = evaluator.evaluate("=DCOUNT(A1:D6, \"Height\", F1:F2)")  // 2

// Count all trees with age > 12
let result2 = evaluator.evaluate("=DCOUNT(A1:D6, \"Age\", F1:G2)")  // 3

// Count cherry trees with yield data
let result3 = evaluator.evaluate("=DCOUNT(A1:D6, 4, F1:F2)")  // 2
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Apple
```

**Note:** DCOUNT only counts cells with numeric values. Use DCOUNTA to count all non-empty cells.

**Excel Documentation:** [DCOUNT function](https://support.microsoft.com/en-us/office/dcount-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1)

**Implementation Status:** ✅ Full implementation

---

### DCOUNTA

Counts non-empty cells in a database column that match specified criteria.

**Syntax:** `DCOUNTA(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to count - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The count of non-empty cells

**Examples:**
```swift
// Count all apple tree records (including text)
let result1 = evaluator.evaluate("=DCOUNTA(A1:D6, \"Tree\", F1:F2)")  // 2

// Count non-empty yield values for tall trees
let result2 = evaluator.evaluate("=DCOUNTA(A1:D6, \"Yield\", F1:G2)")  // 3

// Count all pear tree records
let result3 = evaluator.evaluate("=DCOUNTA(A1:D6, 1, F1:F2)")  // 2
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Pear
```

**Note:** Unlike DCOUNT, DCOUNTA counts all non-empty cells, including text values. This is useful for counting records regardless of data type.

**Excel Documentation:** [DCOUNTA function](https://support.microsoft.com/en-us/office/dcounta-function-00232a6d-5a66-4a01-a25b-c1653fda1244)

**Implementation Status:** ✅ Full implementation

---

### DMAX

Returns the maximum value in a database column that matches specified criteria.

**Syntax:** `DMAX(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to evaluate - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The maximum value among matching records, or 0 if no matches

**Examples:**
```swift
// Find tallest apple tree
let result1 = evaluator.evaluate("=DMAX(A1:D6, \"Height\", F1:F2)")  // 18

// Find maximum yield for all trees
let result2 = evaluator.evaluate("=DMAX(A1:D6, \"Yield\", F1:G2)")  // 14

// Find oldest tree of any type
let result3 = evaluator.evaluate("=DMAX(A1:D6, 3, F1:F2)")  // 20
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Apple
```

**Excel Documentation:** [DMAX function](https://support.microsoft.com/en-us/office/dmax-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2)

**Implementation Status:** ✅ Full implementation

---

### DMIN

Returns the minimum value in a database column that matches specified criteria.

**Syntax:** `DMIN(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to evaluate - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The minimum value among matching records, or 0 if no matches

**Examples:**
```swift
// Find shortest apple tree
let result1 = evaluator.evaluate("=DMIN(A1:D6, \"Height\", F1:F2)")  // 14

// Find minimum age for cherry trees
let result2 = evaluator.evaluate("=DMIN(A1:D6, \"Age\", F1:F2)")  // 9

// Find lowest yield among tall trees
let result3 = evaluator.evaluate("=DMIN(A1:D6, 4, F1:G2)")  // 9
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Apple
```

**Excel Documentation:** [DMIN function](https://support.microsoft.com/en-us/office/dmin-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3)

**Implementation Status:** ✅ Full implementation

---

### DGET

Extracts a single value from a database column that matches specified criteria.

**Syntax:** `DGET(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to retrieve - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Value - The matching cell value, #VALUE! if no match, #NUM! if multiple matches

**Examples:**
```swift
// Get height of a specific tree
let result1 = evaluator.evaluate("=DGET(A1:D6, \"Height\", F1:G3)")  // 18

// Get yield for tree with specific criteria
let result2 = evaluator.evaluate("=DGET(A1:D6, \"Yield\", F1:G3)")  // 14

// Error: Multiple apple trees exist
let result3 = evaluator.evaluate("=DGET(A1:D6, \"Height\", F1:F2)")  // #NUM!
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:G3):
Tree      Age
Apple     20
```

**Note:** DGET is unique among database functions because it expects exactly one matching record. If zero records match, it returns #VALUE!. If multiple records match, it returns #NUM!.

**Excel Documentation:** [DGET function](https://support.microsoft.com/en-us/office/dget-function-455568bf-4eef-45f7-90f0-ec250d00892e)

**Implementation Status:** ✅ Full implementation

---

### DPRODUCT

Multiplies the values in a database column that match specified criteria.

**Syntax:** `DPRODUCT(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to multiply - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The product of matching values, or 0 if no matches

**Examples:**
```swift
// Multiply all apple tree yields
let result1 = evaluator.evaluate("=DPRODUCT(A1:D6, \"Yield\", F1:F2)")  // 140

// Product of heights for tall trees
let result2 = evaluator.evaluate("=DPRODUCT(A1:D6, \"Height\", F1:G2)")  // 3276

// Product of ages for cherry trees
let result3 = evaluator.evaluate("=DPRODUCT(A1:D6, 3, F1:F2)")  // 126
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Apple
```

**Excel Documentation:** [DPRODUCT function](https://support.microsoft.com/en-us/office/dproduct-function-4f96b13e-d49c-47a7-b769-22f6d017cb31)

**Implementation Status:** ✅ Full implementation

---

### DSTDEV

Estimates the standard deviation of a population sample based on values in a database column that match specified criteria.

**Syntax:** `DSTDEV(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to analyze - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The sample standard deviation

**Examples:**
```swift
// Calculate standard deviation of apple tree heights
let result1 = evaluator.evaluate("=DSTDEV(A1:D6, \"Height\", F1:F2)")

// Standard deviation of yields for tall trees
let result2 = evaluator.evaluate("=DSTDEV(A1:D6, \"Yield\", F1:G2)")
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Apple
```

**Excel Documentation:** [DSTDEV function](https://support.microsoft.com/en-us/office/dstdev-function-026b8c73-616d-4b5e-b072-241871c4ab96)

**Implementation Status:** 🔄 Stub implementation (returns #CALC! error)

---

### DVAR

Estimates the variance of a population sample based on values in a database column that match specified criteria.

**Syntax:** `DVAR(database, field, criteria)`

**Parameters:**
- `database`: The range of cells that makes up the database, including column headers
- `field`: The column to analyze - either column name (text) or column number (1-based)
- `criteria`: The range of cells containing the filtering conditions

**Returns:** Number - The sample variance

**Examples:**
```swift
// Calculate variance of apple tree heights
let result1 = evaluator.evaluate("=DVAR(A1:D6, \"Height\", F1:F2)")

// Variance of ages for cherry trees
let result2 = evaluator.evaluate("=DVAR(A1:D6, \"Age\", F1:F2)")
```

**Database Setup:**
```
Database (A1:D6):
Tree      Height    Age    Yield
Apple     18        20     14
Pear      12        12     10
Cherry    13        14     9
Apple     14        15     10

Criteria (F1:F2):
Tree
Cherry
```

**Excel Documentation:** [DVAR function](https://support.microsoft.com/en-us/office/dvar-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1)

**Implementation Status:** 🔄 Stub implementation (returns #CALC! error)

---

## Criteria Syntax

Database functions support flexible criteria matching:

### Exact Match
```
Tree
Apple
```
Matches rows where Tree equals "Apple" (case-insensitive)

### Numeric Comparisons
```
Height    Age
>12       <15
```
Matches rows where Height > 12 AND Age < 15

Supported operators: `>`, `<`, `>=`, `<=`, `<>`, `=`

### Multiple Values (OR Logic)
```
Tree
Apple
Cherry
```
Matches rows where Tree = "Apple" OR Tree = "Cherry"

### Multiple Columns (AND Logic)
```
Tree      Height
Apple     >15
```
Matches rows where Tree = "Apple" AND Height > 15

### Empty Criteria
An empty criteria range (only headers, no conditions) matches all records.

### Wildcards
```
Tree
A*
```
Matches rows where Tree starts with "A" (Apple, etc.)

Supported wildcards: `*` (any characters), `?` (single character)

## Common Patterns

### Counting Records
```swift
// Count all records matching criteria
let totalRecords = evaluator.evaluate("=DCOUNTA(A1:D6, \"Tree\", F1:F2)")

// Count only numeric values
let numericRecords = evaluator.evaluate("=DCOUNT(A1:D6, \"Height\", F1:F2)")
```

### Statistical Analysis
```swift
// Basic statistics for filtered data
let sum = evaluator.evaluate("=DSUM(A1:D6, \"Yield\", F1:F2)")
let avg = evaluator.evaluate("=DAVERAGE(A1:D6, \"Yield\", F1:F2)")
let min = evaluator.evaluate("=DMIN(A1:D6, \"Yield\", F1:F2)")
let max = evaluator.evaluate("=DMAX(A1:D6, \"Yield\", F1:F2)")
let count = evaluator.evaluate("=DCOUNT(A1:D6, \"Yield\", F1:F2)")
```

### Dynamic Criteria
```swift
// Build criteria programmatically
sheet.writeValue("Tree", to: "F1")
sheet.writeValue("Apple", to: "F2")
sheet.writeFormula("=DSUM(A1:D6, \"Yield\", F1:F2)", to: "H1")
```

### Multiple Criteria Sets
```swift
// Complex filtering with multiple conditions
// Criteria: (Tree=Apple AND Height>15) OR (Tree=Cherry AND Age>12)
// Use separate criteria ranges or combine with OR function
```

## Performance Considerations

Database functions are optimized for small to medium-sized datasets (up to ~10,000 rows). For larger datasets:

1. **Use Filtering**: Pre-filter data before applying database functions
2. **Index Columns**: Ensure field names match exactly to avoid scan overhead
3. **Minimize Criteria**: Simpler criteria ranges evaluate faster
4. **Cache Results**: Store intermediate results rather than recalculating

## Error Handling

Database functions return standard Excel errors:

| Error | Cause | Solution |
|-------|-------|----------|
| `#VALUE!` | Invalid database/criteria range | Verify ranges have headers |
| `#VALUE!` | Field name not found | Check spelling and case |
| `#VALUE!` | DGET found no matches | Adjust criteria or use IFERROR |
| `#NUM!` | DGET found multiple matches | Make criteria more specific |
| `#DIV/0!` | DAVERAGE with no matches | Ensure at least one match exists |
| `#NAME?` | Unknown field name | Verify column header exists |

## See Also

- ``StatisticalFunctions`` - Standard statistical analysis
- ``LookupFunctions`` - VLOOKUP, XLOOKUP for record retrieval
- ``FilterFunctions`` - FILTER function for dynamic arrays
- ``AggregationFunctions`` - SUM, AVERAGE, COUNT without criteria
- ``FormulaEvaluator`` - The core formula evaluation engine
