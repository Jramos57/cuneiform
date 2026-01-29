# Lookup & Reference Functions

Lookup values in arrays, reference cells dynamically, and work with dynamic arrays.

## Overview

Cuneiform provides comprehensive lookup and reference functions for finding data, dynamic cell references, and modern dynamic array operations. These functions are essential for data analysis, report generation, and building flexible spreadsheet models.

The library implements ~35 lookup and reference functions compatible with Excel, including traditional lookup functions (VLOOKUP, HLOOKUP), modern Excel 365 functions (XLOOKUP, FILTER, SORT), and dynamic array operations.

### Quick Reference

**Traditional Lookup Functions:**
- ``VLOOKUP`` - Vertical lookup in table ✅
- ``HLOOKUP`` - Horizontal lookup in table ✅
- ``LOOKUP`` - Vector or array lookup ✅
- ``INDEX`` - Return value at row/column ✅
- ``MATCH`` - Find position of value ✅

**Modern Lookup (Excel 365):**
- ``XLOOKUP`` - Modern replacement for VLOOKUP/HLOOKUP ⚠️
- ``XMATCH`` - Enhanced MATCH function ⚠️

**Reference Functions:**
- ``INDIRECT`` - Reference from text string ⚠️
- ``OFFSET`` - Reference offset from base ⚠️
- ``ADDRESS`` - Create cell address text ✅
- ``COLUMN`` - Get column number ✅
- ``ROW`` - Get row number ✅
- ``COLUMNS`` - Count columns in range ✅
- ``ROWS`` - Count rows in range ✅
- ``AREAS`` - Count areas in reference ✅

**Selection Functions:**
- ``CHOOSE`` - Pick value from list ✅
- ``CHOOSECOLS`` - Select columns from array ✅
- ``CHOOSEROWS`` - Select rows from array ✅

**Dynamic Array Functions (Excel 365):**
- ``FILTER`` - Filter array by criteria ✅
- ``SORT`` - Sort array ✅
- ``SORTBY`` - Sort by another array ✅
- ``UNIQUE`` - Extract unique values ✅
- ``SEQUENCE`` - Generate number sequence ✅
- ``RANDARRAY`` - Generate random array ✅
- ``TAKE`` - Take first/last rows/columns ✅
- ``DROP`` - Drop first/last rows/columns ✅
- ``EXPAND`` - Expand array to dimensions ✅
- ``TRANSPOSE`` - Transpose rows/columns ✅

**Array Stacking:**
- ``VSTACK`` - Stack arrays vertically ✅
- ``HSTACK`` - Stack arrays horizontally ✅
- ``TOCOL`` - Convert to single column ✅
- ``TOROW`` - Convert to single row ✅
- ``WRAPCOLS`` - Wrap into columns ✅
- ``WRAPROWS`` - Wrap into rows ✅

**Legend:**
- ✅ Full implementation
- ⚠️ Partial implementation
- 🔄 Stub/placeholder

## Traditional Lookup Functions

### VLOOKUP

Searches for a value in the first column of a table and returns a value in the same row from another column.

**Syntax:** `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

**Parameters:**
- `lookup_value`: Value to search for in the first column
- `table_array`: Table of data to search
- `col_index_num`: Column number to return (1-based)
- `range_lookup` *(optional)*: TRUE for approximate match, FALSE for exact match (default: TRUE)

**Returns:** Value from the specified column, or #N/A if not found

**Examples:**
```swift
// Exact match lookup
let price = try evaluator.evaluate("=VLOOKUP(\"SKU123\", A2:D100, 3, FALSE)")

// Approximate match (assumes sorted first column)
let taxRate = try evaluator.evaluate("=VLOOKUP(A1, TaxTable, 2, TRUE)")

// With error handling
let result = try evaluator.evaluate("=IFERROR(VLOOKUP(A1, Products, 4, FALSE), \"Not Found\")")

// Using in Swift code
sheet.writeFormula("VLOOKUP(A2, PriceList!A:C, 2, FALSE)", to: "B2")
```

**Excel Documentation:** [VLOOKUP function](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)

**Implementation Status:** ✅ Full implementation

**Notes:**
- For exact matches, set `range_lookup` to FALSE
- For approximate matches, first column must be sorted ascending
- Returns #N/A error if no match found
- Consider using XLOOKUP for more flexibility

**See Also:** ``HLOOKUP``, ``XLOOKUP``, ``INDEX``, ``MATCH``

---

### HLOOKUP

Searches for a value in the first row of a table and returns a value in the same column from another row.

**Syntax:** `HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

**Parameters:**
- `lookup_value`: Value to search for in the first row
- `table_array`: Table of data to search
- `row_index_num`: Row number to return (1-based)
- `range_lookup` *(optional)*: TRUE for approximate match, FALSE for exact match (default: TRUE)

**Returns:** Value from the specified row, or #N/A if not found

**Examples:**
```swift
// Lookup month data in horizontal table
let janSales = try evaluator.evaluate("=HLOOKUP(\"Jan\", A1:M5, 3, FALSE)")

// Approximate match
let rate = try evaluator.evaluate("=HLOOKUP(A1, RateTable, 2, TRUE)")

// Using in formulas
sheet.writeFormula("HLOOKUP(\"Q2\", QuarterlyData, 4, FALSE)", to: "B10")
```

**Excel Documentation:** [HLOOKUP function](https://support.microsoft.com/en-us/office/hlookup-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f)

**Implementation Status:** ✅ Full implementation

**Notes:**
- Similar to VLOOKUP but searches horizontally
- Less commonly used than VLOOKUP
- First row must be sorted for approximate matches

**See Also:** ``VLOOKUP``, ``XLOOKUP``

---

### INDEX

Returns the value at a specific row and column intersection in a range.

**Syntax:** `INDEX(array, row_num, [column_num])`

**Parameters:**
- `array`: Array or range to index into
- `row_num`: Row number (1-based, 0 returns entire column)
- `column_num` *(optional)*: Column number (1-based, omit for single column)

**Returns:** Value at the specified position, or #REF! if out of range

**Examples:**
```swift
// Get value at specific position
let value = try evaluator.evaluate("=INDEX(A1:C10, 5, 2)")  // Row 5, Column 2

// Single column lookup
let name = try evaluator.evaluate("=INDEX(A1:A100, 25)")

// Combined with MATCH for flexible lookup
let result = try evaluator.evaluate("=INDEX(B2:B100, MATCH(\"Apple\", A2:A100, 0))")

// Return entire column (row_num = 0)
let column = try evaluator.evaluate("=INDEX(A1:D10, 0, 3)")

// Two-way lookup
sheet.writeFormula("INDEX(Data, MATCH(A1, Names, 0), MATCH(B1, Headers, 0))", to: "C1")
```

**Excel Documentation:** [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd)

**Implementation Status:** ✅ Full implementation

**Notes:**
- More flexible than VLOOKUP (can return any column)
- Commonly paired with MATCH for dynamic lookups
- Row 0 returns entire column, Column 0 returns entire row

**See Also:** ``MATCH``, ``VLOOKUP``, ``XLOOKUP``

---

### MATCH

Returns the relative position of a value in an array.

**Syntax:** `MATCH(lookup_value, lookup_array, [match_type])`

**Parameters:**
- `lookup_value`: Value to find
- `lookup_array`: Array to search
- `match_type` *(optional)*: 
  - `1` = largest value ≤ lookup_value (default, array must be sorted ascending)
  - `0` = exact match
  - `-1` = smallest value ≥ lookup_value (array must be sorted descending)

**Returns:** Position of the match (1-based), or #N/A if not found

**Examples:**
```swift
// Exact match
let position = try evaluator.evaluate("=MATCH(\"Apple\", A1:A100, 0)")

// Find position for two-way lookup
let result = try evaluator.evaluate("=INDEX(Data, MATCH(Product, Products, 0), MATCH(Month, Months, 0))")

// Approximate match (sorted data)
let bracket = try evaluator.evaluate("=MATCH(95, GradeBreaks, 1)")

// Using with INDEX
sheet.writeFormula("INDEX(Prices, MATCH(A1, ProductList, 0))", to: "B1")
```

**Excel Documentation:** [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a)

**Implementation Status:** ✅ Full implementation

**Notes:**
- Returns position, not the value itself
- Pair with INDEX for powerful flexible lookups
- More versatile than VLOOKUP/HLOOKUP combination

**See Also:** ``INDEX``, ``XMATCH``, ``VLOOKUP``

---

### LOOKUP

Looks up a value in a one-row or one-column range and returns a value from the same position in another range.

**Syntax:** `LOOKUP(lookup_value, lookup_vector, [result_vector])`

**Parameters:**
- `lookup_value`: Value to find
- `lookup_vector`: Single row or column to search
- `result_vector` *(optional)*: Single row or column of results (defaults to lookup_vector)

**Returns:** Value from result vector at matching position

**Examples:**
```swift
// Vector form
let result = try evaluator.evaluate("=LOOKUP(A1, B1:B10, C1:C10)")

// Array form (2D lookup)
let value = try evaluator.evaluate("=LOOKUP(95, {0,60,70,80,90}, {\"F\",\"D\",\"C\",\"B\",\"A\"})")

// Using in formulas
sheet.writeFormula("LOOKUP(A1, LookupRange, ResultRange)", to: "B1")
```

**Excel Documentation:** [LOOKUP function](https://support.microsoft.com/en-us/office/lookup-function-446d94af-663b-451d-8251-369d5e3864cb)

**Implementation Status:** ✅ Full implementation

**Notes:**
- Always assumes sorted data (uses approximate match)
- For exact matches, use VLOOKUP, HLOOKUP, or INDEX/MATCH
- Less commonly used than other lookup functions

**See Also:** ``VLOOKUP``, ``HLOOKUP``, ``INDEX``

---

## Modern Lookup Functions (Excel 365)

### XLOOKUP

Modern lookup function that replaces VLOOKUP, HLOOKUP, and INDEX/MATCH combinations.

**Syntax:** `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`

**Parameters:**
- `lookup_value`: Value to search for
- `lookup_array`: Array to search in
- `return_array`: Array to return values from
- `if_not_found` *(optional)*: Value if no match (default: #N/A)
- `match_mode` *(optional)*: 0=exact (default), -1=exact or next smaller, 1=exact or next larger, 2=wildcard
- `search_mode` *(optional)*: 1=first to last (default), -1=last to first, 2=binary ascending, -2=binary descending

**Returns:** Matching value from return_array

**Examples:**
```swift
// Basic exact match
let price = try evaluator.evaluate("=XLOOKUP(A1, Products, Prices)")

// With custom not-found message
let result = try evaluator.evaluate("=XLOOKUP(A1, SKUs, Names, \"Product not found\")")

// Wildcard search
let match = try evaluator.evaluate("=XLOOKUP(\"*Apple*\", Items, Prices, , 2)")

// Reverse search (last to first)
let last = try evaluator.evaluate("=XLOOKUP(A1, Dates, Values, , 0, -1)")

// Using in formulas
sheet.writeFormula("XLOOKUP(A2, ProductIDs, ProductNames, \"N/A\")", to: "B2")
```

**Excel Documentation:** [XLOOKUP function](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)

**Implementation Status:** ⚠️ Partial implementation (match_mode and search_mode not fully supported)

**Advantages over VLOOKUP:**
- Can look to the left (return column before lookup column)
- Default exact match
- Custom if_not_found value
- Can return entire rows or columns
- More intuitive syntax

**See Also:** ``VLOOKUP``, ``HLOOKUP``, ``XMATCH``, ``INDEX``, ``MATCH``

---

### XMATCH

Returns the relative position of an item in an array (enhanced version of MATCH).

**Syntax:** `XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])`

**Parameters:**
- `lookup_value`: Value to find
- `lookup_array`: Array to search
- `match_mode` *(optional)*: 0=exact (default), -1=exact or next smaller, 1=exact or next larger, 2=wildcard
- `search_mode` *(optional)*: 1=first to last (default), -1=last to first, 2=binary ascending, -2=binary descending

**Returns:** Position of match (1-based), or #N/A if not found

**Examples:**
```swift
// Basic usage
let position = try evaluator.evaluate("=XMATCH(\"Apple\", A1:A100)")

// Wildcard match
let pos = try evaluator.evaluate("=XMATCH(\"*tech*\", Categories, 2)")

// Reverse search
let lastPos = try evaluator.evaluate("=XMATCH(A1, DatesList, 0, -1)")

// With INDEX for flexible lookup
sheet.writeFormula("INDEX(Data, XMATCH(A1, IDs))", to: "B1")
```

**Excel Documentation:** [XMATCH function](https://support.microsoft.com/en-us/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312)

**Implementation Status:** ⚠️ Partial implementation (match_mode and search_mode not fully supported)

**See Also:** ``MATCH``, ``XLOOKUP``, ``INDEX``

---

## Reference Functions

### INDIRECT

Returns the reference specified by a text string.

**Syntax:** `INDIRECT(ref_text, [a1])`

**Parameters:**
- `ref_text`: Text string containing a cell reference
- `a1` *(optional)*: TRUE for A1 style (default), FALSE for R1C1 style

**Returns:** Value at the referenced cell, or #REF! if invalid

**Examples:**
```swift
// Reference cell by text
let value = try evaluator.evaluate("=INDIRECT(\"A1\")")

// Build dynamic reference
let result = try evaluator.evaluate("=INDIRECT(\"A\" & B1)")  // If B1=5, returns A5

// Reference another sheet
let sheetValue = try evaluator.evaluate("=INDIRECT(\"Sheet2!A1\")")

// Dynamic column reference
sheet.writeFormula("SUM(INDIRECT(A1 & \"1:\" & A1 & \"100\"))", to: "B1")
```

**Excel Documentation:** [INDIRECT function](https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261)

**Implementation Status:** ⚠️ Partial implementation (basic cell references only)

**Notes:**
- Useful for dynamic range references
- Always volatile (recalculates with every change)
- Can reference other sheets

**See Also:** ``OFFSET``, ``ADDRESS``

---

### OFFSET

Returns a reference offset from a starting reference.

**Syntax:** `OFFSET(reference, rows, cols, [height], [width])`

**Parameters:**
- `reference`: Starting reference
- `rows`: Number of rows to offset (positive=down, negative=up)
- `cols`: Number of columns to offset (positive=right, negative=left)
- `height` *(optional)*: Height of returned range
- `width` *(optional)*: Width of returned range

**Returns:** Reference offset from starting point, or #REF! if invalid

**Examples:**
```swift
// Offset by 2 rows, 1 column
let value = try evaluator.evaluate("=OFFSET(A1, 2, 1)")  // Returns C3

// Dynamic range for SUM
let sum = try evaluator.evaluate("=SUM(OFFSET(A1, 0, 0, 10, 1))")  // Sum A1:A10

// Create dynamic named range
let range = try evaluator.evaluate("=OFFSET(A1, 0, 0, COUNTA(A:A), 1)")
```

**Excel Documentation:** [OFFSET function](https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66)

**Implementation Status:** ⚠️ Partial implementation (simplified, needs sheet coordinate access)

**Notes:**
- Volatile function (recalculates frequently)
- Useful for creating dynamic ranges
- Often used with named ranges

**See Also:** ``INDIRECT``, ``INDEX``

---

### ADDRESS

Creates a cell address as text given row and column numbers.

**Syntax:** `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])`

**Parameters:**
- `row_num`: Row number
- `column_num`: Column number
- `abs_num` *(optional)*: 1=absolute (default), 2=absolute row/relative column, 3=relative row/absolute column, 4=relative
- `a1` *(optional)*: TRUE for A1 style (default), FALSE for R1C1 style
- `sheet_text` *(optional)*: Sheet name to include

**Returns:** Cell address as text

**Examples:**
```swift
// Basic address
let addr = try evaluator.evaluate("=ADDRESS(5, 3)")  // Returns "$C$5"

// Relative reference
let rel = try evaluator.evaluate("=ADDRESS(5, 3, 4)")  // Returns "C5"

// With sheet name
let sheet = try evaluator.evaluate("=ADDRESS(1, 1, 1, TRUE, \"Sheet2\")")  // Returns "Sheet2!$A$1"

// Use with INDIRECT for dynamic reference
sheet.writeFormula("INDIRECT(ADDRESS(ROW(), COLUMN()-1))", to: "B2")
```

**Excel Documentation:** [ADDRESS function](https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89)

**Implementation Status:** ✅ Full implementation

**See Also:** ``INDIRECT``, ``ROW``, ``COLUMN``

---

### COLUMN

Returns the column number of a reference.

**Syntax:** `COLUMN([reference])`

**Parameters:**
- `reference` *(optional)*: Cell or range reference (defaults to current cell)

**Returns:** Column number (1-based)

**Examples:**
```swift
// Get current column
let col = try evaluator.evaluate("=COLUMN()")  // In column C, returns 3

// Get column of specific reference
let colNum = try evaluator.evaluate("=COLUMN(D5)")  // Returns 4

// Create sequential numbers across columns
sheet.writeFormula("COLUMN(A1)", to: "A1")  // Returns 1
sheet.writeFormula("COLUMN(B1)", to: "B1")  // Returns 2

// Dynamic column offset
sheet.writeFormula("INDEX(Data, 1, COLUMN()-1)", to: "B1")
```

**Excel Documentation:** [COLUMN function](https://support.microsoft.com/en-us/office/column-function-44e8c754-711c-4df3-9da4-47a55042554b)

**Implementation Status:** ✅ Full implementation

**See Also:** ``ROW``, ``COLUMNS``, ``ADDRESS``

---

### ROW

Returns the row number of a reference.

**Syntax:** `ROW([reference])`

**Parameters:**
- `reference` *(optional)*: Cell or range reference (defaults to current cell)

**Returns:** Row number (1-based)

**Examples:**
```swift
// Get current row
let row = try evaluator.evaluate("=ROW()")  // In row 5, returns 5

// Get row of specific reference
let rowNum = try evaluator.evaluate("=ROW(A10)")  // Returns 10

// Create sequential numbers down rows
sheet.writeFormula("ROW(A1)", to: "A1")  // Returns 1
sheet.writeFormula("ROW(A2)", to: "A2")  // Returns 2

// Conditional formatting based on row
sheet.writeFormula("MOD(ROW(), 2) = 0", to: "A1")  // TRUE for even rows
```

**Excel Documentation:** [ROW function](https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d)

**Implementation Status:** ✅ Full implementation

**See Also:** ``COLUMN``, ``ROWS``, ``ADDRESS``

---

### COLUMNS

Returns the number of columns in an array or reference.

**Syntax:** `COLUMNS(array)`

**Parameters:**
- `array`: Array or range reference

**Returns:** Number of columns

**Examples:**
```swift
// Count columns in range
let count = try evaluator.evaluate("=COLUMNS(A1:D10)")  // Returns 4

// Single column
let single = try evaluator.evaluate("=COLUMNS(A:A)")  // Returns 1

// Dynamic array sizing
sheet.writeFormula("COLUMNS(DataRange)", to: "B1")
```

**Excel Documentation:** [COLUMNS function](https://support.microsoft.com/en-us/office/columns-function-4e8e7b4e-e603-43e8-b177-956088fa48ca)

**Implementation Status:** ✅ Full implementation

**See Also:** ``ROWS``, ``COLUMN``

---

### ROWS

Returns the number of rows in an array or reference.

**Syntax:** `ROWS(array)`

**Parameters:**
- `array`: Array or range reference

**Returns:** Number of rows

**Examples:**
```swift
// Count rows in range
let count = try evaluator.evaluate("=ROWS(A1:D10)")  // Returns 10

// Entire column
let many = try evaluator.evaluate("=ROWS(A:A)")  // Returns 1048576 (Excel max)

// Dynamic array sizing
sheet.writeFormula("ROWS(DataRange)", to: "B1")
```

**Excel Documentation:** [ROWS function](https://support.microsoft.com/en-us/office/rows-function-b592593e-3fc2-47f2-bec1-bda493811597)

**Implementation Status:** ✅ Full implementation

**See Also:** ``COLUMNS``, ``ROW``

---

### AREAS

Returns the number of areas in a reference.

**Syntax:** `AREAS(reference)`

**Parameters:**
- `reference`: Reference to a cell, range, or multiple ranges

**Returns:** Number of areas

**Examples:**
```swift
// Single range
let one = try evaluator.evaluate("=AREAS(A1:B5)")  // Returns 1

// Multiple ranges
let three = try evaluator.evaluate("=AREAS((A1:A5, C1:C5, E1:E5))")  // Returns 3

// Using in formulas
sheet.writeFormula("AREAS(MyRange)", to: "B1")
```

**Excel Documentation:** [AREAS function](https://support.microsoft.com/en-us/office/areas-function-8392ba32-7a41-43b3-96b0-3695d2ec6152)

**Implementation Status:** ✅ Full implementation

**See Also:** ``ROWS``, ``COLUMNS``

---

## Selection Functions

### CHOOSE

Returns a value from a list based on an index number.

**Syntax:** `CHOOSE(index_num, value1, [value2], ...)`

**Parameters:**
- `index_num`: Which value to return (1-based)
- `value1`, `value2`, ...: List of values to choose from

**Returns:** Value at the specified index, or #VALUE! if index is out of range

**Examples:**
```swift
// Select from list
let day = try evaluator.evaluate("=CHOOSE(3, \"Mon\", \"Tue\", \"Wed\", \"Thu\", \"Fri\")")  // Returns "Wed"

// Dynamic selection
let value = try evaluator.evaluate("=CHOOSE(A1, 100, 200, 300, 400)")

// Select formula result
let result = try evaluator.evaluate("=CHOOSE(2, SUM(A:A), AVERAGE(A:A), MAX(A:A))")

// Random selection
sheet.writeFormula("CHOOSE(RANDBETWEEN(1,5), \"Red\", \"Blue\", \"Green\", \"Yellow\", \"Purple\")", to: "A1")
```

**Excel Documentation:** [CHOOSE function](https://support.microsoft.com/en-us/office/choose-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SWITCH``, ``CHOOSECOLS``, ``CHOOSEROWS``

---

### CHOOSECOLS

Returns specified columns from an array (Excel 365).

**Syntax:** `CHOOSECOLS(array, col_num1, [col_num2], ...)`

**Parameters:**
- `array`: Array to select from
- `col_num1`, `col_num2`, ...: Column numbers to return (1-based)

**Returns:** Array with selected columns, or #REF! if column out of range

**Examples:**
```swift
// Select specific columns
let cols = try evaluator.evaluate("=CHOOSECOLS(A1:E10, 1, 3, 5)")  // Returns columns 1, 3, 5

// Reorder columns
let reorder = try evaluator.evaluate("=CHOOSECOLS(A1:C10, 3, 2, 1)")  // Reverse column order

// Extract single column
let single = try evaluator.evaluate("=CHOOSECOLS(Data, 2)")

// Using in formulas
sheet.writeFormula("CHOOSECOLS(SalesData, 1, 4, 6)", to: "A1")
```

**Excel Documentation:** [CHOOSECOLS function](https://support.microsoft.com/en-us/office/choosecols-function-bf117976-2722-4466-9b9a-1c01ed9aebff)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CHOOSEROWS``, ``CHOOSE``, ``TAKE``, ``DROP``

---

### CHOOSEROWS

Returns specified rows from an array (Excel 365).

**Syntax:** `CHOOSEROWS(array, row_num1, [row_num2], ...)`

**Parameters:**
- `array`: Array to select from
- `row_num1`, `row_num2`, ...: Row numbers to return (1-based)

**Returns:** Array with selected rows, or #REF! if row out of range

**Examples:**
```swift
// Select specific rows
let rows = try evaluator.evaluate("=CHOOSEROWS(A1:E10, 1, 5, 10)")  // Returns rows 1, 5, 10

// Reorder rows
let reorder = try evaluator.evaluate("=CHOOSEROWS(A1:C10, 10, 5, 1)")

// Extract single row
let single = try evaluator.evaluate("=CHOOSEROWS(Data, 3)")

// Using in formulas
sheet.writeFormula("CHOOSEROWS(Products, 2, 4, 6)", to: "A1")
```

**Excel Documentation:** [CHOOSEROWS function](https://support.microsoft.com/en-us/office/chooserows-function-51ace882-9bab-4a44-9625-7274ef7507a3)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CHOOSECOLS``, ``CHOOSE``, ``TAKE``, ``DROP``

---

### TRANSPOSE

Transposes the rows and columns of an array.

**Syntax:** `TRANSPOSE(array)`

**Parameters:**
- `array`: Array to transpose

**Returns:** Transposed array

**Examples:**
```swift
// Transpose table
let transposed = try evaluator.evaluate("=TRANSPOSE(A1:C5)")  // 3 cols × 5 rows → 5 cols × 3 rows

// Convert row to column
let column = try evaluator.evaluate("=TRANSPOSE(A1:E1)")  // 1×5 → 5×1

// Convert column to row
let row = try evaluator.evaluate("=TRANSPOSE(A1:A10)")  // 10×1 → 1×10

// Using in formulas
sheet.writeFormula("TRANSPOSE(MonthlyData)", to: "A1")
```

**Excel Documentation:** [TRANSPOSE function](https://support.microsoft.com/en-us/office/transpose-function-ed039415-ed8a-4a81-93e9-4b6dfac76027)

**Implementation Status:** ✅ Full implementation

**See Also:** ``TOCOL``, ``TOROW``

---

## Dynamic Array Functions (Excel 365)

### FILTER

Returns a filtered array based on criteria.

**Syntax:** `FILTER(array, include, [if_empty])`

**Parameters:**
- `array`: Array to filter
- `include`: Boolean array indicating which rows to include
- `if_empty` *(optional)*: Value if no rows match (default: #CALC!)

**Returns:** Filtered array

**Examples:**
```swift
// Filter by condition
let filtered = try evaluator.evaluate("=FILTER(A2:C100, B2:B100>50)")

// Multiple criteria (AND)
let multi = try evaluator.evaluate("=FILTER(A2:D100, (B2:B100>50)*(C2:C100=\"Active\"))")

// Multiple criteria (OR)
let or = try evaluator.evaluate("=FILTER(A2:D100, (B2:B100>50)+(C2:C100=\"VIP\"))")

// With custom empty message
let result = try evaluator.evaluate("=FILTER(Products, Category=\"Electronics\", \"No products found\")")

// Using in formulas
sheet.writeFormula("FILTER(Sales, Region=\"West\")", to: "A1")
```

**Excel Documentation:** [FILTER function](https://support.microsoft.com/en-us/office/filter-function-f4f7cb66-82eb-4767-8f7c-4877ad80c759)

**Implementation Status:** ✅ Full implementation

**Notes:**
- Returns all matching rows
- Include array must match array height
- Use * for AND, + for OR in criteria

**See Also:** ``SORT``, ``UNIQUE``, ``XLOOKUP``

---

### SORT

Sorts the contents of an array.

**Syntax:** `SORT(array, [sort_index], [sort_order], [by_col])`

**Parameters:**
- `array`: Array to sort
- `sort_index` *(optional)*: Column/row number to sort by (default: 1)
- `sort_order` *(optional)*: 1=ascending (default), -1=descending
- `by_col` *(optional)*: FALSE=sort by row (default), TRUE=sort by column

**Returns:** Sorted array

**Examples:**
```swift
// Sort by first column, ascending
let sorted = try evaluator.evaluate("=SORT(A2:C100)")

// Sort by second column, descending
let desc = try evaluator.evaluate("=SORT(A2:C100, 2, -1)")

// Sort entire table by price column
let prices = try evaluator.evaluate("=SORT(Products, 3, -1)")

// Using in formulas
sheet.writeFormula("SORT(Sales, 4, -1)", to: "A1")  // Sort by 4th column descending
```

**Excel Documentation:** [SORT function](https://support.microsoft.com/en-us/office/sort-function-22f63bd0-ccc8-492f-953d-c20e8e44b86c)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SORTBY``, ``FILTER``, ``UNIQUE``

---

### SORTBY

Sorts an array based on values in another array.

**Syntax:** `SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], ...)`

**Parameters:**
- `array`: Array to sort
- `by_array1`: Array to sort by
- `sort_order1` *(optional)*: 1=ascending (default), -1=descending
- Additional by_array/sort_order pairs *(optional)*: For multi-level sorting

**Returns:** Sorted array

**Examples:**
```swift
// Sort names by corresponding ages
let sorted = try evaluator.evaluate("=SORTBY(A2:A10, B2:B10)")

// Sort by multiple criteria
let multi = try evaluator.evaluate("=SORTBY(Names, Dept, 1, Salary, -1)")  // By dept asc, then salary desc

// Sort by calculated values
let calc = try evaluator.evaluate("=SORTBY(Products, Quantity*Price, -1)")  // By total value

// Using in formulas
sheet.writeFormula("SORTBY(StudentNames, TestScores, -1)", to: "A1")
```

**Excel Documentation:** [SORTBY function](https://support.microsoft.com/en-us/office/sortby-function-cd2d7a62-1b93-435c-b561-d6a35134f28f)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SORT``, ``FILTER``

---

### UNIQUE

Returns unique values from a list or range.

**Syntax:** `UNIQUE(array, [by_col], [exactly_once])`

**Parameters:**
- `array`: Array to extract unique values from
- `by_col` *(optional)*: FALSE=compare rows (default), TRUE=compare columns
- `exactly_once` *(optional)*: FALSE=all unique values (default), TRUE=only values appearing once

**Returns:** Array of unique values

**Examples:**
```swift
// Get unique values
let unique = try evaluator.evaluate("=UNIQUE(A2:A100)")

// Unique rows (multiple columns)
let rows = try evaluator.evaluate("=UNIQUE(A2:C100)")

// Only values appearing exactly once
let distinct = try evaluator.evaluate("=UNIQUE(A2:A100, FALSE, TRUE)")

// Using in formulas
sheet.writeFormula("UNIQUE(CustomerNames)", to: "A1")
```

**Excel Documentation:** [UNIQUE function](https://support.microsoft.com/en-us/office/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e)

**Implementation Status:** ✅ Full implementation

**See Also:** ``FILTER``, ``SORT``

---

### SEQUENCE

Generates an array of sequential numbers.

**Syntax:** `SEQUENCE(rows, [columns], [start], [step])`

**Parameters:**
- `rows`: Number of rows
- `columns` *(optional)*: Number of columns (default: 1)
- `start` *(optional)*: Starting number (default: 1)
- `step` *(optional)*: Increment (default: 1)

**Returns:** Array of sequential numbers

**Examples:**
```swift
// Simple sequence 1-10
let seq = try evaluator.evaluate("=SEQUENCE(10)")

// 2D sequence
let grid = try evaluator.evaluate("=SEQUENCE(5, 3)")  // 5×3 grid: 1-15

// Custom start and step
let custom = try evaluator.evaluate("=SEQUENCE(10, 1, 100, 10)")  // 100, 110, 120, ..., 190

// Dates sequence
sheet.writeFormula("TODAY() + SEQUENCE(7) - 1", to: "A1")  // Next 7 days

// Using in formulas
sheet.writeFormula("SEQUENCE(ROW(A1:A100))", to: "B1")
```

**Excel Documentation:** [SEQUENCE function](https://support.microsoft.com/en-us/office/sequence-function-57467a98-57e0-4817-9f14-2eb78519ca90)

**Implementation Status:** ✅ Full implementation

**See Also:** ``RANDARRAY``

---

### RANDARRAY

Generates an array of random numbers.

**Syntax:** `RANDARRAY([rows], [columns], [min], [max], [integer])`

**Parameters:**
- `rows` *(optional)*: Number of rows (default: 1)
- `columns` *(optional)*: Number of columns (default: 1)
- `min` *(optional)*: Minimum value (default: 0)
- `max` *(optional)*: Maximum value (default: 1)
- `integer` *(optional)*: TRUE for integers, FALSE for decimals (default: FALSE)

**Returns:** Array of random numbers

**Examples:**
```swift
// 5 random decimals between 0 and 1
let rand = try evaluator.evaluate("=RANDARRAY(5)")

// 3×4 grid of random integers 1-100
let grid = try evaluator.evaluate("=RANDARRAY(3, 4, 1, 100, TRUE)")

// Random sample data
let sample = try evaluator.evaluate("=RANDARRAY(10, 1, -50, 50)")

// Using in formulas
sheet.writeFormula("RANDARRAY(100, 1, 1, 1000, TRUE)", to: "A1")
```

**Excel Documentation:** [RANDARRAY function](https://support.microsoft.com/en-us/office/randarray-function-21261e55-3bec-4885-86a6-8b0a47fd4d33)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SEQUENCE``, ``RAND``, ``RANDBETWEEN``

---

### TAKE

Returns specified number of rows or columns from array start or end.

**Syntax:** `TAKE(array, rows, [columns])`

**Parameters:**
- `array`: Array to take from
- `rows`: Number of rows (positive=from start, negative=from end)
- `columns` *(optional)*: Number of columns (positive=from start, negative=from end)

**Returns:** Subset of array

**Examples:**
```swift
// First 5 rows
let top5 = try evaluator.evaluate("=TAKE(A1:C100, 5)")

// Last 5 rows
let bottom5 = try evaluator.evaluate("=TAKE(A1:C100, -5)")

// First 3 rows, last 2 columns
let subset = try evaluator.evaluate("=TAKE(A1:E100, 3, -2)")

// Using with SORT for top N
sheet.writeFormula("TAKE(SORT(Sales, 2, -1), 10)", to: "A1")  // Top 10
```

**Excel Documentation:** [TAKE function](https://support.microsoft.com/en-us/office/take-function-25382ff1-5da1-4f78-ab43-f33bd2e4e003)

**Implementation Status:** ✅ Full implementation

**See Also:** ``DROP``, ``CHOOSECOLS``, ``CHOOSEROWS``

---

### DROP

Returns array with specified number of rows or columns removed.

**Syntax:** `DROP(array, rows, [columns])`

**Parameters:**
- `array`: Array to drop from
- `rows`: Number of rows to drop (positive=from start, negative=from end)
- `columns` *(optional)*: Number of columns to drop (positive=from start, negative=from end)

**Returns:** Array with rows/columns removed

**Examples:**
```swift
// Remove header row
let noHeader = try evaluator.evaluate("=DROP(A1:C100, 1)")

// Remove last row
let noTotal = try evaluator.evaluate("=DROP(A1:C100, -1)")

// Remove first column
let noID = try evaluator.evaluate("=DROP(A1:E100, 0, 1)")

// Using in formulas
sheet.writeFormula("DROP(Data, 1)", to: "A1")  // Skip header
```

**Excel Documentation:** [DROP function](https://support.microsoft.com/en-us/office/drop-function-1cb4e151-9e17-4838-abe5-9ba48d8c6a34)

**Implementation Status:** ✅ Full implementation

**See Also:** ``TAKE``, ``CHOOSECOLS``, ``CHOOSEROWS``

---

### EXPAND

Expands an array to specified dimensions with fill value.

**Syntax:** `EXPAND(array, rows, [columns], [pad_with])`

**Parameters:**
- `array`: Array to expand
- `rows`: Target number of rows (0=don't expand rows)
- `columns` *(optional)*: Target number of columns (0=don't expand columns)
- `pad_with` *(optional)*: Value to fill empty cells (default: #N/A)

**Returns:** Expanded array

**Examples:**
```swift
// Expand to 10 rows
let expanded = try evaluator.evaluate("=EXPAND(A1:A5, 10)")

// Expand to 5×3 with zeros
let grid = try evaluator.evaluate("=EXPAND(A1:B2, 5, 3, 0)")

// Pad with empty string
let padded = try evaluator.evaluate("=EXPAND(Names, 100, 1, \"\")")

// Using in formulas
sheet.writeFormula("EXPAND(Data, 100, 10, 0)", to: "A1")
```

**Excel Documentation:** [EXPAND function](https://support.microsoft.com/en-us/office/expand-function-7433fba5-4ad1-41da-a904-d5d95808bc38)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SEQUENCE``

---

## Array Stacking Functions

### VSTACK

Stacks arrays vertically (row-wise).

**Syntax:** `VSTACK(array1, [array2], ...)`

**Parameters:**
- `array1`, `array2`, ...: Arrays to stack

**Returns:** Vertically combined array, or #VALUE! if column counts differ

**Examples:**
```swift
// Stack two tables
let stacked = try evaluator.evaluate("=VSTACK(A1:C10, A20:C30)")

// Combine multiple arrays
let combined = try evaluator.evaluate("=VSTACK(Headers, Data, Totals)")

// Stack with single row
let withTotal = try evaluator.evaluate("=VSTACK(Data, {\"Total\", SUM(B:B), SUM(C:C)})")

// Using in formulas
sheet.writeFormula("VSTACK(NorthSales, SouthSales, EastSales, WestSales)", to: "A1")
```

**Excel Documentation:** [VSTACK function](https://support.microsoft.com/en-us/office/vstack-function-a4b86897-be0f-48fc-adca-fcc10d795a9c)

**Implementation Status:** ✅ Full implementation

**See Also:** ``HSTACK``, ``TOCOL``

---

### HSTACK

Stacks arrays horizontally (column-wise).

**Syntax:** `HSTACK(array1, [array2], ...)`

**Parameters:**
- `array1`, `array2`, ...: Arrays to stack

**Returns:** Horizontally combined array, or #VALUE! if row counts differ

**Examples:**
```swift
// Stack two tables side-by-side
let stacked = try evaluator.evaluate("=HSTACK(A1:B10, D1:E10)")

// Combine multiple columns
let combined = try evaluator.evaluate("=HSTACK(IDs, Names, Emails, Phones)")

// Add calculated column
let withCalc = try evaluator.evaluate("=HSTACK(Data, B2:B100*C2:C100)")

// Using in formulas
sheet.writeFormula("HSTACK(Products, Prices, Quantities)", to: "A1")
```

**Excel Documentation:** [HSTACK function](https://support.microsoft.com/en-us/office/hstack-function-98c4ab76-10fe-4b4f-8d5f-af1c125fe8c2)

**Implementation Status:** ✅ Full implementation

**See Also:** ``VSTACK``, ``TOROW``

---

### TOCOL

Transforms an array into a single column.

**Syntax:** `TOCOL(array, [ignore], [scan_by_column])`

**Parameters:**
- `array`: Array to transform
- `ignore` *(optional)*: 0=keep all (default), 1=ignore blanks, 2=ignore errors, 3=ignore blanks and errors
- `scan_by_column` *(optional)*: FALSE=by row (default), TRUE=by column

**Returns:** Single-column array

**Examples:**
```swift
// Flatten to column
let column = try evaluator.evaluate("=TOCOL(A1:E10)")

// Ignore blanks
let noBlanks = try evaluator.evaluate("=TOCOL(A1:E10, 1)")

// Scan by column instead of row
let byCol = try evaluator.evaluate("=TOCOL(A1:E10, 0, TRUE)")

// Using in formulas
sheet.writeFormula("TOCOL(DataRange, 1)", to: "A1")
```

**Excel Documentation:** [TOCOL function](https://support.microsoft.com/en-us/office/tocol-function-22839d9b-0b55-4fc1-b4e6-2614c7e7e14e)

**Implementation Status:** ✅ Full implementation

**See Also:** ``TOROW``, ``VSTACK``, ``TRANSPOSE``

---

### TOROW

Transforms an array into a single row.

**Syntax:** `TOROW(array, [ignore], [scan_by_column])`

**Parameters:**
- `array`: Array to transform
- `ignore` *(optional)*: 0=keep all (default), 1=ignore blanks, 2=ignore errors, 3=ignore blanks and errors
- `scan_by_column` *(optional)*: FALSE=by row (default), TRUE=by column

**Returns:** Single-row array

**Examples:**
```swift
// Flatten to row
let row = try evaluator.evaluate("=TOROW(A1:E10)")

// Ignore blanks
let noBlanks = try evaluator.evaluate("=TOROW(A1:E10, 1)")

// Scan by column
let byCol = try evaluator.evaluate("=TOROW(A1:E10, 0, TRUE)")

// Using in formulas
sheet.writeFormula("TOROW(DataRange, 3)", to: "A1")  // Ignore blanks and errors
```

**Excel Documentation:** [TOROW function](https://support.microsoft.com/en-us/office/torow-function-b90d0964-a7d9-44b7-816b-ffa5c2fe2289)

**Implementation Status:** ✅ Full implementation

**See Also:** ``TOCOL``, ``HSTACK``, ``TRANSPOSE``

---

### WRAPCOLS

Wraps a single row or column into multiple columns.

**Syntax:** `WRAPCOLS(vector, wrap_count, [pad_with])`

**Parameters:**
- `vector`: Single row or column array
- `wrap_count`: Number of values per column
- `pad_with` *(optional)*: Fill value for incomplete columns (default: #N/A)

**Returns:** Array wrapped into columns

**Examples:**
```swift
// Wrap into 3 rows per column
let wrapped = try evaluator.evaluate("=WRAPCOLS(A1:A20, 3)")

// Wrap with zero padding
let padded = try evaluator.evaluate("=WRAPCOLS(A1:A25, 10, 0)")

// Using in formulas
sheet.writeFormula("WRAPCOLS(SEQUENCE(100), 10)", to: "A1")  // 10×10 grid
```

**Excel Documentation:** [WRAPCOLS function](https://support.microsoft.com/en-us/office/wrapcols-function-d038b05a-57b7-4ee0-be94-ded0792511e2)

**Implementation Status:** ✅ Full implementation

**See Also:** ``WRAPROWS``, ``TOCOL``

---

### WRAPROWS

Wraps a single row or column into multiple rows.

**Syntax:** `WRAPROWS(vector, wrap_count, [pad_with])`

**Parameters:**
- `vector`: Single row or column array
- `wrap_count`: Number of values per row
- `pad_with` *(optional)*: Fill value for incomplete rows (default: #N/A)

**Returns:** Array wrapped into rows

**Examples:**
```swift
// Wrap into 5 values per row
let wrapped = try evaluator.evaluate("=WRAPROWS(A1:A100, 5)")

// Wrap with blank padding
let padded = try evaluator.evaluate("=WRAPROWS(A1:A23, 10, \"\")")

// Using in formulas
sheet.writeFormula("WRAPROWS(SEQUENCE(100), 10)", to: "A1")  // 10×10 grid
```

**Excel Documentation:** [WRAPROWS function](https://support.microsoft.com/en-us/office/wraprows-function-796825f3-975a-4cee-9c84-1bbddf60ade0)

**Implementation Status:** ✅ Full implementation

**See Also:** ``WRAPCOLS``, ``TOROW``

---

## Common Use Cases

### Basic Lookup Operations

Simple product price lookup:
```swift
// Exact match lookup
sheet.writeFormula("VLOOKUP(A2, Products!A:D, 3, FALSE)", to: "B2")

// With error handling
sheet.writeFormula("IFERROR(VLOOKUP(A2, Products!A:D, 3, FALSE), \"Not Found\")", to: "B2")
```

### Two-Way Lookup

Find value at intersection of row and column:
```swift
// Using INDEX/MATCH
sheet.writeFormula("INDEX(Data, MATCH(A1, RowHeaders, 0), MATCH(B1, ColHeaders, 0))", to: "C1")

// Using XLOOKUP (Excel 365)
sheet.writeFormula("XLOOKUP(A1, RowHeaders, XLOOKUP(B1, ColHeaders, Data))", to: "C1")
```

### Dynamic Range References

Create flexible references:
```swift
// Last N rows
sheet.writeFormula("TAKE(Data, -10)", to: "A1")  // Last 10 rows

// Skip header and total rows
sheet.writeFormula("DROP(DROP(Data, 1), -1)", to: "A1")

// Offset from current position
sheet.writeFormula("OFFSET(A1, 1, 0, 5, 1)", to: "B1")
```

### Data Filtering and Analysis

Filter and analyze data:
```swift
// Filter by multiple criteria
sheet.writeFormula("FILTER(Sales, (Region=\"West\")*(Amount>1000))", to: "A1")

// Top 10 by value
sheet.writeFormula("TAKE(SORT(Data, 3, -1), 10)", to: "A1")

// Unique values
sheet.writeFormula("SORT(UNIQUE(Products))", to: "A1")
```

### Array Manipulation

Transform and combine arrays:
```swift
// Combine multiple regions
sheet.writeFormula("VSTACK(North, South, East, West)", to: "A1")

// Add calculated column
sheet.writeFormula("HSTACK(Products, Qty*Price)", to: "A1")

// Transpose table
sheet.writeFormula("TRANSPOSE(MonthlyData)", to: "A1")
```

### Sequential Data Generation

Create sequences and patterns:
```swift
// Row numbers
sheet.writeFormula("SEQUENCE(ROWS(Data))", to: "A1")

// Date range
sheet.writeFormula("TODAY() + SEQUENCE(30) - 1", to: "A1")  // Next 30 days

// Custom sequence
sheet.writeFormula("SEQUENCE(12, 1, 100, 25)", to: "A1")  // 100, 125, 150, ..., 375
```

### Reshaping Data

Convert between row and column layouts:
```swift
// Column to row
sheet.writeFormula("TOROW(A1:A100)", to: "B1")

// Wrap into grid
sheet.writeFormula("WRAPROWS(A1:A100, 10)", to: "B1")  // 10 columns

// Flatten multi-column to single column
sheet.writeFormula("TOCOL(A1:E20, 1)", to: "G1")  // Ignore blanks
```

## Performance Considerations

### Lookup Function Performance

- **XLOOKUP** vs **VLOOKUP**: XLOOKUP is generally faster for large datasets
- **INDEX/MATCH** vs **VLOOKUP**: INDEX/MATCH is more flexible and can be faster
- **Binary search**: Use MATCH with sorted data and match_type 1 or -1 for O(log n) performance
- **Approximate match**: Requires sorted data but is faster than exact match

### Dynamic Array Memory

- **FILTER**: Memory scales with result size, not input size
- **SORT**: O(n log n) time complexity, temporary memory for sorting
- **UNIQUE**: Memory proportional to unique values
- **Large arrays**: Consider breaking into smaller chunks for very large datasets

### Volatile Functions

These functions recalculate with every change:
- ``INDIRECT`` - Always volatile
- ``OFFSET`` - Always volatile
- ``TODAY`` - Recalculates on date change
- ``RAND``, ``RANDARRAY`` - Recalculate continuously

Use volatile functions sparingly in large workbooks.

### Best Practices

1. **Use exact match by default**: Set range_lookup to FALSE unless you need approximate matching
2. **Prefer XLOOKUP over VLOOKUP**: More features, better error handling, easier syntax
3. **Use INDEX/MATCH for flexibility**: Can look left, return multiple columns, dynamic references
4. **Chain dynamic array functions**: `TAKE(SORT(FILTER(...)))` for complex operations
5. **Handle errors**: Use IFERROR or IFNA to prevent error propagation
6. **Sort before lookup**: Approximate match lookups require sorted data
7. **Minimize volatile functions**: INDIRECT and OFFSET slow down large workbooks

## See Also

- <doc:FormulaReference> - Complete function reference
- <doc:Logical> - IF, IFERROR for error handling in lookups
- <doc:Mathematical> - Aggregate functions to use with filtered data
- <doc:Text> - Text functions for data cleaning before lookup
- ``FormulaEvaluator`` - Evaluate lookup formulas programmatically
