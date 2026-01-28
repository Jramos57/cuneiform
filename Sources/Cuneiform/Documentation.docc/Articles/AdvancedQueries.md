# Advanced Queries

Master complex data queries, filtering, and search operations in Cuneiform.

## Overview

Cuneiform's ``Sheet`` provides powerful querying capabilities that go beyond simple cell access. You can search for specific values, filter rows, iterate over columns and ranges, and build complex data extraction pipelines—all with an intuitive, Swift-native API.

## Cell Access

### Basic Cell Access

```swift
let sheet = try workbook.sheet(at: 0)!

// By string reference
let value = sheet.cell(at: "A1")

// By CellReference
let ref = CellReference(column: "A", row: 1)
let value = sheet.cell(at: ref)

// Access formulas
if let formula = sheet.formula(at: "B1") {
    print("Formula: \(formula.expression)")
    print("Cached value: \(formula.cachedValue ?? "none")")
}
```

### Working with Cell Values

```swift
if let value = sheet.cell(at: "A1") {
    switch value {
    case .text(let str):
        print("Text: \(str)")
    case .number(let num):
        print("Number: \(num)")
    case .date(let dateStr):
        print("Date: \(dateStr)")
    case .boolean(let bool):
        print("Boolean: \(bool)")
    case .error(let err):
        print("Error: \(err)")
    case .empty:
        print("Empty cell")
    case .richText(let runs):
        print("Rich text with \(runs.count) runs")
    }
}
```

## Row Operations

### Accessing Rows

```swift
// Get specific row (1-based indexing)
let row1 = sheet.row(1)
for value in row1 {
    print(value)
}

// Streaming iteration (memory-efficient)
for row in sheet.rows() {
    for (ref, value) in row {
        print("\(ref): \(value)")
    }
}
```

### Filtering Rows

Find rows that match specific criteria:

```swift
// Find rows containing "Active"
let activeRows = sheet.rows { cells in
    cells.contains { $0.value == .text("Active") }
}

// Find rows where column A is a number > 100
let highValueRows = sheet.rows { cells in
    guard let firstCell = cells.first else { return false }
    if case .number(let n) = firstCell.value {
        return n > 100
    }
    return false
}

// Complex filtering with multiple conditions
let filteredRows = sheet.rows { cells in
    guard cells.count >= 3 else { return false }
    
    // Check column A: must be "Active"
    guard case .text(let status) = cells[0].value, status == "Active" else {
        return false
    }
    
    // Check column C: must be number > 50
    guard case .number(let score) = cells[2].value, score > 50 else {
        return false
    }
    
    return true
}
```

## Column Operations

### Accessing Columns

```swift
// By column letter
let columnA = sheet.column("A")
for (row, value) in columnA {
    print("A\(row): \(value ?? .empty)")
}

// By column index (0-based: 0=A, 1=B, etc.)
let columnB = sheet.column(at: 1)
for (row, value) in columnB {
    print("B\(row): \(value ?? .empty)")
}
```

### Column Statistics

```swift
// Sum a column
let columnB = sheet.column("B")
let sum = columnB.reduce(0.0) { result, item in
    if case .number(let n) = item.value {
        return result + n
    }
    return result
}

// Count non-empty cells
let count = columnB.filter { $0.value != nil && $0.value != .empty }.count

// Find max value
let maxValue = columnB.compactMap { item -> Double? in
    if case .number(let n) = item.value {
        return n
    }
    return nil
}.max()

// Average
let numbers = columnB.compactMap { item -> Double? in
    if case .number(let n) = item.value {
        return n
    }
    return nil
}
let average = numbers.reduce(0, +) / Double(numbers.count)
```

### Column Transformations

```swift
// Extract all text values from column A
let names = sheet.column("A").compactMap { item -> String? in
    if case .text(let str) = item.value {
        return str
    }
    return nil
}

// Create lookup map: column A -> column B
var lookup: [String: Double] = [:]
let colA = sheet.column("A")
let colB = sheet.column("B")

for (index, (_, valueA)) in colA.enumerated() {
    if case .text(let key) = valueA,
       let (_, valueB) = colB.first(where: { $0.row == colA[index].row }),
       case .number(let num) = valueB {
        lookup[key] = num
    }
}
```

## Range Operations

### Accessing Ranges

```swift
// Get rectangular range
let range = sheet.range("A1:C10")

for (ref, value) in range {
    print("\(ref): \(value ?? .empty)")
}

// Process range as 2D grid
let rangeData = sheet.range("B2:D4")
var grid: [[CellValue]] = []
var currentRow: [CellValue] = []
var lastRow = 0

for (ref, value) in rangeData {
    if ref.row != lastRow && !currentRow.isEmpty {
        grid.append(currentRow)
        currentRow = []
    }
    currentRow.append(value ?? .empty)
    lastRow = ref.row
}
if !currentRow.isEmpty {
    grid.append(currentRow)
}
```

### Range Analysis

```swift
// Sum a range
let range = sheet.range("B2:B10")
let total = range.reduce(0.0) { sum, item in
    if case .number(let n) = item.value {
        return sum + n
    }
    return sum
}

// Count cells with specific value
let range = sheet.range("C1:C100")
let activeCount = range.filter { $0.value == .text("Active") }.count

// Find min/max in range
let numbers = sheet.range("D1:D50").compactMap { item -> Double? in
    if case .number(let n) = item.value {
        return n
    }
    return nil
}
let min = numbers.min()
let max = numbers.max()
```

## Search Operations

### Find First Match

```swift
// Find first cell with specific value
if let result = sheet.find(where: { _, value in
    value == .text("Alice")
}) {
    print("Found at \(result.reference): \(result.value)")
}

// Find first cell matching condition
if let result = sheet.find(where: { _, value in
    if case .number(let n) = value {
        return n > 100
    }
    return false
}) {
    print("First value > 100 at \(result.reference)")
}

// Find by pattern
if let result = sheet.find(where: { _, value in
    if case .text(let str) = value {
        return str.hasPrefix("Error:")
    }
    return false
}) {
    print("Found error message at \(result.reference)")
}
```

### Find All Matches

```swift
// Find all cells with specific value
let results = sheet.findAll { _, value in
    value == .text("TODO")
}
for match in results {
    print("TODO at \(match.reference)")
}

// Find all high scores
let highScores = sheet.findAll { _, value in
    if case .number(let n) = value {
        return n >= 90
    }
    return false
}

// Find all cells containing a substring
let matches = sheet.findAll { _, value in
    if case .text(let str) = value {
        return str.contains("important")
    }
    return false
}
```

### Search Strategies

```swift
// ✅ Good: Use find() for first match (stops early)
if let first = sheet.find(where: { _, value in value == target }) {
    process(first)
}

// ❌ Less efficient: findAll() then take first
let all = sheet.findAll { _, value in value == target }
if let first = all.first {
    process(first)
}

// ✅ Good: Combine with range for targeted search
let searchRange = sheet.range("A1:A1000")
let result = searchRange.first { $0.value == target }

// ✅ Good: Search within column
let matches = sheet.column("B").filter { row, value in
    guard case .number(let n) = value else { return false }
    return n > 50
}
```

## Data Validations

Access and query data validation rules:

```swift
// Get validations for a range
let validations = sheet.validations(for: "B2:B10")
for validation in validations {
    print("Type: \(validation.validationType)")
    print("Range: \(validation.sqref)")
    if let formula = validation.formula1 {
        print("Rule: \(formula)")
    }
}

// Get validations for specific cell
let cellValidations = sheet.validations(at: CellReference(column: "B", row: 5))
for validation in cellValidations {
    print("Cell has \(validation.validationType) validation")
}

// Check if dropdown list
let validations = sheet.validations(at: "C2")
let hasDropdown = validations.contains { $0.validationType == "list" }
```

## Hyperlinks

Access hyperlink information:

```swift
// Get hyperlinks at specific cell
let links = sheet.hyperlinks(at: "A1")
for link in links {
    if let url = link.url {
        print("External link: \(url)")
    }
    if let location = link.location {
        print("Internal link: \(location)")
    }
    print("Display: \(link.display ?? "none")")
    print("Tooltip: \(link.tooltip ?? "none")")
}

// Find all external links
let allLinks = sheet.hyperlinks.filter { $0.url != nil }
for link in allLinks {
    print("\(link.ref): \(link.url!)")
}
```

## Comments

Access cell comments (notes):

```swift
// Get comments at specific cell
let comments = sheet.comments(at: "B5")
for comment in comments {
    print("Author: \(comment.author ?? "Unknown")")
    print("Text: \(comment.text)")
}

// Find all comments
for comment in sheet.comments {
    print("\(comment.ref): \(comment.text)")
}

// Search comments by content
let importantComments = sheet.comments.filter { comment in
    comment.text.contains("IMPORTANT")
}
```

## Complex Queries

### Multi-Condition Filtering

```swift
struct Record {
    let name: String
    let score: Double
    let status: String
}

// Extract records matching criteria
var records: [Record] = []

let rows = sheet.rows { cells in
    guard cells.count >= 3 else { return false }
    
    // Must have valid name (text)
    guard case .text = cells[0].value else { return false }
    
    // Must have valid score (number)
    guard case .number = cells[1].value else { return false }
    
    // Must have status
    guard case .text = cells[2].value else { return false }
    
    return true
}

for row in rows {
    if case .text(let name) = row[0].value,
       case .number(let score) = row[1].value,
       case .text(let status) = row[2].value {
        records.append(Record(name: name, score: score, status: status))
    }
}

// Filter records
let highScorers = records.filter { $0.score > 90 }
let activeHighScorers = records.filter { $0.status == "Active" && $0.score > 90 }
```

### Aggregation Queries

```swift
// Group by and aggregate
var salesByRegion: [String: Double] = [:]

for row in sheet.rows() {
    guard let regionCell = row.first(where: { $0.reference.column == "A" }),
          let salesCell = row.first(where: { $0.reference.column == "B" }) else {
        continue
    }
    
    if case .text(let region) = regionCell.value,
       case .number(let sales) = salesCell.value {
        salesByRegion[region, default: 0] += sales
    }
}

// Top N query
let topRegions = salesByRegion.sorted { $0.value > $1.value }.prefix(5)
for (region, sales) in topRegions {
    print("\(region): $\(sales)")
}
```

### Join Operations

```swift
// Join data from two sheets
let sheet1 = try workbook.sheet(named: "Employees")!
let sheet2 = try workbook.sheet(named: "Salaries")!

// Build lookup from sheet2 (ID -> Salary)
var salaryLookup: [String: Double] = [:]
for row in sheet2.rows() {
    guard let idCell = row.first(where: { $0.reference.column == "A" }),
          let salaryCell = row.first(where: { $0.reference.column == "B" }) else {
        continue
    }
    
    if case .text(let id) = idCell.value,
       case .number(let salary) = salaryCell.value {
        salaryLookup[id] = salary
    }
}

// Join with sheet1
var employeesWithSalaries: [(name: String, salary: Double)] = []
for row in sheet1.rows() {
    guard let idCell = row.first(where: { $0.reference.column == "A" }),
          let nameCell = row.first(where: { $0.reference.column == "B" }) else {
        continue
    }
    
    if case .text(let id) = idCell.value,
       case .text(let name) = nameCell.value,
       let salary = salaryLookup[id] {
        employeesWithSalaries.append((name, salary))
    }
}
```

### Pivot-Like Queries

```swift
// Summarize data by multiple dimensions
struct Sale {
    let region: String
    let product: String
    let amount: Double
}

// Extract sales
var sales: [Sale] = []
for row in sheet.rows() {
    // ... extract region, product, amount from row
    // sales.append(Sale(...))
}

// Pivot: region x product -> sum
var pivot: [String: [String: Double]] = [:]
for sale in sales {
    pivot[sale.region, default: [:]][sale.product, default: 0] += sale.amount
}

// Query pivot table
for (region, products) in pivot.sorted(by: { $0.key < $1.key }) {
    print("Region: \(region)")
    for (product, total) in products.sorted(by: { $0.key < $1.key }) {
        print("  \(product): $\(total)")
    }
}
```

## Performance Optimization

### Choose the Right Method

```swift
// ✅ Fast: Single value lookup
let value = sheet.cell(at: "A1")

// ✅ Fast: Column access (vertical scan)
let column = sheet.column("B")

// ✅ Fast: Range with known bounds
let range = sheet.range("A1:C10")

// ⚠️ Slower: Full sheet scan
let all = sheet.findAll { _, _ in true }

// ✅ Fast: Stop at first match
let first = sheet.find(where: predicate)

// ⚠️ Slower: Find all then take first
let all = sheet.findAll(where: predicate)
let first = all.first
```

### Stream for Large Sheets

```swift
// ✅ Memory efficient: Stream rows
var count = 0
for row in sheet.rows() {
    if row.contains(where: { $0.value == target }) {
        count += 1
    }
}

// ❌ Memory intensive: Load all rows
let allRows = (1...10000).map { sheet.row($0) }
let count = allRows.filter { row in
    row.contains(target)
}.count
```

### Filter Early

```swift
// ✅ Good: Filter during iteration
let activeRows = sheet.rows { cells in
    cells.contains { $0.value == .text("Active") }
}

// ❌ Bad: Load everything then filter
var allRows: [[CellValue]] = []
for i in 1...sheet.rowCount {
    allRows.append(sheet.row(i))
}
let activeRows = allRows.filter { /* ... */ }
```

## Real-World Examples

### Data Validation Report

```swift
func generateValidationReport(sheet: Sheet) {
    var report: [(cell: String, issues: [String])] = []
    
    // Find all cells with validation errors
    for validation in sheet.dataValidations {
        // Parse range and check each cell
        for (ref, value) in sheet.range(validation.sqref) {
            var issues: [String] = []
            
            // Check if value matches validation
            if validation.validationType == "list" {
                if case .text(let str) = value,
                   let formula = validation.formula1,
                   !formula.contains(str) {
                    issues.append("Not in allowed list")
                }
            }
            
            if !issues.isEmpty {
                report.append((cell: "\(ref)", issues: issues))
            }
        }
    }
    
    // Print report
    for (cell, issues) in report {
        print("\(cell): \(issues.joined(separator: ", "))")
    }
}
```

### Duplicate Detection

```swift
func findDuplicates(in column: String, sheet: Sheet) -> [String: [Int]] {
    var valueCounts: [String: [Int]] = [:]
    
    for (row, value) in sheet.column(column) {
        if case .text(let str) = value {
            valueCounts[str, default: []].append(row)
        }
    }
    
    // Return only values with duplicates
    return valueCounts.filter { $0.value.count > 1 }
}

// Usage
let duplicates = findDuplicates(in: "A", sheet: sheet)
for (value, rows) in duplicates {
    print("\"\(value)\" appears in rows: \(rows.map(String.init).joined(separator: ", "))")
}
```

### Missing Data Analysis

```swift
func analyzeMissingData(sheet: Sheet) {
    var columnMissingCounts: [String: Int] = [:]
    
    // Check each column
    for col in 0..<10 {  // Check first 10 columns
        let letter = String(UnicodeScalar(65 + col)!)
        let column = sheet.column(letter)
        
        let missingCount = column.filter { row, value in
            value == nil || value == .empty
        }.count
        
        if missingCount > 0 {
            columnMissingCounts[letter] = missingCount
        }
    }
    
    // Report
    for (col, count) in columnMissingCounts.sorted(by: { $0.key < $1.key }) {
        print("Column \(col): \(count) missing values")
    }
}
```

## See Also

- <doc:GettingStarted> - Basic querying examples
- <doc:PerformanceTuning> - Query optimization strategies
- ``Sheet`` - Complete query API reference
- ``CellValue`` - Cell value types
- ``CellReference`` - Cell addressing
