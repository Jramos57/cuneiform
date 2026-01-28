# Writing Workbooks

Comprehensive guide to creating and modifying Excel files with Cuneiform.

## Overview

Cuneiform's ``WorkbookWriter`` provides a high-level API for creating .xlsx files. You can write different cell types, formulas, apply protection, add hyperlinks, merge cells, and configure advanced features—all with a clean, type-safe Swift API.

## Basic Writing

### Creating Your First Workbook

```swift
import Cuneiform

// Create a new workbook
var writer = WorkbookWriter()

// Add a sheet
let sheetIndex = writer.addSheet(named: "Data")

// Write some data
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Hello", to: "A1")
    sheet.writeNumber(42, to: "B1")
    sheet.writeBoolean(true, to: "C1")
}

// Save to disk
try writer.save(to: URL(fileURLWithPath: "output.xlsx"))
```

### Writing Different Cell Types

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Text
    sheet.writeText("Product Name", to: "A1")
    
    // Numbers
    sheet.writeNumber(99.99, to: "B1")
    sheet.writeNumber(42, to: "B2")
    
    // Booleans
    sheet.writeBoolean(true, to: "C1")
    sheet.writeBoolean(false, to: "C2")
    
    // Formulas with cached values
    sheet.writeFormula("B1+B2", cachedValue: 141.99, to: "B3")
    
    // Formulas without cached values (Excel will calculate on open)
    sheet.writeFormula("SUM(B1:B2)", to: "B4")
}
```

### Using Cell References

Two ways to specify cell locations:

```swift
// String reference
sheet.writeText("Hello", to: "A1")

// CellReference struct (more type-safe)
let ref = CellReference(column: "A", row: 2)
sheet.writeText("World", to: ref)

// Computed references
for i in 1...10 {
    let ref = CellReference(column: "A", row: i)
    sheet.writeNumber(Double(i), to: ref)
}
```

## Working with Formulas

### Writing Formulas

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Simple formula
    sheet.writeFormula("A1+B1", to: "C1")
    
    // Function call
    sheet.writeFormula("SUM(A1:A10)", to: "A11")
    
    // Complex formula
    sheet.writeFormula("IF(B1>100, \"High\", \"Low\")", to: "C1")
    
    // Cross-sheet reference
    sheet.writeFormula("Sheet2!A1 * 2", to: "D1")
}
```

### Cached Values

Providing cached values makes Excel open faster:

```swift
// Pre-compute the result
let values = [10.0, 20.0, 30.0]
let sum = values.reduce(0, +)  // 60.0

// Write formula with cached value
sheet.writeFormula("SUM(A1:A3)", cachedValue: sum, to: "A4")

// Excel will display 60.0 immediately and recalculate in background
```

### Formula Best Practices

```swift
// ✅ Good: Provide cached values
let result = computeResult()
sheet.writeFormula("COMPLEX_CALCULATION(A1:Z100)", cachedValue: result, to: "AA1")

// ⚠️ Works but slower: Excel must calculate on open
sheet.writeFormula("COMPLEX_CALCULATION(A1:Z100)", to: "AA1")

// ✅ Good: Use absolute references when needed
sheet.writeFormula("$A$1+B1", to: "C1")  // A1 is absolute, B1 is relative
```

## Multiple Sheets

### Adding Multiple Sheets

```swift
var writer = WorkbookWriter()

let dataSheet = writer.addSheet(named: "Data")
let summarySheet = writer.addSheet(named: "Summary")
let chartSheet = writer.addSheet(named: "Charts")

// Write to each sheet
writer.modifySheet(at: dataSheet) { sheet in
    sheet.writeText("Raw Data", to: "A1")
}

writer.modifySheet(at: summarySheet) { sheet in
    sheet.writeText("Summary", to: "A1")
    // Reference data from first sheet
    sheet.writeFormula("SUM(Data!A2:A100)", to: "B2")
}

try writer.save(to: outputURL)
```

### Sheet Organization

```swift
// Create sheets in logical order
let sheets = [
    ("Overview", writer.addSheet(named: "Overview")),
    ("Q1 Data", writer.addSheet(named: "Q1 Data")),
    ("Q2 Data", writer.addSheet(named: "Q2 Data")),
    ("Annual Summary", writer.addSheet(named: "Annual Summary"))
]

// Populate sheets
for (name, index) in sheets {
    writer.modifySheet(at: index) { sheet in
        sheet.writeText(name, to: "A1")
        // ... populate sheet
    }
}
```

## Protection

### Sheet Protection

Protect sheets to prevent unwanted modifications:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Basic protection with password
    sheet.protectSheet(password: "secret123")
    
    // Write protected content
    sheet.writeText("Protected Data", to: "A1")
}
```

### Protection Options

Customize what users can and cannot do:

```swift
// Use preset options
writer.modifySheet(at: sheetIndex) { sheet in
    // Strict: prevent all modifications
    sheet.protectSheet(password: "pwd", options: .strict)
    
    // Readonly: allow viewing only
    sheet.protectSheet(password: "pwd", options: .readonly)
}

// Custom protection
var options = SheetProtectionOptions()
options.formatCells = false       // Prevent formatting
options.insertRows = false         // Prevent row insertion
options.deleteColumns = true       // Allow column deletion
options.selectUnlockedCells = true // Allow selecting unlocked cells

writer.modifySheet(at: sheetIndex) { sheet in
    sheet.protectSheet(password: "pwd", options: options)
}
```

### Protection Options Reference

Available ``SheetProtectionOptions``:

```swift
public struct SheetProtectionOptions {
    var formatCells: Bool           // Format cells
    var formatColumns: Bool         // Format columns
    var formatRows: Bool            // Format rows
    var insertColumns: Bool         // Insert columns
    var insertRows: Bool            // Insert rows
    var insertHyperlinks: Bool      // Insert hyperlinks
    var deleteColumns: Bool         // Delete columns
    var deleteRows: Bool            // Delete rows
    var selectLockedCells: Bool     // Select locked cells
    var selectUnlockedCells: Bool   // Select unlocked cells
    var sort: Bool                  // Sort data
    var autoFilter: Bool            // Use auto filter
    var pivotTables: Bool           // Use pivot tables
}
```

### Workbook Protection

Protect workbook structure to prevent sheet operations:

```swift
var writer = WorkbookWriter()
_ = writer.addSheet(named: "Sheet1")

// Protect workbook structure
writer.protectWorkbook(password: "secret")

// Or use preset options
writer.protectWorkbook(options: .structureOnly)  // Prevent sheet add/delete/rename
writer.protectWorkbook(options: .strict)         // Protect structure and windows

// Custom options
var options = WorkbookProtectionOptions()
options.structure = true  // Prevent sheet operations
options.windows = false   // Allow window changes
writer.protectWorkbook(password: "pwd", options: options)
```

## Hyperlinks

### External Hyperlinks

Link to websites or files:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Basic URL
    sheet.addHyperlinkExternal(
        at: CellReference(column: "A", row: 1),
        url: "https://example.com"
    )
    
    // With display text
    sheet.addHyperlinkExternal(
        at: CellReference(column: "A", row: 2),
        url: "https://example.com",
        display: "Visit Example"
    )
    
    // With tooltip
    sheet.addHyperlinkExternal(
        at: CellReference(column: "A", row: 3),
        url: "https://example.com",
        display: "Example Site",
        tooltip: "Click to visit example.com"
    )
    
    // File link
    sheet.addHyperlinkExternal(
        at: CellReference(column: "B", row: 1),
        url: "file:///Users/name/Documents/report.pdf",
        display: "View Report"
    )
}
```

### Internal Hyperlinks

Link to locations within the workbook:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Link to another sheet
    sheet.addHyperlinkInternal(
        at: CellReference(column: "A", row: 1),
        location: "Sheet2!A1",
        display: "Go to Sheet2"
    )
    
    // Link to named range
    sheet.addHyperlinkInternal(
        at: CellReference(column: "A", row: 2),
        location: "SummaryData",
        display: "View Summary"
    )
    
    // With tooltip
    sheet.addHyperlinkInternal(
        at: CellReference(column: "A", row: 3),
        location: "Sheet2!B5",
        display: "Q2 Sales",
        tooltip: "Jump to Q2 sales data"
    )
}
```

## Merge Cells

Combine cells for headers or formatting:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Write header text
    sheet.writeText("Quarterly Report", to: "A1")
    
    // Merge header across multiple columns
    sheet.mergeCells("A1:D1")
    
    // Merge multiple ranges at once
    sheet.mergeCells([
        "A1:D1",  // Header row
        "A2:A4",  // Left sidebar
        "E2:F2"   // Another merged cell
    ])
}
```

## Data Validations

Add dropdown lists and input restrictions:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Dropdown list
    let validation = WorksheetBuilder.DataValidation(
        sqref: "B2:B10",
        validationType: .list,
        formula1: "\"Low,Medium,High\""
    )
    sheet.addDataValidation(validation)
    
    // Number range
    let numberValidation = WorksheetBuilder.DataValidation(
        sqref: "C2:C10",
        validationType: .whole,
        formula1: "1",
        formula2: "100"
    )
    sheet.addDataValidation(numberValidation)
}
```

## Auto Filters

Add column filtering capabilities:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Write headers
    sheet.writeText("Name", to: "A1")
    sheet.writeText("Age", to: "B1")
    sheet.writeText("City", to: "C1")
    
    // Add auto filter for entire data range
    sheet.setAutoFilter(range: "A1:C100")
}
```

## Conditional Formatting

Add visual formatting based on cell values:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    // Highlight cells greater than 100
    let rule = WorksheetData.ConditionalRule(
        type: "cellIs",
        operator: "greaterThan",
        formula: ["100"],
        dxfId: nil,
        priority: 1
    )
    sheet.addConditionalFormat(range: "B2:B10", rule: rule)
    
    // Multiple rules
    let rules = [
        WorksheetData.ConditionalRule(
            type: "cellIs",
            operator: "greaterThan",
            formula: ["100"],
            dxfId: nil,
            priority: 1
        ),
        WorksheetData.ConditionalRule(
            type: "cellIs",
            operator: "lessThan",
            formula: ["50"],
            dxfId: nil,
            priority: 2
        )
    ]
    sheet.addConditionalFormat(range: "C2:C10", rules: rules)
}
```

## Batch Operations

### Writing Large Datasets Efficiently

```swift
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

writer.modifySheet(at: sheetIndex) { sheet in
    // Write headers
    let headers = ["ID", "Name", "Email", "Score"]
    for (col, header) in headers.enumerated() {
        let colLetter = String(UnicodeScalar(65 + col)!)  // A, B, C, D
        sheet.writeText(header, to: CellReference(column: colLetter, row: 1))
    }
    
    // Write data rows
    for (index, record) in records.enumerated() {
        let row = index + 2  // Start at row 2 (after headers)
        sheet.writeNumber(Double(record.id), to: CellReference(column: "A", row: row))
        sheet.writeText(record.name, to: CellReference(column: "B", row: row))
        sheet.writeText(record.email, to: CellReference(column: "C", row: row))
        sheet.writeNumber(record.score, to: CellReference(column: "D", row: row))
    }
}

// Single save operation
try writer.save(to: outputURL)
```

### Performance Tips

```swift
// ✅ Good: Write all data, then save once
writer.modifySheet(at: sheetIndex) { sheet in
    for item in items {
        sheet.writeText(item.name, to: item.cell)
    }
}
try writer.save(to: url)  // Single ZIP operation

// ❌ Bad: Multiple saves
for item in items {
    writer.modifySheet(at: sheetIndex) { sheet in
        sheet.writeText(item.name, to: item.cell)
    }
    try writer.save(to: url)  // Very slow!
}
```

## Advanced Patterns

### Template-Based Writing

```swift
struct ReportTemplate {
    func apply(to sheet: inout WorkbookWriter.SheetWriter, data: ReportData) {
        // Title
        sheet.writeText(data.title, to: "A1")
        sheet.mergeCells("A1:D1")
        
        // Headers
        sheet.writeText("Metric", to: "A2")
        sheet.writeText("Value", to: "B2")
        sheet.writeText("Target", to: "C2")
        sheet.writeText("Status", to: "D2")
        
        // Data rows
        for (index, metric) in data.metrics.enumerated() {
            let row = index + 3
            sheet.writeText(metric.name, to: CellReference(column: "A", row: row))
            sheet.writeNumber(metric.value, to: CellReference(column: "B", row: row))
            sheet.writeNumber(metric.target, to: CellReference(column: "C", row: row))
            
            // Formula to determine status
            let formula = "IF(B\(row)>=C\(row),\"✓\",\"✗\")"
            sheet.writeFormula(formula, to: CellReference(column: "D", row: row))
        }
    }
}

// Usage
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Report")
writer.modifySheet(at: sheetIndex) { sheet in
    ReportTemplate().apply(to: &sheet, data: reportData)
}
```

### Multi-Sheet Reports

```swift
struct MultiSheetReport {
    let data: [String: [[String]]]
    
    func write(to writer: inout WorkbookWriter) {
        for (sheetName, sheetData) in data {
            let index = writer.addSheet(named: sheetName)
            writer.modifySheet(at: index) { sheet in
                for (rowIndex, row) in sheetData.enumerated() {
                    for (colIndex, value) in row.enumerated() {
                        let col = String(UnicodeScalar(65 + colIndex)!)
                        sheet.writeText(value, to: CellReference(column: col, row: rowIndex + 1))
                    }
                }
            }
        }
    }
}

// Usage
var writer = WorkbookWriter()
var report = MultiSheetReport(data: reportData)
report.write(to: &writer)
try writer.save(to: outputURL)
```

### Incremental Sheet Building

```swift
struct SheetBuilder {
    var sheet: WorkbookWriter.SheetWriter
    var currentRow = 1
    
    mutating func writeHeader(_ text: String) {
        sheet.writeText(text, to: CellReference(column: "A", row: currentRow))
        sheet.mergeCells("A\(currentRow):D\(currentRow)")
        currentRow += 1
    }
    
    mutating func writeRow(_ values: [String]) {
        for (index, value) in values.enumerated() {
            let col = String(UnicodeScalar(65 + index)!)
            sheet.writeText(value, to: CellReference(column: col, row: currentRow))
        }
        currentRow += 1
    }
    
    mutating func writeBlankRow() {
        currentRow += 1
    }
}

// Usage
writer.modifySheet(at: sheetIndex) { sheet in
    var builder = SheetBuilder(sheet: sheet)
    builder.writeHeader("Q1 Report")
    builder.writeRow(["Name", "Sales", "Target", "Status"])
    for record in records {
        builder.writeRow([record.name, "\(record.sales)", "\(record.target)", record.status])
    }
}
```

## Error Handling

Handle write operations safely:

```swift
do {
    var writer = WorkbookWriter()
    let sheetIndex = writer.addSheet(named: "Data")
    
    writer.modifySheet(at: sheetIndex) { sheet in
        sheet.writeText("Hello", to: "A1")
    }
    
    try writer.save(to: outputURL)
    print("Workbook saved successfully")
    
} catch let error as CuneiformError {
    switch error {
    case .accessDenied(let path):
        print("Cannot write to: \(path)")
    case .zipError(let message):
        print("ZIP error: \(message)")
    default:
        print("Error: \(error)")
    }
} catch {
    print("Unexpected error: \(error)")
}
```

## Complete Example

Putting it all together:

```swift
import Cuneiform

struct SalesReport {
    let title: String
    let data: [(name: String, q1: Double, q2: Double, q3: Double, q4: Double)]
}

func createSalesReport(_ report: SalesReport, outputURL: URL) throws {
    var writer = WorkbookWriter()
    let sheetIndex = writer.addSheet(named: "Sales Report")
    
    writer.modifySheet(at: sheetIndex) { sheet in
        // Title
        sheet.writeText(report.title, to: "A1")
        sheet.mergeCells("A1:F1")
        
        // Headers
        let headers = ["Name", "Q1", "Q2", "Q3", "Q4", "Total"]
        for (col, header) in headers.enumerated() {
            let colLetter = String(UnicodeScalar(65 + col)!)
            sheet.writeText(header, to: CellReference(column: colLetter, row: 2))
        }
        
        // Data rows
        for (index, record) in report.data.enumerated() {
            let row = index + 3
            sheet.writeText(record.name, to: CellReference(column: "A", row: row))
            sheet.writeNumber(record.q1, to: CellReference(column: "B", row: row))
            sheet.writeNumber(record.q2, to: CellReference(column: "C", row: row))
            sheet.writeNumber(record.q3, to: CellReference(column: "D", row: row))
            sheet.writeNumber(record.q4, to: CellReference(column: "E", row: row))
            
            // Total formula
            sheet.writeFormula(
                "SUM(B\(row):E\(row))",
                cachedValue: record.q1 + record.q2 + record.q3 + record.q4,
                to: CellReference(column: "F", row: row)
            )
        }
        
        // Summary row
        let summaryRow = report.data.count + 3
        sheet.writeText("TOTAL", to: CellReference(column: "A", row: summaryRow))
        for col in ["B", "C", "D", "E", "F"] {
            let firstRow = 3
            let lastRow = summaryRow - 1
            sheet.writeFormula(
                "SUM(\(col)\(firstRow):\(col)\(lastRow))",
                to: CellReference(column: col, row: summaryRow)
            )
        }
        
        // Add auto filter
        sheet.setAutoFilter(range: "A2:F\(summaryRow - 1)")
        
        // Protect sheet (allow filtering)
        var options = SheetProtectionOptions()
        options.autoFilter = true
        sheet.protectSheet(password: "report123", options: options)
    }
    
    // Protect workbook structure
    writer.protectWorkbook(options: .structureOnly)
    
    try writer.save(to: outputURL)
}

// Usage
let report = SalesReport(
    title: "2024 Sales Report",
    data: [
        ("Alice", 100, 120, 115, 130),
        ("Bob", 90, 95, 100, 105),
        ("Charlie", 110, 115, 120, 125)
    ]
)

try createSalesReport(report, outputURL: URL(fileURLWithPath: "sales_2024.xlsx"))
```

## See Also

- <doc:GettingStarted> - Basic writing examples
- <doc:PerformanceTuning> - Optimize write performance
- ``WorkbookWriter`` - API reference
- ``SheetProtectionOptions`` - Protection configuration
- ``WorkbookProtectionOptions`` - Workbook protection
