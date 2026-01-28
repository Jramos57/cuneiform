# Getting Started

Learn how to read and write Excel files with Cuneiform.

## Overview

Cuneiform provides a simple, Swift-native API for working with Excel .xlsx files. This guide will walk you through installation and basic usage patterns to get you up and running quickly.

## Installation

### Swift Package Manager

Add Cuneiform to your `Package.swift` dependencies:

```swift
dependencies: [
    .package(url: "https://github.com/jramos57/cuneiform.git", from: "0.1.0")
]
```

Then add it to your target dependencies:

```swift
targets: [
    .target(
        name: "YourTarget",
        dependencies: ["Cuneiform"]
    )
]
```

### Xcode

1. In Xcode, select **File → Add Package Dependencies...**
2. Enter the repository URL: `https://github.com/jramos57/cuneiform.git`
3. Select version `0.1.0` or later
4. Add the package to your target

## Reading Your First Workbook

Opening and reading an Excel file is straightforward:

```swift
import Cuneiform

// Open an .xlsx file
let workbook = try Workbook.open(url: URL(fileURLWithPath: "data.xlsx"))

// List all sheets
for sheet in workbook.sheets {
    print("Sheet: \(sheet.name)")
}

// Access a sheet by name
if let sheet = try workbook.sheet(named: "Sheet1") {
    // Get a cell value
    if let value = sheet.cell(at: "A1") {
        print("A1: \(value)")
    }
    
    // Get a row
    let row = sheet.row(1)
    for (ref, value) in row {
        print("\(ref): \(value)")
    }
}
```

### Cell Values

The ``CellValue`` enum represents fully-resolved cell content:

```swift
let value = sheet.cell(at: "B2")

switch value {
case .text(let str):
    print("Text: \(str)")
case .number(let num):
    print("Number: \(num)")
case .date(let dateStr):
    print("Date: \(dateStr)")  // ISO 8601 format
case .boolean(let bool):
    print("Boolean: \(bool)")
case .error(let error):
    print("Error: \(error)")
case .empty:
    print("Empty cell")
case .none:
    print("Cell doesn't exist")
}
```

## Writing a Workbook

Creating and writing Excel files is just as simple:

```swift
import Cuneiform

// Create a new workbook
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

// Write different cell types
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Name", to: "A1")
    sheet.writeText("Score", to: "B1")
    sheet.writeText("Pass", to: "C1")
    
    sheet.writeText("Alice", to: "A2")
    sheet.writeNumber(95, to: "B2")
    sheet.writeBoolean(true, to: "C2")
    
    sheet.writeText("Bob", to: "A3")
    sheet.writeNumber(72, to: "B3")
    sheet.writeBoolean(false, to: "C3")
}

// Save to file
try writer.save(to: URL(fileURLWithPath: "output.xlsx"))
```

### Writing Formulas

You can write formulas with optional cached values:

```swift
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeNumber(10, to: "A1")
    sheet.writeNumber(20, to: "A2")
    
    // Formula with cached value
    sheet.writeFormula("SUM(A1:A2)", cachedValue: 30, to: "A3")
    
    // Formula without cached value (Excel will calculate on open)
    sheet.writeFormula("AVERAGE(A1:A2)", to: "A4")
}
```

## Working with Ranges

Access and iterate over ranges of cells:

```swift
// Access a range
let range = sheet.range("A1:C10")
for (ref, value) in range {
    print("\(ref): \(value)")
}

// Access entire columns
let columnA = sheet.column("A")
let columnB = sheet.column(at: 1)  // 0-based index

// Filter rows
let highScores = sheet.rows { cells in
    cells.contains { cell in
        if case .number(let n) = cell.value {
            return n >= 90
        }
        return false
    }
}
```

## Working with Formulas

Cuneiform includes a comprehensive formula evaluator with 467 Excel functions:

```swift
import Cuneiform

// Create a formula evaluator
var evaluator = FormulaEvaluator()

// Set cell values
evaluator.setCell("A1", value: .number(10))
evaluator.setCell("A2", value: .number(20))
evaluator.setCell("A3", value: .number(30))

// Evaluate formulas
let sum = evaluator.evaluate("=SUM(A1:A3)")       // 60.0
let avg = evaluator.evaluate("=AVERAGE(A1:A3)")   // 20.0
let max = evaluator.evaluate("=MAX(A1:A3)")       // 30.0

// Complex formulas
let result = evaluator.evaluate("=IF(SUM(A1:A3)>50, \"High\", \"Low\")")  // "High"
```

See <doc:FormulaEngine> and <doc:FormulaReference> for complete formula documentation.

## Finding Data

Cuneiform provides powerful search capabilities:

```swift
// Find first matching cell
if let cell = sheet.find(where: { _, value in 
    value == .text("Alice") 
}) {
    print("Found at \(cell.reference): \(cell.value)")
}

// Find all matching cells
let matches = sheet.findAll { _, value in
    if case .number(let n) = value {
        return n > 50
    }
    return false
}

for match in matches {
    print("\(match.reference): \(match.value)")
}
```

## Performance Tips

For large files, use streaming iteration to minimize memory:

```swift
// Streaming iteration (memory-efficient)
for row in sheet.rows() {
    for (ref, value) in row {
        process(ref, value)
    }
}
```

Sheets are loaded lazily—only when accessed via `sheet(named:)` or `sheet(at:)`. This improves startup time for workbooks with many sheets.

See <doc:PerformanceTuning> for detailed optimization strategies.

## Next Steps

Now that you understand the basics, explore more advanced features:

- **<doc:DataAnalysis>** - Learn to extract and analyze spreadsheet data
- **<doc:ReportGeneration>** - Create structured reports programmatically
- **<doc:FormulaEngine>** - Deep dive into the formula evaluation engine
- **<doc:AdvancedQueries>** - Master complex data queries and filtering
- **<doc:WritingWorkbooks>** - Learn about styling, protection, and advanced writing features
- **<doc:PerformanceTuning>** - Optimize for your specific use case

## See Also

- ``Workbook``
- ``WorkbookWriter``
- ``Sheet``
- ``CellValue``
- ``FormulaEvaluator``
