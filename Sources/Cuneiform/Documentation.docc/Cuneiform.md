# ``Cuneiform``

Pure Swift library for reading and writing Office Open XML SpreadsheetML (.xlsx) files.

@Metadata {
    @DisplayName("Cuneiform")
    @PageImage(purpose: icon, source: "cuneiform-icon", alt: "A spreadsheet grid icon")
    @PageColor(blue)
}

## Overview

**Cuneiform** is a pure Swift implementation of the Office Open XML SpreadsheetML standard (ECMA-376 / ISO/IEC 29500). It provides comprehensive support for reading and writing Excel .xlsx files with a modern, idiomatic Swift API.

Built with Swift 6, Cuneiform leverages value semantics, Sendable types, and typed throws to provide a safe, performant, and ergonomic experience for working with spreadsheet data.

### Key Features

- **Full OOXML Compliance**: ~90-92% standard compliance with complete SpreadsheetML support
- **Comprehensive Formula Engine**: 467 Excel formula functions with full evaluation (97% full implementations)
- **Read and Write**: Complete support for opening, querying, creating, and modifying .xlsx files
- **High Performance**: Lazy loading, streaming iteration, and efficient memory usage for large files
- **Advanced Features**: Charts, pivot tables, data validations, named ranges, hyperlinks, and tables
- **Protection**: Sheet protection and workbook protection with password support
- **Modern Swift**: Swift 6, Sendable, value semantics, typed throws, comprehensive error handling
- **Thoroughly Tested**: 834 passing tests ensuring reliability

### Quick Example

```swift
import Cuneiform

// Open an .xlsx file
let workbook = try Workbook.open(url: URL(fileURLWithPath: "data.xlsx"))

// Access a sheet and read cells
if let sheet = try workbook.sheet(named: "Sheet1") {
    if let value = sheet.cell(at: "A1") {
        print("A1: \(value)")  // text, number, date, boolean, error, or empty
    }
}

// Create a new workbook
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

// Write cells
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Name", to: "A1")
    sheet.writeNumber(42, to: "B1")
    sheet.writeFormula("B1*2", cachedValue: 84, to: "C1")
}

// Save to file
try writer.save(to: URL(fileURLWithPath: "output.xlsx"))
```

## Topics

### Essentials

- <doc:GettingStarted>
- ``Workbook``
- ``WorkbookWriter``
- ``Sheet``

### Tutorials

- <doc:DataAnalysis>
- <doc:ReportGeneration>

### Reading Workbooks

- ``Workbook``
- ``Sheet``
- ``CellValue``
- ``SheetInfo``
- ``RawCellValue``

### Writing Workbooks

- ``WorkbookWriter``
- ``WorkbookWriterSheet``
- ``CellReference``

### Formula Engine

- <doc:FormulaEngine>
- ``FormulaEvaluator``
- ``FormulaParser``
- <doc:FormulaReference>

### Advanced Features

- <doc:AdvancedQueries>
- <doc:WritingWorkbooks>
- <doc:PerformanceTuning>
- ``DefinedName``
- ``DataValidation``
- ``Hyperlink``
- ``TableData``
- ``PivotTableData``
- ``ChartData``

### Protection

- ``SheetProtection``
- ``SheetProtectionOptions``
- ``WorkbookProtection``
- ``WorkbookProtectionOptions``

### Architecture and Guides

- <doc:Architecture>
- <doc:ErrorHandling>
- <doc:MigrationGuide>

### Core Types

- ``OPCPackage``
- ``PartPath``
- ``Relationship``
- ``ContentType``

### Parsers

- ``WorkbookParser``
- ``WorksheetParser``
- ``SharedStringsParser``
- ``StylesParser``

### Builders

- ``WorkbookBuilder``
- ``WorksheetBuilder``
- ``SharedStringsBuilder``
- ``ContentTypesBuilder``
- ``RelationshipsBuilder``

### Error Handling

- ``CuneiformError``
