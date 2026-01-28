# Error Handling

Comprehensive guide to handling errors and failures in Cuneiform.

## Overview

Cuneiform uses Swift 6's typed throws feature to provide precise, actionable error information. All public APIs that can fail throw ``CuneiformError``, making error handling predictable and type-safe.

## CuneiformError

The ``CuneiformError`` enum defines all possible error conditions:

```swift
public enum CuneiformError: Error, Sendable {
    // Package Errors
    case invalidZipArchive(reason: String)
    case missingPart(path: String)
    case invalidContentType(path: String, expected: String, found: String)
    case invalidPackageStructure(reason: String)
    
    // XML Parsing Errors
    case malformedXML(part: String, detail: String)
    case missingRequiredElement(element: String, inPart: String)
    case invalidAttributeValue(attribute: String, value: String, inPart: String)
    
    // Data Errors
    case invalidCellReference(String)
    case sharedStringIndexOutOfRange(index: Int, count: Int)
    case styleIndexOutOfRange(index: Int, count: Int)
    
    // File Errors
    case fileNotFound(path: String)
    case accessDenied(path: String)
    case notAnXlsxFile(path: String)
}
```

## Basic Error Handling

### Reading Files

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
    let sheet = try workbook.sheet(at: 0)!
    
    // Process data
    for row in sheet.rows() {
        process(row)
    }
    
} catch let error as CuneiformError {
    print("Error: \(error)")
} catch {
    print("Unexpected error: \(error)")
}
```

### Writing Files

```swift
do {
    var writer = WorkbookWriter()
    let sheetIndex = writer.addSheet(named: "Data")
    
    writer.modifySheet(at: sheetIndex) { sheet in
        sheet.writeText("Hello", to: "A1")
    }
    
    try writer.save(to: outputURL)
    
} catch let error as CuneiformError {
    print("Failed to save: \(error)")
} catch {
    print("Unexpected error: \(error)")
}
```

## Error Categories

### Package Errors

Errors related to the ZIP archive structure:

#### invalidZipArchive

The file is not a valid ZIP archive:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.invalidZipArchive(let reason) {
    print("Not a valid ZIP: \(reason)")
    // User action: Check file format, try re-downloading
}
```

#### missingPart

A required part of the Office Open XML structure is missing:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.missingPart(let path) {
    print("Missing required part: \(path)")
    // Common parts: [Content_Types].xml, xl/workbook.xml, etc.
    // User action: File may be corrupted, try re-saving from Excel
}
```

#### invalidPackageStructure

The package structure doesn't conform to the standard:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.invalidPackageStructure(let reason) {
    print("Invalid structure: \(reason)")
    // User action: File may be corrupted or use unsupported features
}
```

### XML Parsing Errors

Errors when parsing XML content:

#### malformedXML

XML content cannot be parsed:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.malformedXML(let part, let detail) {
    print("Malformed XML in '\(part)': \(detail)")
    // Common causes: File corruption, manual editing, non-Excel tool
    // User action: Try opening and re-saving in Excel
}
```

#### missingRequiredElement

A required XML element is absent:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.missingRequiredElement(let element, let inPart) {
    print("Missing <\(element)> in \(inPart)")
    // User action: File may be incomplete or corrupted
}
```

#### invalidAttributeValue

An XML attribute has an invalid value:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.invalidAttributeValue(let attribute, let value, let inPart) {
    print("Invalid \(attribute)='\(value)' in \(inPart)")
    // User action: File may use unsupported feature or be corrupted
}
```

### Data Errors

Errors related to spreadsheet data:

#### invalidCellReference

A cell reference string cannot be parsed:

```swift
// Invalid reference format
do {
    let ref = CellReference("Invalid")
    // Won't reach here - returns nil
} catch {
    // CellReference initializer returns nil for invalid refs
}

// May occur when parsing formulas or ranges
do {
    let sheet = try workbook.sheet(at: 0)!
    let range = sheet.range("InvalidRange")
} catch CuneiformError.invalidCellReference(let ref) {
    print("Invalid reference: \(ref)")
}
```

#### sharedStringIndexOutOfRange

A cell references a non-existent shared string:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
    let sheet = try workbook.sheet(at: 0)!
    let value = sheet.cell(at: "A1")
} catch CuneiformError.sharedStringIndexOutOfRange(let index, let count) {
    print("Shared string index \(index) out of range (count: \(count))")
    // User action: File corruption likely
}
```

#### styleIndexOutOfRange

A cell references a non-existent style:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
    let sheet = try workbook.sheet(at: 0)!
    let style = sheet.cellStyle(at: "A1")
} catch CuneiformError.styleIndexOutOfRange(let index, let count) {
    print("Style index \(index) out of range (count: \(count))")
    // User action: File corruption likely
}
```

### File Errors

Errors related to file system operations:

#### fileNotFound

The specified file doesn't exist:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.fileNotFound(let path) {
    print("File not found: \(path)")
    // User action: Check path, verify file exists
}
```

#### accessDenied

Permission denied to read or write the file:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.accessDenied(let path) {
    print("Access denied: \(path)")
    // User action: Check file permissions, close file in other apps
}
```

#### notAnXlsxFile

The file is not an .xlsx file:

```swift
do {
    let workbook = try Workbook.open(url: fileURL)
} catch CuneiformError.notAnXlsxFile(let path) {
    print("Not an xlsx file: \(path)")
    // User action: Verify file extension and format
}
```

## Error Recovery Patterns

### Graceful Degradation

Continue processing despite errors:

```swift
func processWorkbook(url: URL) {
    do {
        let workbook = try Workbook.open(url: url)
        
        // Try to process each sheet, skipping failures
        for sheetInfo in workbook.sheets {
            do {
                if let sheet = try workbook.sheet(named: sheetInfo.name) {
                    processSheet(sheet)
                }
            } catch {
                print("Skipping sheet '\(sheetInfo.name)': \(error)")
                continue
            }
        }
        
    } catch {
        print("Cannot open workbook: \(error)")
    }
}
```

### Retry with Fallback

Try alternative approaches:

```swift
func readData(from url: URL) -> [String] {
    do {
        let workbook = try Workbook.open(url: url)
        
        // Try specific sheet first
        if let sheet = try? workbook.sheet(named: "Data") {
            return extractData(from: sheet)
        }
        
        // Fallback to first sheet
        if let sheet = try? workbook.sheet(at: 0) {
            print("Warning: Using first sheet as fallback")
            return extractData(from: sheet)
        }
        
        return []
        
    } catch {
        print("Error: \(error)")
        return []
    }
}
```

### Validation Before Operations

Check preconditions to avoid errors:

```swift
func writeToFile(url: URL, data: [[String]]) throws {
    // Check if directory exists
    let directory = url.deletingLastPathComponent()
    var isDirectory: ObjCBool = false
    guard FileManager.default.fileExists(atPath: directory.path, isDirectory: &isDirectory),
          isDirectory.boolValue else {
        throw NSError(domain: "Directory doesn't exist", code: 1)
    }
    
    // Check write permissions
    guard FileManager.default.isWritableFile(atPath: directory.path) else {
        throw NSError(domain: "No write permission", code: 2)
    }
    
    // Now safe to write
    var writer = WorkbookWriter()
    let sheetIndex = writer.addSheet(named: "Data")
    
    writer.modifySheet(at: sheetIndex) { sheet in
        for (rowIndex, row) in data.enumerated() {
            for (colIndex, value) in row.enumerated() {
                let col = String(UnicodeScalar(65 + colIndex)!)
                sheet.writeText(value, to: CellReference(column: col, row: rowIndex + 1))
            }
        }
    }
    
    try writer.save(to: url)
}
```

### User-Friendly Error Messages

Convert errors to user-facing messages:

```swift
func displayError(_ error: Error) -> String {
    guard let cuneiformError = error as? CuneiformError else {
        return "An unexpected error occurred: \(error.localizedDescription)"
    }
    
    switch cuneiformError {
    case .fileNotFound:
        return "The file could not be found. Please check the file path and try again."
        
    case .accessDenied:
        return "Access to the file was denied. Please check file permissions and ensure the file is not open in another application."
        
    case .invalidZipArchive:
        return "The file is not a valid Excel file. Please ensure you're opening an .xlsx file."
        
    case .missingPart:
        return "The file appears to be corrupted or incomplete. Try opening and re-saving it in Excel."
        
    case .malformedXML:
        return "The file contains invalid data. This may indicate file corruption or unsupported features."
        
    case .notAnXlsxFile:
        return "This is not an Excel .xlsx file. Please select a valid Excel workbook."
        
    case .invalidCellReference(let ref):
        return "Invalid cell reference: '\(ref)'. Please check your formula or range specification."
        
    case .sharedStringIndexOutOfRange, .styleIndexOutOfRange:
        return "The file contains corrupted data. Try opening and re-saving it in Excel."
        
    case .invalidPackageStructure:
        return "The file structure is not recognized. The file may be corrupted or created by an incompatible tool."
        
    case .missingRequiredElement, .invalidAttributeValue, .invalidContentType:
        return "The file contains data that cannot be processed. This may indicate file corruption or use of unsupported Excel features."
    }
}

// Usage
do {
    let workbook = try Workbook.open(url: fileURL)
    // ...
} catch {
    let userMessage = displayError(error)
    showAlert(userMessage)
}
```

## Error Logging

### Structured Logging

Log errors with context:

```swift
import os.log

let logger = Logger(subsystem: "com.example.app", category: "excel")

func processFile(url: URL) {
    logger.info("Processing file: \(url.path)")
    
    do {
        let workbook = try Workbook.open(url: url)
        logger.info("Opened workbook with \(workbook.sheets.count) sheets")
        
        for sheetInfo in workbook.sheets {
            logger.debug("Processing sheet: \(sheetInfo.name)")
            
            do {
                if let sheet = try workbook.sheet(named: sheetInfo.name) {
                    processSheet(sheet)
                    logger.info("Successfully processed sheet: \(sheetInfo.name)")
                }
            } catch let error as CuneiformError {
                logger.error("Failed to process sheet '\(sheetInfo.name)': \(error.description)")
            }
        }
        
    } catch let error as CuneiformError {
        logger.error("Failed to open workbook: \(error.description)")
    } catch {
        logger.fault("Unexpected error: \(error.localizedDescription)")
    }
}
```

### Debug Information

Capture detailed error context:

```swift
struct ErrorContext {
    let operation: String
    let fileURL: URL
    let timestamp: Date
    let error: CuneiformError
    
    var debugDescription: String {
        """
        Operation: \(operation)
        File: \(fileURL.path)
        Time: \(timestamp)
        Error: \(error)
        """
    }
}

func captureError(_ error: CuneiformError, operation: String, fileURL: URL) -> ErrorContext {
    ErrorContext(
        operation: operation,
        fileURL: fileURL,
        timestamp: Date(),
        error: error
    )
}

// Usage
do {
    let workbook = try Workbook.open(url: fileURL)
} catch let error as CuneiformError {
    let context = captureError(error, operation: "open_workbook", fileURL: fileURL)
    print(context.debugDescription)
    // Send to crash reporting service, save to log file, etc.
}
```

## Testing Error Conditions

### Unit Tests

Test error handling in your code:

```swift
import Testing
@testable import Cuneiform

@Test func testFileNotFound() throws {
    let nonExistentURL = URL(fileURLWithPath: "/nonexistent/file.xlsx")
    
    #expect(throws: CuneiformError.self) {
        try Workbook.open(url: nonExistentURL)
    }
}

@Test func testInvalidZipArchive() throws {
    // Create a file that's not a ZIP
    let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("invalid.xlsx")
    try "Not a ZIP file".write(to: tempURL, atomically: true, encoding: .utf8)
    defer { try? FileManager.default.removeItem(at: tempURL) }
    
    #expect(throws: CuneiformError.invalidZipArchive) {
        try Workbook.open(url: tempURL)
    }
}

@Test func testAccessDenied() throws {
    let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent("readonly.xlsx")
    // Create file and make it read-only
    // ... test implementation
}
```

## Best Practices

### 1. Always Handle Typed Errors

```swift
// ✅ Good: Handle specific error types
do {
    let workbook = try Workbook.open(url: fileURL)
} catch let error as CuneiformError {
    handleCuneiformError(error)
} catch {
    handleUnexpectedError(error)
}

// ❌ Bad: Generic catch
do {
    let workbook = try Workbook.open(url: fileURL)
} catch {
    print("Error: \(error)")  // Loses type information
}
```

### 2. Provide Context in Error Messages

```swift
// ✅ Good: Include context
func processWorkbook(url: URL, name: String) {
    do {
        let workbook = try Workbook.open(url: url)
        // ...
    } catch {
        print("Failed to process workbook '\(name)' at \(url.path): \(error)")
    }
}

// ❌ Bad: Generic message
func processWorkbook(url: URL) {
    do {
        let workbook = try Workbook.open(url: url)
    } catch {
        print("Error")  // Unhelpful
    }
}
```

### 3. Don't Swallow Errors Silently

```swift
// ✅ Good: Log or propagate errors
func processFile(url: URL) throws {
    do {
        let workbook = try Workbook.open(url: url)
        // ...
    } catch let error as CuneiformError {
        logger.error("Failed to open: \(error)")
        throw error  // Propagate
    }
}

// ❌ Bad: Silent failure
func processFile(url: URL) {
    do {
        let workbook = try Workbook.open(url: url)
    } catch {
        // Silent - caller has no idea it failed
    }
}
```

### 4. Clean Up Resources

```swift
// ✅ Good: Use defer for cleanup
func processFiles(urls: [URL]) {
    for url in urls {
        var tempFile: URL?
        
        defer {
            if let temp = tempFile {
                try? FileManager.default.removeItem(at: temp)
            }
        }
        
        do {
            // Process file
            let workbook = try Workbook.open(url: url)
            // ...
        } catch {
            print("Error processing \(url): \(error)")
            continue
        }
    }
}
```

### 5. Validate Input Early

```swift
// ✅ Good: Validate before expensive operations
func importData(from url: URL) throws {
    // Quick checks first
    guard FileManager.default.fileExists(atPath: url.path) else {
        throw CuneiformError.fileNotFound(path: url.path)
    }
    
    guard url.pathExtension.lowercased() == "xlsx" else {
        throw CuneiformError.notAnXlsxFile(path: url.path)
    }
    
    // Now do expensive operations
    let workbook = try Workbook.open(url: url)
    // ...
}
```

## See Also

- <doc:Architecture> - Understanding error sources
- ``CuneiformError`` - Complete error reference
- ``Workbook`` - Reading API with typed throws
- ``WorkbookWriter`` - Writing API with typed throws
