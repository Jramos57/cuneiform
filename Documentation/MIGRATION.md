# Migration Guide

Guide for migrating to Cuneiform from other Swift Excel libraries.

## Table of Contents

1. [Why Migrate to Cuneiform?](#why-migrate-to-cuneiform)
2. [Feature Comparison](#feature-comparison)
3. [Common Patterns](#common-patterns)
4. [API Mappings](#api-mappings)
5. [Breaking Changes](#breaking-changes)
6. [Migration Checklist](#migration-checklist)

## Why Migrate to Cuneiform?

Cuneiform offers several advantages:

- **Pure Swift 6**: Modern concurrency, Sendable types, typed throws
- **Read & Write**: Both read and create Excel files
- **Advanced Queries**: Built-in filtering, finding, range access
- **Performance**: Memory-efficient streaming, lazy loading
- **Type Safety**: Strong typing for cell values with Swift enums
- **Formula Support**: Parse and create formulas
- **Zero Dependencies**: No external libraries required
- **Cross-Platform**: macOS, iOS, tvOS, watchOS, visionOS

## Feature Comparison

### Cuneiform vs CoreXLSX

| Feature | CoreXLSX | Cuneiform |
|---------|----------|-----------|
| Read .xlsx | ✅ | ✅ |
| Write .xlsx | ❌ | ✅ |
| Cell formulas | ❌ | ✅ |
| Swift 6 | ⚠️ Partial | ✅ Full |
| Sendable types | ❌ | ✅ |
| Lazy loading | ⚠️ Manual | ✅ Automatic |
| Query API | ❌ | ✅ |
| Type-safe values | ⚠️ XML | ✅ Enum |

### Cuneiform vs XLSXWriter (other platforms)

| Feature | XLSXWriter | Cuneiform |
|---------|------------|-----------|
| Language | Python | Swift |
| Read .xlsx | ❌ | ✅ |
| Write .xlsx | ✅ | ✅ |
| Formulas | ✅ | ✅ |
| Styling | ✅ | ⚠️ Future |
| Charts | ✅ | ⚠️ Future |

## Common Patterns

### Opening a File

**CoreXLSX:**
```swift
import CoreXLSX

guard let file = XLSXFile(filepath: path) else {
    fatalError("Cannot open file")
}

for wbk in try file.parseWorkbooks() {
    for (name, path) in try file.parseWorksheetPathsAndNames(workbook: wbk) {
        let worksheet = try file.parseWorksheet(at: path)
        // Process worksheet
    }
}
```

**Cuneiform:**
```swift
import Cuneiform

let workbook = try Workbook.open(url: fileURL)

// List all sheets
for sheetInfo in workbook.sheets {
    print(sheetInfo.name)
}

// Access specific sheet
if let sheet = try workbook.sheet(named: "Sheet1") {
    // Process sheet
}
```

### Reading Cell Values

**CoreXLSX:**
```swift
let worksheet = try file.parseWorksheet(at: path)
let sharedStrings = try file.parseSharedStrings()

for row in worksheet.data?.rows ?? [] {
    for cell in row.cells {
        let value: String
        if cell.type == .sharedString,
           let stringIndex = cell.value.flatMap(Int.init) {
            value = sharedStrings[stringIndex].text ?? ""
        } else {
            value = cell.value ?? ""
        }
    }
}
```

**Cuneiform:**
```swift
let sheet = try workbook.sheet(at: 0)!

// Single cell
if let value = sheet.cell(at: "A1") {
    switch value {
    case .text(let text): print(text)
    case .number(let num): print(num)
    case .boolean(let bool): print(bool)
    case .date(let date): print(date)
    case .empty: print("Empty")
    case .error(let err): print("Error: \(err)")
    }
}

// Iterate rows
for row in sheet.rows() {
    for (ref, value) in row {
        print("\(ref): \(value)")
    }
}
```

### Filtering Data

**CoreXLSX:**
```swift
// Manual filtering required
var filtered: [Row] = []
for row in worksheet.data?.rows ?? [] {
    // Check cells manually
    if let cell = row.cells.first(where: { $0.reference == "C1" }),
       cell.value == "Active" {
        filtered.append(row)
    }
}
```

**Cuneiform:**
```swift
// Built-in filtering
let activeRows = sheet.rows { cells in
    cells.contains { $0.value == .text("Active") }
}

// Or find specific cells
let results = sheet.findAll { _, value in
    value == .text("Active")
}
```

### Writing Files

**CoreXLSX:**
```swift
// Not supported - read-only library
```

**Cuneiform:**
```swift
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Name", to: "A1")
    sheet.writeNumber(42, to: "B1")
    sheet.writeFormula("SUM(B1:B10)", to: "B11")
}

try writer.save(to: outputURL)
```

## API Mappings

### Basic Operations

| Task | CoreXLSX | Cuneiform |
|------|----------|-----------|
| Open file | `XLSXFile(filepath:)` | `Workbook.open(url:)` |
| Get sheets | `parseWorksheetPathsAndNames()` | `workbook.sheets` |
| Access sheet | `parseWorksheet(at:)` | `workbook.sheet(named:)` |
| Get cell | Manual iteration | `sheet.cell(at:)` |
| Cell value | String parsing | `CellValue` enum |
| Shared strings | `parseSharedStrings()` | Automatic resolution |

### Advanced Operations

| Task | CoreXLSX | Cuneiform |
|------|----------|-----------|
| Find cells | Manual loop | `sheet.find(where:)` |
| Filter rows | Manual loop | `sheet.rows(where:)` |
| Get column | Manual iteration | `sheet.column(_:)` |
| Get range | Manual iteration | `sheet.range(_:)` |
| Formulas | Not parsed | `sheet.formula(at:)` |
| Streaming | Manual | `sheet.rows()` |

## Breaking Changes

When migrating from CoreXLSX or similar libraries:

### 1. Cell References

**Old:**
```swift
cell.reference  // String like "A1"
```

**New:**
```swift
cell.reference  // CellReference struct
cell.reference.column  // "A"
cell.reference.row     // 1
```

### 2. Cell Values

**Old:**
```swift
let value = cell.value  // String?
let type = cell.type    // CellType enum
// Manual type checking and conversion
```

**New:**
```swift
let value = sheet.cell(at: "A1")  // CellValue?
switch value {
case .text(let str): // Already a String
case .number(let num): // Already a Double
case .boolean(let bool): // Already a Bool
// ...
}
```

### 3. Shared Strings

**Old:**
```swift
let sharedStrings = try file.parseSharedStrings()
if cell.type == .sharedString,
   let index = Int(cell.value ?? "") {
    let value = sharedStrings[index].text
}
```

**New:**
```swift
// Automatic - no manual handling needed
let value = sheet.cell(at: ref)
// Shared strings resolved transparently
```

### 4. Error Handling

**Old:**
```swift
enum XLSXReaderError: Error {
    case archiveEntryNotFound
    // ...
}
```

**New:**
```swift
enum CuneiformError: Error, Sendable {
    case missingRequiredPart(PartPath)
    case malformedXML(part: String, detail: String)
    // ... more specific errors
}
```

## Migration Checklist

### Phase 1: Dependency Updates

- [ ] Add Cuneiform to Package.swift
```swift
dependencies: [
    .package(url: "https://github.com/yourusername/cuneiform.git", from: "1.0.0")
]
```

- [ ] Remove old Excel library dependencies
- [ ] Update imports in affected files

### Phase 2: API Migration

- [ ] Replace file opening code
- [ ] Update sheet access patterns
- [ ] Convert cell value handling to CellValue enum
- [ ] Update error handling for typed throws
- [ ] Replace manual shared string resolution

### Phase 3: Feature Additions

- [ ] Add write capabilities (if needed)
- [ ] Utilize advanced query APIs
- [ ] Implement streaming for large files
- [ ] Add formula parsing/writing (if needed)

### Phase 4: Testing & Validation

- [ ] Verify all existing functionality works
- [ ] Test with production data files
- [ ] Benchmark performance (should improve)
- [ ] Check memory usage (should decrease)

### Phase 5: Optimization

- [ ] Use streaming for large files
- [ ] Implement lazy loading patterns
- [ ] Add query optimizations
- [ ] Profile and tune as needed

## Example Migration

### Before (CoreXLSX)

```swift
import CoreXLSX

func processExcelFile(path: String) throws {
    guard let file = XLSXFile(filepath: path) else {
        throw MyError.cannotOpenFile
    }
    
    let sharedStrings = try file.parseSharedStrings()
    
    for wbk in try file.parseWorkbooks() {
        for (name, path) in try file.parseWorksheetPathsAndNames(workbook: wbk) {
            guard name == "Sales" else { continue }
            
            let worksheet = try file.parseWorksheet(at: path)
            
            var total = 0.0
            for row in worksheet.data?.rows ?? [] {
                for cell in row.cells {
                    guard cell.reference.hasPrefix("B") else { continue }
                    
                    if let valueStr = cell.value,
                       let value = Double(valueStr) {
                        total += value
                    }
                }
            }
            
            print("Total: \(total)")
        }
    }
}
```

### After (Cuneiform)

```swift
import Cuneiform

func processExcelFile(url: URL) throws {
    let workbook = try Workbook.open(url: url)
    
    guard let sheet = try workbook.sheet(named: "Sales") else {
        throw MyError.sheetNotFound
    }
    
    // Much simpler with built-in column access
    let salesColumn = sheet.column("B")
    
    let total = salesColumn.reduce(0.0) { sum, (_, value) in
        if case .number(let num) = value {
            return sum + num
        }
        return sum
    }
    
    print("Total: \(total)")
}
```

### Benefits of Migration

- **38 lines → 21 lines**: Simpler, clearer code
- **Automatic shared string resolution**: No manual handling
- **Type-safe values**: No string parsing
- **Built-in queries**: Column access, filtering, finding
- **Error safety**: Typed throws with specific error cases
- **Performance**: Lazy loading and streaming built-in

## Getting Help

If you encounter issues during migration:

1. Check the [Performance Guide](PERFORMANCE.md) for optimization tips
2. Review the [Examples](../Examples/) for common patterns
3. See the [README](../README.md) for API documentation
4. Run the test suite to understand behavior:
   ```bash
   swift test
   ```

## Common Pitfalls

### 1. Cell References

**Pitfall**: Assuming string-based references work everywhere

```swift
// ❌ Wrong
let value = sheet.cell(at: "A" + String(row))

// ✅ Correct
let ref = CellReference(column: "A", row: row)
let value = sheet.cell(at: ref)
```

### 2. Shared Strings

**Pitfall**: Trying to manually resolve shared strings

```swift
// ❌ Wrong - don't do this
// Cuneiform handles it automatically

// ✅ Correct
let value = sheet.cell(at: ref)
// Already resolved if it was a shared string
```

### 3. Memory Management

**Pitfall**: Loading entire sheets for simple queries

```swift
// ❌ Wrong
var allRows: [[CellValue]] = []
for i in 1...10000 {
    allRows.append(sheet.row(i))
}
let filtered = allRows.filter { /* ... */ }

// ✅ Correct
let filtered = sheet.rows { cells in
    // Filters during iteration
}
```

## Summary

Migrating to Cuneiform provides:

- **Simpler API**: Less boilerplate, clearer intent
- **Type Safety**: Swift enums instead of string parsing
- **More Features**: Write support, formulas, advanced queries
- **Better Performance**: Lazy loading, streaming, optimization
- **Modern Swift**: Swift 6, Sendable, typed throws

The migration process is straightforward and typically results in cleaner, more maintainable code with better performance characteristics.
