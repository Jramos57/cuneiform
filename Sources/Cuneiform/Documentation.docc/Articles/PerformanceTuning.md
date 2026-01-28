# Performance Tuning

Comprehensive guide to optimizing Cuneiform performance for your use case.

## Overview

Cuneiform is designed for performance out of the box, but understanding its memory management and access patterns can help you optimize for specific workloads. This guide covers strategies for reading, writing, and querying Excel files efficiently.

## Memory Management

### Understanding Lazy Loading

Cuneiform uses lazy evaluation wherever possible to minimize memory usage:

- **Sheets are loaded on-demand**: Only accessed sheets are parsed
- **Row streaming available**: Use `sheet.rows()` for memory-efficient iteration
- **No caching by default**: Re-accessing cells re-resolves values (trade-off for memory)

### Memory-Efficient Patterns

#### ✅ Good: Streaming Rows

```swift
// Process large files without loading everything into memory
for row in sheet.rows() {
    for (ref, value) in row {
        process(ref, value)
    }
}
```

#### ❌ Bad: Loading All Rows at Once

```swift
// Loads all rows into memory immediately
var allData: [[CellValue]] = []
for i in 1...sheet.rowCount {
    allData.append(sheet.row(i))  // Inefficient for large files
}
```

### Memory Footprint by File Size

Approximate memory usage (measured on typical hardware):

| File Size | Rows × Cols | Memory (Eager) | Memory (Streaming) |
|-----------|-------------|----------------|-------------------|
| Small     | 100 × 10    | ~1 MB         | ~0.5 MB          |
| Medium    | 1,000 × 10  | ~10 MB        | ~2 MB            |
| Large     | 10,000 × 10 | ~100 MB       | ~5 MB            |
| Very Large| 100,000 × 10| ~1 GB         | ~20 MB           |

## Reading Strategies

### For Small Files (< 1,000 rows)

Direct access is fast and convenient:

```swift
let workbook = try Workbook.open(url: fileURL)
let sheet = try workbook.sheet(at: 0)!

// Access cells directly
let value = sheet.cell(at: "A1")

// Use convenience methods
let column = sheet.column("B")
let range = sheet.range("A1:C10")
```

### For Medium Files (1,000 - 10,000 rows)

Use streaming to balance performance and memory:

```swift
let workbook = try Workbook.open(url: fileURL)
let sheet = try workbook.sheet(at: 0)!

// Stream rows to limit memory usage
for row in sheet.rows() {
    // Process each row
    for (ref, value) in row {
        process(ref, value)
    }
}
```

### For Large Files (> 10,000 rows)

Optimize for streaming and avoid random access:

```swift
// ✅ Good: Sequential streaming
for row in sheet.rows() {
    // Process in order
    handleRow(row)
}

// ❌ Bad: Random access in loop
for i in 1...10000 {
    let cell = sheet.cell(at: CellReference(column: "A", row: i))
    // Each access re-searches the data structure
}
```

### Multi-Sheet Workbooks

Only load sheets you need:

```swift
// ✅ Good: Load specific sheet
if let targetSheet = try workbook.sheet(named: "ImportantData") {
    process(targetSheet)
}

// ❌ Bad: Iterate all sheets unnecessarily
for sheetInfo in workbook.sheets {
    let sheet = try workbook.sheet(named: sheetInfo.name)
    // Loads and parses every sheet
}
```

## Writing Strategies

### Batch Writes

Write all data before saving:

```swift
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

// ✅ Good: Write everything, then save once
writer.modifySheet(at: sheetIndex) { sheet in
    for (index, item) in data.enumerated() {
        let row = index + 1
        sheet.writeText(item.name, to: CellReference(column: "A", row: row))
        sheet.writeNumber(item.value, to: CellReference(column: "B", row: row))
    }
}

try writer.save(to: url)  // Single write operation
```

### Avoid Incremental Saves

```swift
// ❌ Bad: Saving repeatedly
for item in data {
    writer.modifySheet(at: sheetIndex) { sheet in
        sheet.writeText(item.name, to: "A\(item.index)")
    }
    try writer.save(to: url)  // Multiple ZIP operations - very slow!
}
```

### Formula Caching

Provide cached values for formulas when possible:

```swift
// ✅ Good: Pre-compute and cache
let sum = values.reduce(0, +)
sheet.writeFormula("SUM(A1:A100)", cachedValue: sum, to: "A101")

// ⚠️ Works but Excel must recalculate on open
sheet.writeFormula("SUM(A1:A100)", to: "A101")
```

## Query Optimization

### Choose the Right Query Method

Different methods for different needs:

```swift
// Finding single value - stops at first match
if let result = sheet.find(where: { _, value in value == target }) {
    // Fast: stops searching after first match
}

// Finding all matches - searches entire sheet
let results = sheet.findAll(where: { _, value in value == target })
// Slower: must check every cell

// Column access - optimized for vertical data
let column = sheet.column("B")  // Faster than iterating all cells

// Range access - good for rectangular regions
let range = sheet.range("A1:C10")  // Efficient for known bounds
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

### Pattern Matching Performance

```swift
// ✅ Fast: Direct equality
if value == .number(100) { }

// ⚠️ Slower: Pattern matching with binding
if case .number(let n) = value, n > 100 { }

// Consider extracting value once
if case .number(let n) = value {
    if n > 100 { /* ... */ }
    if n < 200 { /* ... */ }
}
```

## Profiling & Measurement

### Built-in Benchmarks

Run the performance test suite:

```bash
cd cuneiform
swift test --filter PerformanceBenchmarks
```

Results show timing for:
- Reading various file sizes
- Writing various file sizes
- Round-trip operations
- Query operations
- Streaming vs eager loading

### Using Instruments

Profile with Xcode Instruments:

```bash
# Generate Instruments trace
swift build -c release
xcrun xctrace record --template 'Time Profiler' \
  --launch .build/release/YourApp
```

Focus on:
- **Time Profiler**: Identify hot paths
- **Allocations**: Track memory growth
- **Leaks**: Verify no retain cycles

### Custom Benchmarking

```swift
func measure(label: String, operation: () throws -> Void) rethrows {
    let start = Date()
    try operation()
    let elapsed = Date().timeIntervalSince(start)
    print("\(label): \(String(format: "%.3f", elapsed))s")
}

measure(label: "Read 1000 rows") {
    let workbook = try Workbook.open(url: fileURL)
    let sheet = try workbook.sheet(at: 0)!
    for row in sheet.rows() {
        _ = row.count
    }
}
```

## Best Practices

### 1. Use Streaming for Large Files

```swift
// Memory-efficient for any size
for row in sheet.rows() {
    process(row)
}
```

### 2. Access Sheets Selectively

```swift
// Only load what you need
if let sheet = try workbook.sheet(named: "ImportantSheet") {
    process(sheet)
}
```

### 3. Batch Write Operations

```swift
// Collect all data, write once
var writer = WorkbookWriter()
// ... configure all sheets
try writer.save(to: url)  // Single save
```

### 4. Pre-compile Cell References

```swift
// ✅ Good: Reuse references
let cellRef = CellReference(column: "A", row: 1)
for sheet in sheets {
    _ = sheet.cell(at: cellRef)
}

// ❌ Bad: Parse repeatedly
for sheet in sheets {
    _ = sheet.cell(at: "A1")  // Re-parses "A1" each time
}
```

### 5. Use Appropriate Query Methods

```swift
// For single value
sheet.find(where: predicate)

// For all matches
sheet.findAll(where: predicate)

// For column data
sheet.column("B")

// For rectangular region
sheet.range("A1:D10")
```

## Troubleshooting

### Problem: Slow Reading

**Symptoms**: Opening files takes seconds

**Possible Causes**:
1. Large shared strings table
2. Complex cell formulas
3. Many sheets in workbook

**Solutions**:
```swift
// Load only needed sheet
guard let sheet = try workbook.sheet(named: "TargetSheet") else { return }

// Use streaming
for row in sheet.rows() { /* ... */ }

// Avoid accessing all cells
// Use targeted queries instead
```

### Problem: High Memory Usage

**Symptoms**: App using gigabytes of RAM

**Possible Causes**:
1. Loading entire sheet into memory
2. Accumulating row data in arrays
3. Not releasing workbook references

**Solutions**:
```swift
// Use streaming
for row in sheet.rows() {
    process(row)
    // Row data released each iteration
}

// Process and discard
autoreleasepool {
    let workbook = try Workbook.open(url: url)
    process(workbook)
    // Released at end of pool
}
```

### Problem: Slow Writing

**Symptoms**: Writing takes much longer than reading

**Possible Causes**:
1. Multiple save operations
2. Large string tables (many unique strings)

**Solutions**:
```swift
// Write all data before saving
writer.modifySheet(at: 0) { sheet in
    for item in largeDataset {
        sheet.writeNumber(item.value, to: item.ref)
    }
}
try writer.save(to: url)  // Single ZIP operation
```

### Problem: Query Performance

**Symptoms**: findAll or rows(where:) is slow

**Solutions**:
```swift
// Use find() if you only need first match
if let first = sheet.find(where: predicate) {
    // Stops after first match
}

// Stream and filter together
for row in sheet.rows() {
    if matchesCriteria(row) {
        process(row)
        // Can break early if needed
    }
}
```

## Performance Summary

| Operation | Small Files | Large Files | Best Practice |
|-----------|-------------|-------------|--------------|
| Read | Direct access | Streaming | `sheet.rows()` |
| Write | Any method | Batch writes | Write once, save once |
| Find single | `find(where:)` | `find(where:)` | Stops at first match |
| Find all | `findAll(where:)` | Stream + filter | Avoid for very large files |
| Column access | `column(_:)` | `column(_:)` | Efficient vertical scan |
| Range access | `range(_:)` | Targeted ranges | Avoid large ranges |

## Benchmark Reference

From test suite (typical hardware):

- **Read 1,000 rows**: ~34ms
- **Write 1,000 rows (10 cols)**: ~50ms
- **Round-trip 500 rows**: ~80ms
- **Find in 1,000 rows**: <1ms
- **Range query 2,600 cells**: ~660ms
- **Streaming vs Eager**: 2-3x memory reduction

For more benchmarks, run:
```bash
swift test --filter PerformanceBenchmarks
```

## Advanced Optimization Techniques

### 1. Parallel Processing

Process multiple sheets concurrently:

```swift
import Foundation

let workbook = try Workbook.open(url: fileURL)

await withTaskGroup(of: Void.self) { group in
    for sheetInfo in workbook.sheets {
        group.addTask {
            if let sheet = try? workbook.sheet(named: sheetInfo.name) {
                await processSheet(sheet)
            }
        }
    }
}
```

### 2. Incremental Processing

Process data in chunks:

```swift
let chunkSize = 1000
var currentChunk: [CellValue] = []

for row in sheet.rows() {
    for (_, value) in row {
        currentChunk.append(value)
        
        if currentChunk.count >= chunkSize {
            processChunk(currentChunk)
            currentChunk.removeAll(keepingCapacity: true)
        }
    }
}

// Process remaining
if !currentChunk.isEmpty {
    processChunk(currentChunk)
}
```

### 3. Memory Pooling

Reuse allocations for repeated operations:

```swift
var referencePool = (0..<1000).map { CellReference(column: "A", row: $0 + 1) }

for (index, item) in data.enumerated() {
    let ref = referencePool[index]
    sheet.writeText(item.name, to: ref)
}
```

### 4. Lazy Computation

Defer expensive operations until needed:

```swift
struct LazySheet {
    let sheet: Sheet
    
    lazy var statistics: Statistics = {
        computeStatistics(from: sheet)
    }()
    
    lazy var uniqueValues: Set<CellValue> = {
        Set(sheet.column("A").map { $0.1 })
    }()
}
```

## See Also

- <doc:Architecture> - Understanding Cuneiform's internal design
- <doc:AdvancedQueries> - Efficient query patterns
- ``Workbook`` - High-level read API
- ``Sheet`` - Query and access methods
- ``WorkbookWriter`` - High-level write API
