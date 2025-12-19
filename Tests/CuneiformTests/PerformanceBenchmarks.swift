import Testing
import Foundation
@testable import Cuneiform

@Suite("Performance Benchmarks")
struct PerformanceBenchmarks {
    
    /// Measure execution time of a block
    func measure(name: String, _ block: () throws -> Void) rethrows -> TimeInterval {
        let start = Date()
        try block()
        let elapsed = Date().timeIntervalSince(start)
        print("⏱️  \(name): \(String(format: "%.3f", elapsed))s")
        return elapsed
    }
    
    @Test("Benchmark: Read small workbook (100 rows)")
    func benchmarkReadSmall() throws {
        let tempFile = try createTestWorkbook(rows: 100, columns: 5)
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        _ = try measure(name: "Read 100 rows") {
            let workbook = try Workbook.open(url: tempFile)
            let sheet = try workbook.sheet(at: 0)!
            
            var cellCount = 0
            for row in sheet.rows() {
                cellCount += row.count
            }
            #expect(cellCount == 500) // 100 rows × 5 columns
        }
    }
    
    @Test("Benchmark: Read medium workbook (1000 rows)")
    func benchmarkReadMedium() throws {
        let tempFile = try createTestWorkbook(rows: 1000, columns: 10)
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        _ = try measure(name: "Read 1000 rows") {
            let workbook = try Workbook.open(url: tempFile)
            let sheet = try workbook.sheet(at: 0)!
            
            var cellCount = 0
            for row in sheet.rows() {
                cellCount += row.count
            }
            #expect(cellCount == 10000) // 1000 rows × 10 columns
        }
    }
    
    @Test("Benchmark: Write small workbook (100 rows)")
    func benchmarkWriteSmall() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("write-bench-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        _ = try measure(name: "Write 100 rows") {
            var writer = WorkbookWriter()
            let sheetIndex = writer.addSheet(named: "Data")
            
            writer.modifySheet(at: sheetIndex) { sheet in
                for row in 1...100 {
                    for col in 1...5 {
                        let ref = CellReference(column: String(UnicodeScalar(64 + col)!), row: row)
                        sheet.writeNumber(Double(row * col), to: ref)
                    }
                }
            }
            
            try writer.save(to: tempFile)
        }
    }
    
    @Test("Benchmark: Write medium workbook (1000 rows)")
    func benchmarkWriteMedium() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("write-med-bench-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        _ = try measure(name: "Write 1000 rows") {
            var writer = WorkbookWriter()
            let sheetIndex = writer.addSheet(named: "Data")
            
            writer.modifySheet(at: sheetIndex) { sheet in
                for row in 1...1000 {
                    for col in 1...10 {
                        let colLetter = columnLetter(col)
                        let ref = CellReference(column: colLetter, row: row)
                        sheet.writeNumber(Double(row * col), to: ref)
                    }
                }
            }
            
            try writer.save(to: tempFile)
        }
    }
    
    @Test("Benchmark: Round-trip (write + read 500 rows)")
    func benchmarkRoundTrip() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("roundtrip-bench-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        _ = try measure(name: "Round-trip 500 rows") {
            // Write
            var writer = WorkbookWriter()
            let sheetIndex = writer.addSheet(named: "Data")
            
            writer.modifySheet(at: sheetIndex) { sheet in
                for row in 1...500 {
                    sheet.writeNumber(Double(row), to: CellReference(column: "A", row: row))
                    sheet.writeText("Row \(row)", to: CellReference(column: "B", row: row))
                }
            }
            
            try writer.save(to: tempFile)
            
            // Read back
            let workbook = try Workbook.open(url: tempFile)
            let sheet = try workbook.sheet(at: 0)!
            
            var count = 0
            for row in sheet.rows() {
                count += row.count
            }
            #expect(count == 1000) // 500 rows × 2 columns
        }
    }
    
    @Test("Benchmark: Streaming vs eager loading")
    func benchmarkStreamingVsEager() throws {
        let tempFile = try createTestWorkbook(rows: 500, columns: 5)
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        let workbook = try Workbook.open(url: tempFile)
        let sheet = try workbook.sheet(at: 0)!
        
        // Streaming (lazy iteration)
        let streamingTime = try measure(name: "Streaming iteration") {
            var sum = 0.0
            for row in sheet.rows() {
                for (_, value) in row {
                    if case .number(let n) = value {
                        sum += n
                    }
                }
            }
            #expect(sum > 0)
        }
        
        // Eager (load all at once)
        let eagerTime = try measure(name: "Eager loading") {
            var sum = 0.0
            for rowIndex in 1...500 {
                let cells = sheet.row(rowIndex)
                for value in cells {
                    if case .number(let n) = value {
                        sum += n
                    }
                }
            }
            #expect(sum > 0)
        }
        
        print("📊 Streaming vs Eager: \(String(format: "%.2fx", eagerTime / streamingTime))")
    }
    
    @Test("Benchmark: Find operations")
    func benchmarkFindOperations() throws {
        let tempFile = try createTestWorkbook(rows: 1000, columns: 5)
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        let workbook = try Workbook.open(url: tempFile)
        let sheet = try workbook.sheet(at: 0)!
        
        _ = try measure(name: "Find first match in 1000 rows") {
            let result = sheet.find { _, value in
                value == .number(500)
            }
            #expect(result != nil)
        }
        
        _ = try measure(name: "Find all matches in 1000 rows") {
            let results = sheet.findAll { _, value in
                if case .number(let n) = value {
                    return n > 900
                }
                return false
            }
            #expect(results.count > 0)
        }
    }
    
    @Test("Benchmark: Range queries")
    func benchmarkRangeQueries() throws {
        let tempFile = try createTestWorkbook(rows: 100, columns: 26)
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        let workbook = try Workbook.open(url: tempFile)
        let sheet = try workbook.sheet(at: 0)!
        
        _ = try measure(name: "Range query A1:Z100 (2600 cells)") {
            let range = sheet.range("A1:Z100")
            #expect(range.count == 2600)
        }
        
        _ = try measure(name: "Column query (100 cells)") {
            let column = sheet.column("M")
            #expect(column.count == 100)
        }
    }
    
    // MARK: - Helper Methods
    
    /// Create a test workbook with specified size
    func createTestWorkbook(rows: Int, columns: Int) throws -> URL {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Data")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            for row in 1...rows {
                for col in 1...columns {
                    let colLetter = columnLetter(col)
                    let ref = CellReference(column: colLetter, row: row)
                    sheet.writeNumber(Double(row), to: ref)
                }
            }
        }
        
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("benchmark-\(UUID()).xlsx")
        try writer.save(to: tempFile)
        return tempFile
    }
    
    func columnLetter(_ index: Int) -> String {
        var result = ""
        var num = index
        
        while num > 0 {
            let remainder = (num - 1) % 26
            result = String(UnicodeScalar(65 + remainder)!) + result
            num = (num - 1) / 26
        }
        
        return result
    }
}
