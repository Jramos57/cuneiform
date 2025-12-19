import Testing
import Foundation
@testable import Cuneiform

@Suite("Advanced Query Tests")
struct AdvancedQueryTests {
    
    /// Create a test workbook with sample data
    func createTestWorkbook() throws -> (Workbook, Sheet) {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Data")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            // Header row
            sheet.writeText("Name", to: "A1")
            sheet.writeText("Age", to: "B1")
            sheet.writeText("City", to: "C1")
            
            // Data rows
            sheet.writeText("Alice", to: "A2")
            sheet.writeNumber(30, to: "B2")
            sheet.writeText("NYC", to: "C2")
            
            sheet.writeText("Bob", to: "A3")
            sheet.writeNumber(25, to: "B3")
            sheet.writeText("LA", to: "C3")
            
            sheet.writeText("Charlie", to: "A4")
            sheet.writeNumber(35, to: "B4")
            sheet.writeText("NYC", to: "C4")
            
            sheet.writeText("Diana", to: "A5")
            sheet.writeNumber(28, to: "B5")
            sheet.writeText("Boston", to: "C5")
        }
        
        // Save and reload
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("query-test-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        let workbook = try Workbook.open(url: fileURL)
        let sheet = try workbook.sheet(at: 0)!
        
        return (workbook, sheet)
    }
    
    @Test("Range access with A1:C3 notation")
    func rangeAccess() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let range = sheet.range("A1:C2")
        
        // Should have 6 cells (2 rows x 3 columns)
        #expect(range.count == 6)
        
        // Check first row
        #expect(range[0].value == .text("Name"))
        #expect(range[1].value == .text("Age"))
        #expect(range[2].value == .text("City"))
        
        // Check second row
        #expect(range[3].value == .text("Alice"))
        #expect(range[4].value == .number(30))
        #expect(range[5].value == .text("NYC"))
    }
    
    @Test("Range access single cell")
    func rangeSingleCell() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let range = sheet.range("B2:B2")
        
        #expect(range.count == 1)
        #expect(range[0].value == .number(30))
    }
    
    @Test("Range access invalid format")
    func rangeInvalidFormat() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let range = sheet.range("invalid")
        #expect(range.isEmpty)
    }
    
    @Test("Column access by letter")
    func columnAccessByLetter() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let colA = sheet.column("A")
        
        // Should have 5 cells (header + 4 data rows)
        #expect(colA.count == 5)
        #expect(colA[0].value == .text("Name"))
        #expect(colA[1].value == .text("Alice"))
        #expect(colA[2].value == .text("Bob"))
        #expect(colA[3].value == .text("Charlie"))
        #expect(colA[4].value == .text("Diana"))
    }
    
    @Test("Column access by index")
    func columnAccessByIndex() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let colB = sheet.column(at: 1) // B = index 1
        
        #expect(colB.count == 5)
        #expect(colB[0].value == .text("Age"))
        #expect(colB[1].value == .number(30))
        #expect(colB[2].value == .number(25))
        #expect(colB[3].value == .number(35))
        #expect(colB[4].value == .number(28))
    }
    
    @Test("Column access case insensitive")
    func columnCaseInsensitive() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let colLower = sheet.column("a")
        let colUpper = sheet.column("A")
        
        #expect(colLower.count == colUpper.count)
        #expect(colLower[0].value == colUpper[0].value)
    }
    
    @Test("Row filtering by predicate")
    func rowFiltering() throws {
        let (_, sheet) = try createTestWorkbook()
        
        // Find rows where city is NYC
        let nycRows = sheet.rows { cells in
            cells.contains { $0.reference.column == "C" && $0.value == .text("NYC") }
        }
        
        #expect(nycRows.count == 2) // Alice and Charlie
        
        // Check first NYC row (Alice)
        let aliceRow = nycRows[0]
        #expect(aliceRow[0].value == .text("Alice"))
        #expect(aliceRow[1].value == .number(30))
        
        // Check second NYC row (Charlie)
        let charlieRow = nycRows[1]
        #expect(charlieRow[0].value == .text("Charlie"))
        #expect(charlieRow[1].value == .number(35))
    }
    
    @Test("Row filtering by age threshold")
    func rowFilteringByAge() throws {
        let (_, sheet) = try createTestWorkbook()
        
        // Find rows where age >= 30
        let olderRows = sheet.rows { cells in
            cells.contains { cell in
                if case .number(let age) = cell.value {
                    return age >= 30
                }
                return false
            }
        }
        
        #expect(olderRows.count == 2) // Alice (30) and Charlie (35)
    }
    
    @Test("Find first matching cell")
    func findFirst() throws {
        let (_, sheet) = try createTestWorkbook()
        
        // Find first cell with "NYC"
        let result = sheet.find { ref, value in
            value == .text("NYC")
        }
        
        #expect(result != nil)
        #expect(result?.reference.column == "C")
        #expect(result?.reference.row == 2) // Alice's row
        #expect(result?.value == .text("NYC"))
    }
    
    @Test("Find first with no match")
    func findFirstNoMatch() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let result = sheet.find { _, value in
            value == .text("NotFound")
        }
        
        #expect(result == nil)
    }
    
    @Test("Find all matching cells")
    func findAll() throws {
        let (_, sheet) = try createTestWorkbook()
        
        // Find all cells with "NYC"
        let results = sheet.findAll { _, value in
            value == .text("NYC")
        }
        
        #expect(results.count == 2)
        #expect(results[0].reference.row == 2) // Alice
        #expect(results[1].reference.row == 4) // Charlie
    }
    
    @Test("Find all numbers greater than 30")
    func findAllNumbersGreaterThan30() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let results = sheet.findAll { _, value in
            if case .number(let n) = value {
                return n > 30
            }
            return false
        }
        
        #expect(results.count == 1)
        #expect(results[0].value == .number(35)) // Charlie's age
    }
    
    @Test("Complex query: filter rows then extract values")
    func complexQuery() throws {
        let (_, sheet) = try createTestWorkbook()
        
        // Find all people in NYC and get their names
        let nycPeople = sheet.rows { cells in
            cells.contains { $0.reference.column == "C" && $0.value == .text("NYC") }
        }.map { row in
            row.first { $0.reference.column == "A" }?.value
        }
        
        #expect(nycPeople.count == 2)
        #expect(nycPeople[0] == .text("Alice"))
        #expect(nycPeople[1] == .text("Charlie"))
    }
    
    @Test("Range with gaps in data")
    func rangeWithGaps() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sparse")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeNumber(1, to: "A1")
            sheet.writeNumber(2, to: "C1")
            sheet.writeNumber(3, to: "A2")
            // B1, B2, C2 are empty
        }
        
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("sparse-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        let workbook = try Workbook.open(url: fileURL)
        let sheet = try workbook.sheet(at: 0)!
        
        let range = sheet.range("A1:C2")
        
        // Should have 6 cells, some nil
        #expect(range.count == 6)
        #expect(range[0].value == .number(1))
        #expect(range[1].value == nil) // B1
        #expect(range[2].value == .number(2))
        #expect(range[3].value == .number(3))
        #expect(range[4].value == nil) // B2
        #expect(range[5].value == nil) // C2
    }
    
    @Test("Empty column returns empty array")
    func emptyColumn() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let colZ = sheet.column("Z")
        #expect(colZ.isEmpty)
    }
    
    @Test("Row filtering with no matches")
    func rowFilteringNoMatches() throws {
        let (_, sheet) = try createTestWorkbook()
        
        let impossibleRows = sheet.rows { cells in
            cells.contains { $0.value == .number(999) }
        }
        
        #expect(impossibleRows.isEmpty)
    }
}
