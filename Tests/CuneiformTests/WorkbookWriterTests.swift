import Testing
import Foundation
@testable import Cuneiform

@Suite("WorkbookWriter Tests")
struct WorkbookWriterTests {
    
    @Test("Create empty workbook")
    func createEmptyWorkbook() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        
        let data = try writer.buildData()
        #expect(data.count > 0)
    }
    
    @Test("Write simple cell values")
    func writeSimpleValues() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Data")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Hello", to: "A1")
            sheet.writeNumber(42.5, to: "B1")
            sheet.writeBoolean(true, to: "C1")
        }
        
        let data = try writer.buildData()
        #expect(data.count > 0)
    }
    
    @Test("Write formula")
    func writeFormula() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Calc")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeNumber(10, to: "A1")
            sheet.writeNumber(20, to: "A2")
            sheet.writeFormula("A1+A2", cachedValue: 30, to: "A3")
        }
        
        let data = try writer.buildData()
        #expect(data.count > 0)
    }
    
    @Test("Multiple sheets")
    func multipleSheets() throws {
        var writer = WorkbookWriter()
        let sheet1 = writer.addSheet(named: "First")
        let sheet2 = writer.addSheet(named: "Second")
        
        writer.modifySheet(at: sheet1) { sheet in
            sheet.writeText("Sheet 1", to: "A1")
        }
        
        writer.modifySheet(at: sheet2) { sheet in
            sheet.writeText("Sheet 2", to: "A1")
        }
        
        let data = try writer.buildData()
        #expect(data.count > 0)
    }
    
    @Test("Round-trip: write and read back")
    func roundTrip() throws {
        // Create and write workbook
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "TestSheet")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Test", to: "A1")
            sheet.writeNumber(123.45, to: "B1")
            sheet.writeBoolean(false, to: "C1")
        }
        
        // Save to temp file
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("roundtrip-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        
        // Read back
        let workbook = try Workbook.open(url: fileURL)
        let sheet = try workbook.sheet(at: 0)!
        
        // Verify values
        #expect(sheet.cell(at: "A1") == .text("Test"))
        #expect(sheet.cell(at: "B1") == .number(123.45))
        #expect(sheet.cell(at: "C1") == .boolean(false))
    }
    
    @Test("Round-trip: formulas")
    func roundTripFormulas() throws {
        // Create and write workbook with formulas
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Formulas")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeNumber(5, to: "A1")
            sheet.writeNumber(10, to: "A2")
            sheet.writeFormula("SUM(A1:A2)", cachedValue: 15, to: "A3")
        }
        
        // Save and read back
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("formulas-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        
        let workbook = try Workbook.open(url: fileURL)
        let sheet = try workbook.sheet(at: 0)!
        
        // Verify values and formulas
        #expect(sheet.cell(at: "A1") == .number(5))
        #expect(sheet.cell(at: "A2") == .number(10))
        
        if let formula = sheet.formula(at: "A3") {
            #expect(formula.formula == "SUM(A1:A2)")
        } else {
            Issue.record("Expected formula at A3")
        }
    }
    
    @Test("Round-trip: multiple sheets")
    func roundTripMultipleSheets() throws {
        // Create workbook with multiple sheets
        var writer = WorkbookWriter()
        let sheet1 = writer.addSheet(named: "Alpha")
        let sheet2 = writer.addSheet(named: "Beta")
        
        writer.modifySheet(at: sheet1) { sheet in
            sheet.writeText("First", to: "A1")
        }
        
        writer.modifySheet(at: sheet2) { sheet in
            sheet.writeText("Second", to: "A1")
        }
        
        // Save and read back
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("multisheets-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        
        let workbook = try Workbook.open(url: fileURL)
        
        #expect(workbook.sheets.count == 2)
        #expect(workbook.sheets[0].name == "Alpha")
        #expect(workbook.sheets[1].name == "Beta")
        
        let alpha = try workbook.sheet(at: 0)!
        let beta = try workbook.sheet(at: 1)!
        
        #expect(alpha.cell(at: "A1") == .text("First"))
        #expect(beta.cell(at: "A1") == .text("Second"))
    }
    
    @Test("Write grid of values")
    func writeGrid() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Grid")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            // Write 3x3 grid
            for row in 1...3 {
                for col in 1...3 {
                    let value = row * 10 + col
                    sheet.writeNumber(Double(value), to: CellReference(column: String(UnicodeScalar(64 + col)!), row: row))
                }
            }
        }
        
        // Save and read back
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("grid-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        
        let workbook = try Workbook.open(url: fileURL)
        let sheet = try workbook.sheet(at: 0)!
        
        // Verify grid
        for row in 1...3 {
            for col in 1...3 {
                let ref = CellReference(column: String(UnicodeScalar(64 + col)!), row: row)
                let expected = Double(row * 10 + col)
                #expect(sheet.cell(at: ref) == .number(expected))
            }
        }
    }
    
    @Test("Write large dataset")
    func writeLargeDataset() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Large")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            // Write 100 rows
            for row in 1...100 {
                sheet.writeNumber(Double(row), to: CellReference(column: "A", row: row))
                sheet.writeText("Row \(row)", to: CellReference(column: "B", row: row))
            }
        }
        
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("large-\(UUID()).xlsx")
        defer { try? FileManager.default.removeItem(at: fileURL) }
        
        try writer.save(to: fileURL)
        
        let workbook = try Workbook.open(url: fileURL)
        let sheet = try workbook.sheet(at: 0)!
        
        // Spot check values
        #expect(sheet.cell(at: "A1") == .number(1))
        #expect(sheet.cell(at: "B1") == .text("Row 1"))
        #expect(sheet.cell(at: "A50") == .number(50))
        #expect(sheet.cell(at: "B50") == .text("Row 50"))
        #expect(sheet.cell(at: "A100") == .number(100))
        #expect(sheet.cell(at: "B100") == .text("Row 100"))
    }
}
