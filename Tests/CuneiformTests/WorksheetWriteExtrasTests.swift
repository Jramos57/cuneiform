import Foundation
import Testing

@testable import Cuneiform

@Suite("Worksheet Write Extras")
struct WorksheetWriteExtrasTests {
    @Test("Write and read merged cells")
    func writeMergedCells() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("merged_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "Merge")
        writer.modifySheet(at: idx) { sheet in
            sheet.writeText("A", to: CellReference(column: "A", row: 1))
            sheet.writeText("B", to: CellReference(column: "B", row: 1))
            sheet.mergeCells("A1:B1")
        }
        try writer.save(to: tempFile)

        let wb = try Workbook.open(url: tempFile)
        let maybeSheet = try wb.sheet(named: "Merge")
        let sheet = try #require(maybeSheet)
        #expect(sheet.mergedCells.contains("A1:B1"))
    }

    @Test("Write list data validation (smoke)")
    func writeDataValidationList() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("validation_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "Validate")
        writer.modifySheet(at: idx) { sheet in
            sheet.writeText("Select:", to: CellReference(column: "A", row: 1))
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .list,
                allowBlank: true,
                sqref: "B2:B10",
                formula1: "\"Yes,No,Maybe\""
            ))
        }
        try writer.save(to: tempFile)

        // Inspect the worksheet XML to verify presence of dataValidations
        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let info = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[info.relationshipId] else {
            #expect(false)
            return
        }
        let path = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try String(data: pkg.readPart(path), encoding: .utf8)!
        #expect(sheetXML.contains("<dataValidations"))
        #expect(sheetXML.contains("type=\"list\""))
        #expect(sheetXML.contains("sqref=\"B2:B10\""))
    }
    
    @Test("Write and read row heights and column widths")
    func writeRowAndColumnDimensions() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dimensions_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "Dimensions")
        writer.modifySheet(at: idx) { sheet in
            // Add some data
            sheet.writeText("Header 1", to: CellReference(column: "A", row: 1))
            sheet.writeText("Header 2", to: CellReference(column: "B", row: 1))
            sheet.writeText("Data", to: CellReference(column: "A", row: 2))
            
            // Set row heights
            sheet.setRowHeight(row: 1, height: 30.0)  // Header row taller
            sheet.setRowHeight(row: 2, height: 20.0)
            sheet.setRowHeight(row: 3, height: 15.0, hidden: true)  // Hidden row
            
            // Set column widths
            sheet.setColumnWidth(min: 1, max: 1, width: 20.0)  // Column A
            sheet.setColumnWidth(column: 2, width: 15.0)  // Column B
            sheet.setColumnWidth(column: 3, width: 10.0, hidden: true)  // Hidden column C
        }
        try writer.save(to: tempFile)

        // Read back and verify
        let wb = try Workbook.open(url: tempFile)
        let maybeSheet = try wb.sheet(named: "Dimensions")
        let sheet = try #require(maybeSheet)
        
        // Verify we can access the data (basic sanity check)
        #expect(sheet.cell(at: "A1")?.asString == "Header 1")
        
        // Verify dimensions are present in raw data
        let rawData = sheet.getRawData()
        
        // Check rows
        #expect(rawData.rows.count >= 2)
        let row1 = rawData.rows.first { $0.index == 1 }
        #expect(row1 != nil)
        #expect(row1?.height == 30.0)
        #expect(row1?.customHeight == true)
        #expect(row1?.hidden == false)
        
        let row3 = rawData.rows.first { $0.index == 3 }
        if row3 != nil {
            #expect(row3?.hidden == true)
        }
        
        // Check columns
        #expect(rawData.columns.count >= 3)
        let colA = rawData.columns.first { $0.min == 1 && $0.max == 1 }
        #expect(colA != nil)
        #expect(colA?.width == 20.0)
        #expect(colA?.customWidth == true)
        #expect(colA?.hidden == false)
        
        let colC = rawData.columns.first { $0.min == 3 && $0.max == 3 }
        #expect(colC != nil)
        #expect(colC?.hidden == true)
    }
}
