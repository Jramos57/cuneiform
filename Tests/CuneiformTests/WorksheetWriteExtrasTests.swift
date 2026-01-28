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
        // Test using low-level builder API
        var builder = WorksheetBuilder()
        
        // Add some data
        builder.addCell(at: CellReference(column: "A", row: 1), value: .text("Header 1"))
        builder.addCell(at: CellReference(column: "B", row: 1), value: .text("Header 2"))
        builder.addCell(at: CellReference(column: "A", row: 2), value: .text("Data"))
        
        // Set row heights
        builder.setRowHeight(row: 1, height: 30.0)  // Header row taller
        builder.setRowHeight(row: 2, height: 20.0)
        builder.setRowHeight(row: 3, height: 15.0, hidden: true)  // Hidden row
        
        // Set column widths
        builder.setColumnWidth(min: 1, max: 1, width: 20.0)  // Column A
        builder.setColumnWidth(column: 2, width: 15.0)  // Column B
        builder.setColumnWidth(column: 3, width: 10.0, hidden: true)  // Hidden column C
        
        // Build and parse XML
        let xml = builder.build()
        let xmlString = String(data: xml, encoding: .utf8)!
        
        // Verify XML contains dimension elements
        #expect(xmlString.contains("<cols>"))
        #expect(xmlString.contains("ht=\"30"))  // Allow for decimals
        #expect(xmlString.contains("ht=\"20"))  // Allow for decimals
        #expect(xmlString.contains("customHeight=\"1\""))
        #expect(xmlString.contains("width=\"20"))  // Allow for decimals
        #expect(xmlString.contains("width=\"15"))  // Allow for decimals
        #expect(xmlString.contains("hidden=\"1\""))
        
        // Parse the XML back to verify round-trip
        let data = try WorksheetParser.parse(data: xml)
        
        // Verify row dimensions
        let row1 = data.rows.first { $0.index == 1 }
        #expect(row1 != nil)
        #expect(row1?.height == 30.0)
        #expect(row1?.customHeight == true)
        #expect(row1?.hidden == false)
        
        let row3 = data.rows.first { $0.index == 3 }
        // Row 3 won't exist in parsed data because it has no cells
        // (empty rows with only dimension attributes are not output)
        #expect(row3 == nil)
        
        // Verify column dimensions
        #expect(data.columns.count >= 3)
        
        let colA = data.columns.first { $0.min == 1 && $0.max == 1 }
        #expect(colA != nil)
        #expect(colA?.width == 20.0)
        #expect(colA?.customWidth == true)
        #expect(colA?.hidden == false)
        
        let colC = data.columns.first { $0.min == 3 && $0.max == 3 }
        #expect(colC != nil)
        #expect(colC?.width == 10.0)
        #expect(colC?.hidden == true)
    }
}
