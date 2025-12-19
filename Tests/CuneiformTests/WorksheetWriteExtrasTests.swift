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
}
