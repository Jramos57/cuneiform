import Foundation
import Testing

@testable import Cuneiform

@Suite("Data Validation Variants")
struct DataValidationVariantTests {
    @Test("Numeric between validation emits operator and formulas")
    func numericBetween() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dv_between_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "Validate")
        writer.modifySheet(at: idx) { sheet in
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .whole,
                allowBlank: false,
                sqref: "C2:C100",
                formula1: "1",
                formula2: "10",
                op: "between"
            ))
        }
        try writer.save(to: tempFile)

        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let sheetInfo = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[sheetInfo.relationshipId] else {
            #expect(false)
            return
        }
        let path = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try String(data: pkg.readPart(path), encoding: .utf8)!
        #expect(sheetXML.contains("<dataValidations"))
        #expect(sheetXML.contains("type=\"whole\""))
        #expect(sheetXML.contains("operator=\"between\""))
        #expect(sheetXML.contains("<formula1>1</formula1>"))
        #expect(sheetXML.contains("<formula2>10</formula2>"))
    }
}
