import Foundation
import Testing

@testable import Cuneiform

@Suite("Data Validation Read")
struct DataValidationReadTests {
    @Test("Read date between validation from worksheet")
    func readDateBetween() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dv_read_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "Validate")
        writer.modifySheet(at: idx) { sheet in
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .date,
                allowBlank: true,
                sqref: "B2:B10",
                formula1: "DATE(2025,1,1)",
                formula2: "DATE(2025,12,31)",
                op: "between"
            ))
        }
        try writer.save(to: tempFile)

        let wb = try Workbook.open(url: tempFile)
        let maybeSheet = try wb.sheet(named: "Validate")
        let sheet = try #require(maybeSheet)
        let dvs = sheet.dataValidations
        let maybe = dvs.first { $0.sqref == "B2:B10" }
        let dv = try #require(maybe)
        #expect(dv.type == .date)
        #expect(dv.op == "between")
        #expect(dv.formula1 == "DATE(2025,1,1)")
        #expect(dv.formula2 == "DATE(2025,12,31)")
        #expect(dv.allowBlank == true)
    }
}
