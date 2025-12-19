import Foundation
import Testing

@testable import Cuneiform

@Suite("Named Ranges Read")
struct NamedRangesReadTests {
    @Test("Read definedNames from workbook.xml")
    func readDefinedNames() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("named_read_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        writer.addNamedRange(name: "MyRange", refersTo: "Sheet1!$A$1:$A$5")
        let s1 = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: s1) { sheet in
            sheet.writeText("x", to: "A1")
        }
        try writer.save(to: tempFile)

        let wb = try Workbook.open(url: tempFile)
        let names = wb.definedNamesList
        #expect(names.contains(DefinedName(name: "MyRange", refersTo: "Sheet1!$A$1:$A$5")))
    }
}
