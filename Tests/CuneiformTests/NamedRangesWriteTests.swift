import Foundation
import Testing

@testable import Cuneiform

@Suite("Named Ranges Write")
struct NamedRangesWriteTests {
    @Test("Write definedNames to workbook.xml")
    func writeDefinedNames() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("namedranges_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Data")
        writer.addNamedRange(name: "MyRange", refersTo: "Sheet1!$A$1:$B$10")
        try writer.save(to: tempFile)

        // Inspect workbook.xml for definedNames
        var pkg = try OPCPackage.open(url: tempFile)
        let xml = try String(data: pkg.readPart(.workbook), encoding: .utf8)!
        #expect(xml.contains("<definedNames>"))
        #expect(xml.contains("name=\"MyRange\""))
        #expect(xml.contains("Sheet1!$A$1:$B$10"))
    }
}
