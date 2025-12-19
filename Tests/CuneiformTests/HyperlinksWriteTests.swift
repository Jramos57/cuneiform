import Foundation
import Testing

@testable import Cuneiform

@Suite("Hyperlinks Write")
struct HyperlinksWriteTests {
    @Test("WorksheetWriter: external + internal hyperlinks emission")
    func writeHyperlinks() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("hyperlinks_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "Links")
        writer.modifySheet(at: idx) { sheet in
            sheet.writeText("Go to Web", to: "B1")
            sheet.addHyperlinkExternal(at: "B2", url: "https://example.com", display: "Example", tooltip: "Open example.com")

            sheet.writeText("Go to Cell", to: "C1")
            sheet.addHyperlinkInternal(at: "C3", location: "Links!A1", display: "Top", tooltip: "Jump to A1")
        }

        try writer.save(to: tempFile)

        // Inspect worksheet XML
        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let info = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[info.relationshipId] else {
            #expect(false)
            return
        }
        let sheetPath = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try pkg.readPartAsString(sheetPath)

        // Verify hyperlinks section in sheet XML
        #expect(sheetXML.contains("<hyperlinks>"))
        #expect(sheetXML.contains("ref=\"B2\""))
        #expect(sheetXML.contains("display=\"Example\""))
        #expect(sheetXML.contains("tooltip=\"Open example.com\""))
        #expect(sheetXML.contains("r:id=\"")) // external link should have relationship id

        #expect(sheetXML.contains("ref=\"C3\""))
        #expect(sheetXML.contains("location=\"Links!A1\""))

        // Verify worksheet relationships for external hyperlink
        let wsRels = try pkg.relationships(for: sheetPath)
        let hyperlinkRels = wsRels[.hyperlink]
        #expect(hyperlinkRels.count == 1)
        let hRel = try #require(hyperlinkRels.first)
        #expect(hRel.isExternal)
        #expect(hRel.target == "https://example.com")

        // Relationships part must exist
        #expect(pkg.partExists(sheetPath.relationshipsPath))
    }
}
