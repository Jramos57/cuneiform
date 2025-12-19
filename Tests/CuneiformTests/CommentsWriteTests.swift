import Foundation
import Testing
@testable import Cuneiform

@Suite struct CommentsWriteTests {
    @Test func writesCommentsPartAndRelationship() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.addComment(at: "B2", text: "Hello there", author: "Vek")
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        var package = try OPCPackage.open(url: tempURL)
        #expect(package.partExists(PartPath("/xl/comments1.xml")))

        let commentsData = try package.readPart(PartPath("/xl/comments1.xml"))
        let parsed = try CommentsParser.parse(data: commentsData)
        #expect(parsed.authors == ["Vek"])
        #expect(parsed.comments.count == 1)
        let comment = parsed.comments[0]
        let b2: CellReference = "B2"
        #expect(comment.ref == b2)
        #expect(comment.author == "Vek")
        #expect(comment.text == "Hello there")

        let sheetRels = try package.relationships(for: PartPath("/xl/worksheets/sheet1.xml"))
        #expect(sheetRels[.comments].count == 1)
        let rel = sheetRels[.comments][0]
        #expect(rel.target == "../comments1.xml")
    }
}
