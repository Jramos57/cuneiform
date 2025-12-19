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

    @Test func writesVMLDrawingForComments() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.addComment(at: "A1", text: "First note", author: "Author A")
            sheet.addComment(at: "C3", text: "Second note", author: "Author B")
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        var package = try OPCPackage.open(url: tempURL)
        
        // Verify VML drawing part exists
        #expect(package.partExists(PartPath("/xl/drawings/vmlDrawing1.vml")))
        
        // Verify VML content is valid XML with shapes
        let vmlData = try package.readPart(PartPath("/xl/drawings/vmlDrawing1.vml"))
        let vmlString = String(data: vmlData, encoding: .utf8)!
        #expect(vmlString.contains("<xml"))
        #expect(vmlString.contains("v:shapetype"))
        #expect(vmlString.contains("v:shape"))
        #expect(vmlString.contains("ClientData"))
        
        // Verify VML drawing relationship exists
        let sheetRels = try package.relationships(for: PartPath("/xl/worksheets/sheet1.xml"))
        let vmlRels = sheetRels[.vmlDrawing]
        #expect(vmlRels.count == 1)
        #expect(vmlRels[0].target == "../drawings/vmlDrawing1.vml")
        
        // Verify worksheet XML contains legacyDrawing element with correct r:id
        let wsData = try package.readPart(PartPath("/xl/worksheets/sheet1.xml"))
        let wsString = String(data: wsData, encoding: .utf8)!
        #expect(wsString.contains("<legacyDrawing"))
        #expect(wsString.contains("r:id"))
    }
}

