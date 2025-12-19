import Foundation
import Testing
@testable import Cuneiform

@Suite struct HyperlinksReadTests {
    @Test func parseSingleHyperlink() throws {
        let xml = """
        <?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"
                   xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
          <sheetData/>
          <hyperlinks>
            <hyperlink ref=\"A1\" r:id=\"rId1\" display=\"Click\" tooltip=\"Go to site\"/>
          </hyperlinks>
        </worksheet>
        """.data(using: .utf8)!

        let ws = try WorksheetParser.parse(data: xml)
        #expect(ws.hyperlinks.count == 1)
        let hl = ws.hyperlinks[0]
        let a1: CellReference = "A1"
        #expect(hl.ref == a1)
        #expect(hl.relationshipId == "rId1")
        #expect(hl.display == "Click")
        #expect(hl.tooltip == "Go to site")
        #expect(hl.location == nil)
    }

    @Test func parseInternalHyperlink() throws {
        let xml = """
        <?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData/>
          <hyperlinks>
            <hyperlink ref=\"B2\" location=\"Sheet2!A1\" display=\"Jump\" />
          </hyperlinks>
        </worksheet>
        """.data(using: .utf8)!

        let ws = try WorksheetParser.parse(data: xml)
        #expect(ws.hyperlinks.count == 1)
        let hl = ws.hyperlinks[0]
        let b2: CellReference = "B2"
        #expect(hl.ref == b2)
        #expect(hl.relationshipId == nil)
        #expect(hl.display == "Jump")
        #expect(hl.location == "Sheet2!A1")
    }
}
