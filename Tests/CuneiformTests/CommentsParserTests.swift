import Foundation
import Testing
@testable import Cuneiform

@Suite struct CommentsParserTests {
    @Test func parseSingleComment() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <authors>
            <author>Vek</author>
          </authors>
          <commentList>
            <comment ref="B2" authorId="0">
              <text>
                <r><t>Hello there</t></r>
              </text>
            </comment>
          </commentList>
        </comments>
        """.data(using: .utf8)!

        let parsed = try CommentsParser.parse(data: xml)
        #expect(parsed.authors == ["Vek"])
        #expect(parsed.comments.count == 1)
        let c = parsed.comments[0]
        let b2: CellReference = "B2"
        #expect(c.ref == b2)
        #expect(c.author == "Vek")
        #expect(c.text == "Hello there")
    }
}
