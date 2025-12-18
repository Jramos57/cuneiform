import Testing
@testable import Cuneiform
import Foundation

@Suite struct SharedStringsTests {
    @Test func parseSimpleStrings() throws {
        let xml = """
        <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <si><t>A</t></si>
          <si><t>B</t></si>
        </sst>
        """.data(using: .utf8)!

        let ss = try SharedStringsParser.parse(data: xml)
        #expect(ss.count == 2)
        #expect(ss[0] == "A")
        #expect(ss[1] == "B")
    }

    @Test func parseRichText() throws {
        let xml = """
        <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <si>
            <r><t>Hello </t></r>
            <r><rPr><b/></rPr><t>World</t></r>
          </si>
        </sst>
        """.data(using: .utf8)!

        let ss = try SharedStringsParser.parse(data: xml)
        #expect(ss.count == 1)
        #expect(ss[0] == "Hello World")
    }

    @Test func parsePreservedWhitespace() throws {
        let xml = """
        <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <si><t xml:space=\"preserve\">  A  </t></si>
        </sst>
        """.data(using: .utf8)!

        let ss = try SharedStringsParser.parse(data: xml)
        #expect(ss.count == 1)
        #expect(ss[0] == "  A  ")
    }

    @Test func parseEmptyString() throws {
        let xml = """
        <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <si><t/></si>
          <si></si>
        </sst>
        """.data(using: .utf8)!

        let ss = try SharedStringsParser.parse(data: xml)
        #expect(ss.count == 2)
        #expect(ss[0] == "")
        #expect(ss[1] == "")
    }

    @Test func subscriptOutOfBounds() throws {
        let xml = """
        <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <si><t>A</t></si>
        </sst>
        """.data(using: .utf8)!
        let ss = try SharedStringsParser.parse(data: xml)

        #expect(ss[-1] == nil)
        #expect(ss[99] == nil)
    }
}
