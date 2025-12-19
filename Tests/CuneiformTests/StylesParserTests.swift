import Testing
@testable import Cuneiform
import Foundation

@Suite struct StylesParserTests {
    @Test func parseCustomNumberFormats() throws {
        let xml = """
        <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <numFmts count=\"2\">
            <numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/>
            <numFmt numFmtId=\"200\" formatCode=\"#,##0.00\"/>
          </numFmts>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        #expect(styles.numberFormats[164] == "yyyy-mm-dd")
        #expect(styles.numberFormats[200] == "#,##0.00")
    }

    @Test func parseCellFormats() throws {
        let xml = """
        <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <cellXfs count=\"3\">
            <xf numFmtId=\"0\"/>
            <xf numFmtId=\"14\"/>
            <xf numFmtId=\"164\"/>
          </cellXfs>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        #expect(styles.cellFormats.count == 3)
        #expect(styles.cellFormats[0].numFmtId == 0)
        #expect(styles.cellFormats[1].numFmtId == 14)
        #expect(styles.cellFormats[2].numFmtId == 164)
    }

    @Test func builtInDateFormats() throws {
        let xml = """
        <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <cellXfs count=\"1\"><xf numFmtId=\"14\"/></cellXfs>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        #expect(styles.isDateFormat(styleIndex: 0))
    }

    @Test func customDateFormat() throws {
        let xml = """
        <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <numFmts count=\"1\">
            <numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/>
          </numFmts>
          <cellXfs count=\"1\"><xf numFmtId=\"164\"/></cellXfs>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        #expect(styles.isDateFormat(styleIndex: 0))
    }

    @Test func numberFormatNotDate() throws {
        let xml = """
        <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <numFmts count=\"1\">
            <numFmt numFmtId=\"200\" formatCode=\"#,##0.00\"/>
          </numFmts>
          <cellXfs count=\"1\"><xf numFmtId=\"200\"/></cellXfs>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        #expect(styles.isDateFormat(styleIndex: 0) == false)
    }

    @Test func missingStylesFile() {
        let styles = StylesInfo.empty
        #expect(styles.cellFormats.count == 1)  // Default format at index 0
        #expect(styles.numberFormats.isEmpty)
        #expect(styles.fonts.count == 1)  // Default font
        #expect(styles.fills.count == 2)  // Default fills
        #expect(styles.borders.count == 1)  // Default border
    }
}
