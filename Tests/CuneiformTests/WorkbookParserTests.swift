import Testing
@testable import Cuneiform
import Foundation

@Suite struct WorkbookParserTests {
    @Test func parseSheets() throws {
        let xml = """
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"
                  xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
          <sheets>
            <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>
            <sheet name=\"Data\" sheetId=\"2\" r:id=\"rId2\"/>
          </sheets>
        </workbook>
        """.data(using: .utf8)!

        let info = try WorkbookParser.parse(data: xml)
        #expect(info.sheets.count == 2)
        #expect(info.sheets.map(\.name) == ["Sheet1", "Data"])
        #expect(info.sheets.map(\.sheetId) == [1, 2])
        #expect(info.sheets.map(\.relationshipId) == ["rId1", "rId2"])
    }

    @Test func parseSheetStates() throws {
        let xml = """
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"
                  xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
          <sheets>
            <sheet name=\"A\" sheetId=\"1\" r:id=\"rId1\"/>
            <sheet name=\"B\" sheetId=\"2\" state=\"hidden\" r:id=\"rId2\"/>
            <sheet name=\"C\" sheetId=\"3\" state=\"veryHidden\" r:id=\"rId3\"/>
          </sheets>
        </workbook>
        """.data(using: .utf8)!

        let info = try WorkbookParser.parse(data: xml)
        #expect(info.sheets.map(\.state) == [.visible, .hidden, .veryHidden])
    }

    @Test func sheetByName() throws {
        let xml = """
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"
                  xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
          <sheets>
            <sheet name=\"X\" sheetId=\"10\" r:id=\"r1\"/>
            <sheet name=\"Y\" sheetId=\"20\" r:id=\"r2\"/>
          </sheets>
        </workbook>
        """.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: xml)
        #expect(info.sheet(named: "X")?.sheetId == 10)
        #expect(info.sheet(named: "Z") == nil)
    }

    @Test func missingNameThrows() {
        let xml = """
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"
                  xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
          <sheets>
            <sheet sheetId=\"1\" r:id=\"rId1\"/>
          </sheets>
        </workbook>
        """.data(using: .utf8)!

        #expect(throws: CuneiformError.self) {
            _ = try WorkbookParser.parse(data: xml)
        }
    }

    @Test func emptyWorkbook() throws {
        let xml = """
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheets/>
        </workbook>
        """.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: xml)
        #expect(info.sheets.isEmpty)
    }
}
