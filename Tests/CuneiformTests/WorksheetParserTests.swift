import Testing
@testable import Cuneiform
import Foundation

@Suite struct WorksheetParserTests {
    @Test func parseCellReference() {
        // Using string literals (ExpressibleByStringLiteral)
        let a1: CellReference = "A1"
        #expect(a1.column == "A")
        #expect(a1.row == 1)
        #expect(a1.columnIndex == 0)

        let z1: CellReference = "Z1"
        #expect(z1.columnIndex == 25)

        let aa1: CellReference = "AA1"
        #expect(aa1.columnIndex == 26)

        let az1: CellReference = "AZ1"
        #expect(az1.columnIndex == 51)
    }

    @Test func parseSharedStringCell() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\">
              <c r=\"A1\" t=\"s\"><v>0</v></c>
            </row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        let cell = ws.cell(at: "A1")
        if case .sharedString(index: 0)? = cell?.value {
            #expect(true)
        } else {
            #expect(false)
        }
    }

    @Test func parseNumberCell() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\"><v>42.5</v></c></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        let cell = ws.cell(at: "A1")
        if case .number(let n)? = cell?.value {
            #expect(n == 42.5)
        } else { #expect(false) }
    }

    @Test func parseBooleanCell() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\" t=\"b\"><v>1</v></c></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        let cell = ws.cell(at: "A1")
        if case .boolean(true)? = cell?.value { #expect(true) } else { #expect(false) }
    }

    @Test func parseInlineStringCell() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\" t=\"str\"><v>hello</v></c></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        let cell = ws.cell(at: "A1")
        if case .inlineString(let s)? = cell?.value { #expect(s == "hello") } else { #expect(false) }
    }

    @Test func parseErrorCell() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        let cell = ws.cell(at: "A1")
        if case .error(let e)? = cell?.value { #expect(e == "#DIV/0!") } else { #expect(false) }
    }

    @Test func parseEmptyCell() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\"/></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        #expect(ws.cell(at: "A1")?.value == .empty)
    }

    @Test func parseMergedCells() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\"/></row>
          </sheetData>
          <mergeCells count=\"1\"><mergeCell ref=\"A5:C5\"/></mergeCells>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        #expect(ws.mergedCells == ["A5:C5"])
    }

    @Test func cellLookup() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"A1\"><v>1</v></c></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        #expect(ws.cell(at: "A1") != nil)
        #expect(ws.cell(at: "B1") == nil)
    }

    @Test func sparseData() throws {
        let xml = """
        <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <sheetData>
            <row r=\"1\"><c r=\"C1\"><v>3.14</v></c></row>
            <row r=\"5\"><c r=\"A5\"/></row>
          </sheetData>
        </worksheet>
        """.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: xml)
        #expect(ws.rows.count == 2)
        #expect(ws.cell(at: "C1") != nil)
        #expect(ws.cell(at: "A5") != nil)
        #expect(ws.cell(at: "B2") == nil)
    }
}
