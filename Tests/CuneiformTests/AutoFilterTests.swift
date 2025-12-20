import Testing
import Foundation
@testable import Cuneiform

@Suite struct AutoFilterTests {
    // MARK: - Parser Tests

    @Test func parseSimpleAutoFilter() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <autoFilter ref="A1:D100"/>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.autoFilter != nil)
        #expect(ws.autoFilter?.ref == "A1:D100")
        #expect(ws.autoFilter?.columnFilters.isEmpty == true)
    }

    @Test func parseAutoFilterWithDiscreteValues() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <autoFilter ref="A1:C50">
            <filterColumn colId="0">
              <filters>
                <filter val="Apple"/>
                <filter val="Orange"/>
                <filter val="Banana"/>
              </filters>
            </filterColumn>
          </autoFilter>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.autoFilter != nil)
        #expect(ws.autoFilter?.ref == "A1:C50")
        #expect(ws.autoFilter?.columnFilters.count == 1)

        let filter = ws.autoFilter!.columnFilters[0]
        #expect(filter.colId == 0)
        if case let .values(vals) = filter.criterion {
            #expect(vals.count == 3)
            #expect(vals.contains("Apple"))
            #expect(vals.contains("Orange"))
            #expect(vals.contains("Banana"))
        } else {
            Issue.record("Expected discrete values filter")
        }
    }

    @Test func parseAutoFilterMultipleColumns() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <autoFilter ref="A1:E100">
            <filterColumn colId="1">
              <filters>
                <filter val="Active"/>
              </filters>
            </filterColumn>
            <filterColumn colId="3">
              <filters>
                <filter val="High"/>
                <filter val="Critical"/>
              </filters>
            </filterColumn>
          </autoFilter>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.autoFilter?.columnFilters.count == 2)
        #expect(ws.autoFilter?.columnFilters[0].colId == 1)
        #expect(ws.autoFilter?.columnFilters[1].colId == 3)
    }

    // MARK: - Write Tests

    @Test func writeAutoFilterRange() throws {
        var builder = WorksheetBuilder()
        builder.addCell(at: CellReference(column: "A", row: 1), value: .text("Name"))
        builder.addCell(at: CellReference(column: "B", row: 1), value: .text("Value"))
        builder.setAutoFilter(range: "A1:B10")

        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        #expect(xml.contains("<autoFilter ref=\"A1:B10\"/>"))
    }

    // MARK: - Round-trip Tests

    @Test func roundTripAutoFilter() throws {
        let tempFile = FileManager.default.temporaryDirectory.appendingPathComponent("autofilter_roundtrip.xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Data")
        writer.modifySheet(at: 0) { sheet in
            // Header row
            sheet.writeText("Name", to: CellReference(column: "A", row: 1))
            sheet.writeText("Value", to: CellReference(column: "B", row: 1))
            sheet.writeText("Category", to: CellReference(column: "C", row: 1))

            // Data rows
            for i in 2...20 {
                sheet.writeText("Item \(i)", to: CellReference(column: "A", row: i))
                sheet.writeNumber(Double(i * 10), to: CellReference(column: "B", row: i))
                sheet.writeText(i % 2 == 0 ? "Even" : "Odd", to: CellReference(column: "C", row: i))
            }

            // Set autofilter
            sheet.setAutoFilter(range: "A1:C20")
        }

        try writer.save(to: tempFile)

        let workbook = try Workbook.open(url: tempFile)
        guard let sheet = try workbook.sheet(named: "Data") else {
            Issue.record("Data sheet not found after round-trip")
            return
        }

        #expect(sheet.autoFilter != nil)
        #expect(sheet.autoFilter?.ref == "A1:C20")
    }

    // MARK: - API Tests

    @Test func sheetAutoFilterProperty() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <autoFilter ref="B2:F50"/>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.autoFilter?.ref == "B2:F50")
    }
}
