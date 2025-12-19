import Foundation
import Testing
@testable import Cuneiform

@Suite("Data Validation Helpers")
struct DataValidationHelpersTests {
    @Test("Filter validations by range and cell")
    func filterValidations() throws {
        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1"><v>1</v></c>
            </row>
          </sheetData>
          <dataValidations count="2">
            <dataValidation type="whole" allowBlank="1" operator="between" sqref="A1:A3 C5 D7:F9">
              <formula1>1</formula1>
              <formula2>10</formula2>
            </dataValidation>
            <dataValidation type="list" allowBlank="0" sqref="B2">
              <formula1>"Yes,No"</formula1>
            </dataValidation>
          </dataValidations>
        </worksheet>
        """.data(using: .utf8)!

        let ws = try WorksheetParser.parse(data: xml)
        let sheet = Sheet(data: ws, sharedStrings: .empty, styles: .empty)

        // by range intersection (captures A1:A3 and single B2)
        let r1 = sheet.validations(for: "A2:B2")
        #expect(r1.count == 2)
        let kinds1 = Set(r1.map(\.type))
        #expect(kinds1.contains(.whole))
        #expect(kinds1.contains(.list))

        let r2 = sheet.validations(for: "E8:G10")
        #expect(r2.count == 1)
        #expect(r2.first?.type == .whole)

        let r3 = sheet.validations(for: "Z1:Z5")
        #expect(r3.isEmpty)

        // by single cell
        let b2: CellReference = "B2"
        let c1 = sheet.validations(at: b2)
        #expect(c1.count == 1)
        #expect(c1.first?.type == .list)

        let c5: CellReference = "C5"
        let c2 = sheet.validations(at: c5)
        #expect(c2.count == 1)
        #expect(c2.first?.type == .whole)

        let b3: CellReference = "B3"
        let c3 = sheet.validations(at: b3)
        #expect(c3.isEmpty)
    }
}
