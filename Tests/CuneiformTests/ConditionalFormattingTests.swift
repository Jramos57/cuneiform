import Testing
import Foundation
@testable import Cuneiform

@Suite struct ConditionalFormattingTests {
    @Test func parseCellIsRule() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <conditionalFormatting sqref="A1">
            <cfRule type="cellIs" operator="greaterThan" priority="1">
              <formula>5</formula>
            </cfRule>
          </conditionalFormatting>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.conditionalFormats.count == 1)
        let cf = ws.conditionalFormats[0]
        #expect(cf.range == "A1")
        #expect(cf.rules.count == 1)
        if case let .cellIs(op, f1, f2) = cf.rules[0].type {
            #expect(op == .greaterThan)
            #expect(f1 == "5")
            #expect(f2 == nil)
        } else {
            Issue.record("Expected cellIs rule")
        }
    }

    @Test func parseDataBarRule() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <conditionalFormatting sqref="B1:B10">
            <cfRule type="dataBar" priority="1">
              <dataBar showValue="0">
                <cfvo type="min"/>
                <cfvo type="max"/>
                <color rgb="FF63BE7B"/>
              </dataBar>
            </cfRule>
          </conditionalFormatting>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.conditionalFormats.count == 1)
        let rule = ws.conditionalFormats[0].rules[0]
        if case let .dataBar(db) = rule.type {
            #expect(db.min.type == .min)
            #expect(db.max.type == .max)
            #expect(db.color == "FF63BE7B")
            #expect(db.showValue == false)
        } else {
            Issue.record("Expected dataBar rule")
        }
    }

    @Test func writeCellIsConditionalFormatting() throws {
        var builder = WorksheetBuilder()
        let rule = WorksheetData.ConditionalRule(
            type: .cellIs(op: .between, formula1: "10", formula2: "20"),
            priority: nil,
            dxfId: nil,
            stopIfTrue: false
        )
        builder.addConditionalFormat(WorksheetData.ConditionalFormat(range: "A1:A10", rules: [rule]))
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        #expect(xml.contains("<conditionalFormatting sqref=\"A1:A10\">"))
        #expect(xml.contains("type=\"cellIs\""))
        #expect(xml.contains("operator=\"between\""))
        #expect(xml.contains("<formula>10</formula>"))
        #expect(xml.contains("<formula>20</formula>"))
    }

    @Test func writeColorScale() throws {
        var builder = WorksheetBuilder()
        let cfvo1 = WorksheetData.CFValueObject(type: .min, value: nil)
        let cfvo2 = WorksheetData.CFValueObject(type: .percentile, value: "50")
        let cfvo3 = WorksheetData.CFValueObject(type: .max, value: nil)
        let rule = WorksheetData.ConditionalRule(
            type: .colorScale(WorksheetData.ConditionalRule.ColorScale(cfvos: [cfvo1, cfvo2, cfvo3], colors: ["FF0000", "FFFF00", "00FF00"])),
            priority: 1,
            dxfId: nil,
            stopIfTrue: false
        )
        builder.addConditionalFormat(WorksheetData.ConditionalFormat(range: "C1:C10", rules: [rule]))
        let xml = String(data: builder.build(), encoding: .utf8)!
        #expect(xml.contains("colorScale"))
        #expect(xml.contains("cfvo type=\"percentile\" val=\"50\""))
        #expect(xml.contains("<color rgb=\"FFFF00\"/>"))
    }

    @Test func roundTripConditionalFormatting() throws {
        let tempFile = FileManager.default.temporaryDirectory.appendingPathComponent("cf_roundtrip.xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: 0) { sheet in
            sheet.writeNumber(5, to: "A1")
            let rule = WorksheetData.ConditionalRule(
                type: .cellIs(op: .greaterThan, formula1: "3", formula2: nil),
                priority: 1,
                dxfId: nil,
                stopIfTrue: false
            )
            sheet.addConditionalFormat(range: "A1", rule: rule)
        }

        try writer.save(to: tempFile)

        let workbook = try Workbook.open(url: tempFile)
        guard let sheet = try workbook.sheet(named: "Sheet1") else {
            Issue.record("Sheet1 not found after round-trip")
            return
        }

        #expect(sheet.conditionalFormats.count == 1)
        if case let .cellIs(op, f1, _) = sheet.conditionalFormats[0].rules.first?.type {
            #expect(op == .greaterThan)
            #expect(f1 == "3")
        } else {
            Issue.record("Expected cellIs rule after round-trip")
        }
    }
}
