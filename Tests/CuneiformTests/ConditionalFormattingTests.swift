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

    @Test func parseIconSetRule() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <conditionalFormatting sqref="D1:D10">
            <cfRule type="iconSet" priority="1">
              <iconSet iconSet="3Arrows" showValue="0" reverse="1">
                <cfvo type="percent" val="0"/>
                <cfvo type="percent" val="33"/>
                <cfvo type="percent" val="67"/>
              </iconSet>
            </cfRule>
          </conditionalFormatting>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.conditionalFormats.count == 1)
        let rule = ws.conditionalFormats[0].rules[0]
        if case let .iconSet(iset) = rule.type {
            #expect(iset.name == "3Arrows")
            #expect(iset.showValue == false)
            #expect(iset.reverse == true)
            #expect(iset.cfvos.count == 3)
            #expect(iset.cfvos[0].type == .percent)
            #expect(iset.cfvos[0].value == "0")
            #expect(iset.cfvos[1].value == "33")
            #expect(iset.cfvos[2].value == "67")
        } else {
            Issue.record("Expected iconSet rule")
        }
    }

    @Test func writeIconSet() throws {
        var builder = WorksheetBuilder()
        let cfvo1 = WorksheetData.CFValueObject(type: .percent, value: "0")
        let cfvo2 = WorksheetData.CFValueObject(type: .percent, value: "33")
        let cfvo3 = WorksheetData.CFValueObject(type: .percent, value: "67")
        let iconSet = WorksheetData.ConditionalRule.IconSet(
            name: "3TrafficLights1",
            cfvos: [cfvo1, cfvo2, cfvo3],
            showValue: false,
            reverse: nil,
            percent: true
        )
        let rule = WorksheetData.ConditionalRule(
            type: .iconSet(iconSet),
            priority: 1,
            dxfId: nil,
            stopIfTrue: false
        )
        builder.addConditionalFormat(WorksheetData.ConditionalFormat(range: "E1:E20", rules: [rule]))
        let xml = String(data: builder.build(), encoding: .utf8)!
        #expect(xml.contains("iconSet iconSet=\"3TrafficLights1\""))
        #expect(xml.contains("showValue=\"0\""))
        #expect(xml.contains("percent=\"1\""))
        #expect(xml.contains("cfvo type=\"percent\" val=\"33\""))
    }

    @Test func roundTripIconSet() throws {
        let tempFile = FileManager.default.temporaryDirectory.appendingPathComponent("iconset_roundtrip.xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: 0) { sheet in
            for i in 1...10 {
                sheet.writeNumber(Double(i * 10), to: CellReference(column: "A", row: i))
            }
            let cfvos = [
                WorksheetData.CFValueObject(type: .percent, value: "0"),
                WorksheetData.CFValueObject(type: .percent, value: "25"),
                WorksheetData.CFValueObject(type: .percent, value: "50"),
                WorksheetData.CFValueObject(type: .percent, value: "75")
            ]
            let iconSet = WorksheetData.ConditionalRule.IconSet(
                name: "4Rating",
                cfvos: cfvos,
                showValue: true,
                reverse: false,
                percent: nil
            )
            let rule = WorksheetData.ConditionalRule(
                type: .iconSet(iconSet),
                priority: 1,
                dxfId: nil,
                stopIfTrue: false
            )
            sheet.addConditionalFormat(range: "A1:A10", rule: rule)
        }

        try writer.save(to: tempFile)

        let workbook = try Workbook.open(url: tempFile)
        guard let sheet = try workbook.sheet(named: "Sheet1") else {
            Issue.record("Sheet1 not found after round-trip")
            return
        }

        #expect(sheet.conditionalFormats.count == 1)
        if case let .iconSet(iset) = sheet.conditionalFormats[0].rules.first?.type {
            #expect(iset.name == "4Rating")
            #expect(iset.cfvos.count == 4)
            #expect(iset.showValue == true)
            #expect(iset.reverse == false)
        } else {
            Issue.record("Expected iconSet rule after round-trip")
        }
    }

    @Test func parseColorScale3Point() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <conditionalFormatting sqref="F1:F100">
            <cfRule type="colorScale" priority="1">
              <colorScale>
                <cfvo type="min"/>
                <cfvo type="percentile" val="50"/>
                <cfvo type="max"/>
                <color rgb="FFF8696B"/>
                <color rgb="FFFFEB84"/>
                <color rgb="FF63BE7B"/>
              </colorScale>
            </cfRule>
          </conditionalFormatting>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)
        #expect(ws.conditionalFormats.count == 1)
        let rule = ws.conditionalFormats[0].rules[0]
        if case let .colorScale(cs) = rule.type {
            #expect(cs.cfvos.count == 3)
            #expect(cs.cfvos[0].type == .min)
            #expect(cs.cfvos[1].type == .percentile)
            #expect(cs.cfvos[1].value == "50")
            #expect(cs.cfvos[2].type == .max)
            #expect(cs.colors.count == 3)
            #expect(cs.colors[0] == "FFF8696B")
            #expect(cs.colors[1] == "FFFFEB84")
            #expect(cs.colors[2] == "FF63BE7B")
        } else {
            Issue.record("Expected colorScale rule")
        }
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
