import Foundation
import Testing

@testable import Cuneiform

@Suite("Data Validation Expanded Variants")
struct DataValidationExpandedTests {
    @Test("Decimal >= validation")
    func decimalGreaterOrEqual() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dv_decimal_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "ValidateOps")
        writer.modifySheet(at: idx) { sheet in
            sheet.writeText("Value:", to: CellReference(column: "B", row: 1))
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .decimal,
                allowBlank: true,
                sqref: "C3:C10",
                formula1: "1.5",
                formula2: nil,
                op: "greaterThanOrEqual"
            ))
        }
        try writer.save(to: tempFile)

        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let info = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[info.relationshipId] else {
            #expect(false)
            return
        }
        let path = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try String(data: pkg.readPart(path), encoding: .utf8)!
        #expect(sheetXML.contains("<dataValidations"))
        #expect(sheetXML.contains("type=\"decimal\""))
        #expect(sheetXML.contains("operator=\"greaterThanOrEqual\""))
        #expect(sheetXML.contains("sqref=\"C3:C10\""))
        #expect(sheetXML.contains("<formula1>1.5</formula1>"))
    }

    @Test("Date between validation")
    func dateBetween() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dv_date_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "ValidateDates")
        writer.modifySheet(at: idx) { sheet in
            sheet.writeText("Date:", to: CellReference(column: "A", row: 1))
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .date,
                allowBlank: true,
                sqref: "D2:D20",
                formula1: "DATE(2024,1,1)",
                formula2: "DATE(2024,12,31)",
                op: "between"
            ))
        }
        try writer.save(to: tempFile)

        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let info = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[info.relationshipId] else {
            #expect(false)
            return
        }
        let path = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try String(data: pkg.readPart(path), encoding: .utf8)!
        #expect(sheetXML.contains("type=\"date\""))
        #expect(sheetXML.contains("operator=\"between\""))
        #expect(sheetXML.contains("sqref=\"D2:D20\""))
        #expect(sheetXML.contains("<formula1>DATE(2024,1,1)</formula1>"))
        #expect(sheetXML.contains("<formula2>DATE(2024,12,31)</formula2>"))
    }

    @Test("Whole number <= validation")
    func wholeLessOrEqual() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dv_whole_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "ValidateWhole")
        writer.modifySheet(at: idx) { sheet in
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .whole,
                allowBlank: false,
                sqref: "E1:E5",
                formula1: "100",
                formula2: nil,
                op: "lessThanOrEqual"
            ))
        }
        try writer.save(to: tempFile)

        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let info = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[info.relationshipId] else {
            #expect(false)
            return
        }
        let path = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try String(data: pkg.readPart(path), encoding: .utf8)!
        #expect(sheetXML.contains("type=\"whole\""))
        #expect(sheetXML.contains("operator=\"lessThanOrEqual\""))
        #expect(sheetXML.contains("sqref=\"E1:E5\""))
        #expect(sheetXML.contains("<formula1>100</formula1>"))
    }

    @Test("List validation with range reference")
    func listRangeReference() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dv_list_range_\(UUID().uuidString).xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }

        var writer = WorkbookWriter()
        let idx = writer.addSheet(named: "ValidateList")
        writer.modifySheet(at: idx) { sheet in
            sheet.addDataValidation(WorksheetBuilder.DataValidation(
                type: .list,
                allowBlank: true,
                sqref: "B2:B20",
                formula1: "ValidateList!$A$1:$A$10"
            ))
        }
        try writer.save(to: tempFile)

        var pkg = try OPCPackage.open(url: tempFile)
        let wbRels = try pkg.relationships(for: .workbook)
        guard let info = try Workbook.open(url: tempFile).sheets.first,
              let rel = wbRels[info.relationshipId] else {
            #expect(false)
            return
        }
        let path = rel.resolveTarget(relativeTo: .workbook)
        let sheetXML = try String(data: pkg.readPart(path), encoding: .utf8)!
        #expect(sheetXML.contains("type=\"list\""))
        #expect(sheetXML.contains("sqref=\"B2:B20\""))
        #expect(sheetXML.contains("<formula1>ValidateList!$A$1:$A$10</formula1>"))
    }
}
