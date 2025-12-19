import Foundation
import Testing
@testable import Cuneiform

@Suite struct SheetProtectionWriteTests {
    @Test func writesUnprotectedSheet() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Data", to: "A1")
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        var package = try OPCPackage.open(url: tempURL)
        let wsData = try package.readPart(PartPath("/xl/worksheets/sheet1.xml"))
        let wsString = String(data: wsData, encoding: .utf8)!
        
        // Should not have sheetProtection element
        #expect(!wsString.contains("<sheetProtection"))
    }

    @Test func writesProtectedSheetNoPassword() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Protected Data", to: "A1")
            sheet.protectSheet()
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        var package = try OPCPackage.open(url: tempURL)
        let wsData = try package.readPart(PartPath("/xl/worksheets/sheet1.xml"))
        let wsString = String(data: wsData, encoding: .utf8)!
        
        #expect(wsString.contains("<sheetProtection"))
        #expect(wsString.contains("sheet=\"1\""))
    }

    @Test func writesProtectedSheetWithPassword() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Secret", to: "A1")
            sheet.protectSheet(password: "myPassword123")
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        var package = try OPCPackage.open(url: tempURL)
        let wsData = try package.readPart(PartPath("/xl/worksheets/sheet1.xml"))
        let wsString = String(data: wsData, encoding: .utf8)!
        
        #expect(wsString.contains("password=\"myPassword123\""))
    }

    @Test func writesProtectionWithStrictOptions() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Data", to: "A1")
            sheet.protectSheet(options: .strict)
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        var package = try OPCPackage.open(url: tempURL)
        let wsData = try package.readPart(PartPath("/xl/worksheets/sheet1.xml"))
        let wsString = String(data: wsData, encoding: .utf8)!
        
        // Strict blocks all operations
        #expect(wsString.contains("formatCells=\"1\""))
        #expect(wsString.contains("insertRows=\"1\""))
        #expect(wsString.contains("deleteColumns=\"1\""))
    }

    @Test func roundTripProtection() throws {
        // Write
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Protected")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Data", to: "B2")
            sheet.protectSheet(password: "secret")
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        // Read and verify
        let workbook = try Workbook.open(url: tempURL)
        guard let sheet = try workbook.sheet(at: 0) else {
            #expect(false)
            return
        }
        
        #expect(sheet.protection != nil)
        let prot = sheet.protection!
        #expect(prot.sheet == true)
        #expect(prot.content == true)
        #expect(prot.passwordHash == "secret")
    }

    @Test func writeProtectionWithCustomOptions() throws {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Data", to: "A1")
            var options = SheetProtectionOptions()
            options.formatCells = false  // Block formatting
            options.insertRows = false   // Block row insertion
            options.deleteColumns = true // Allow column deletion
            sheet.protectSheet(password: "pwd", options: options)
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        let workbook = try Workbook.open(url: tempURL)
        guard let sheet = try workbook.sheet(at: 0) else {
            #expect(false)
            return
        }
        
        let prot = sheet.protection!
        #expect(prot.formatCells == false)
        #expect(prot.insertRows == false)
        #expect(prot.deleteColumns == true)
    }

    @Test func multipleSheetProtection() throws {
        var writer = WorkbookWriter()
        
        let sheet1Index = writer.addSheet(named: "Protected1")
        writer.modifySheet(at: sheet1Index) { sheet in
            sheet.writeText("Secret1", to: "A1")
            sheet.protectSheet(password: "pwd1")
        }
        
        let sheet2Index = writer.addSheet(named: "Open")
        writer.modifySheet(at: sheet2Index) { sheet in
            sheet.writeText("Public", to: "A1")
        }
        
        let sheet3Index = writer.addSheet(named: "Protected2")
        writer.modifySheet(at: sheet3Index) { sheet in
            sheet.writeText("Secret2", to: "A1")
            sheet.protectSheet(password: "pwd2", options: .strict)
        }

        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsx")
        try writer.save(to: tempURL)

        let workbook = try Workbook.open(url: tempURL)
        
        guard let sheet1 = try workbook.sheet(at: 0) else {
            #expect(false)
            return
        }
        #expect(sheet1.protection != nil)
        #expect(sheet1.protection!.passwordHash == "pwd1")
        
        guard let sheet2 = try workbook.sheet(at: 1) else {
            #expect(false)
            return
        }
        #expect(sheet2.protection == nil)
        
        guard let sheet3 = try workbook.sheet(at: 2) else {
            #expect(false)
            return
        }
        #expect(sheet3.protection != nil)
        #expect(sheet3.protection!.passwordHash == "pwd2")
    }
}
