import Testing
import Foundation
@testable import Cuneiform

@Suite("Workbook Protection Write Tests")
struct WorkbookProtectionWriteTests {
    
    /// Test: protect workbook with default options (no protection)
    @Test("Protect workbook with no flags (default)")
    func protectWorkbookDefault() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        writer.protectWorkbook()
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        #expect(workbook.protection != nil)
        #expect(workbook.protection?.structureProtected == false)
        #expect(workbook.protection?.windowsProtected == false)
    }
    
    /// Test: protect workbook structure only
    @Test("Protect workbook structure only")
    func protectWorkbookStructureOnly() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        writer.protectWorkbook(options: .structureOnly)
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        #expect(workbook.protection != nil)
        #expect(workbook.protection?.structureProtected == true)
        #expect(workbook.protection?.windowsProtected == false)
    }
    
    /// Test: protect workbook structure and windows
    @Test("Protect workbook structure and windows (strict)")
    func protectWorkbookStrict() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        writer.protectWorkbook(options: .strict)
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        #expect(workbook.protection != nil)
        #expect(workbook.protection?.structureProtected == true)
        #expect(workbook.protection?.windowsProtected == true)
    }
    
    /// Test: protect workbook windows only
    @Test("Protect workbook windows only")
    func protectWorkbookWindowsOnly() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        
        var options = WorkbookProtectionOptions()
        options.windows = true
        writer.protectWorkbook(options: options)
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        #expect(workbook.protection != nil)
        #expect(workbook.protection?.structureProtected == false)
        #expect(workbook.protection?.windowsProtected == true)
    }
    
    /// Test: protect workbook with password
    @Test("Protect workbook with password hash")
    func protectWorkbookWithPassword() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        writer.protectWorkbook(password: "secret123", options: .structureOnly)
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        #expect(workbook.protection != nil)
        #expect(workbook.protection?.structureProtected == true)
        #expect(workbook.protection?.windowsProtected == false)
        #expect(workbook.protection?.passwordHash != nil)
        #expect(!workbook.protection!.passwordHash!.isEmpty)
    }
    
    /// Test: round-trip with workbook protection preserves protection settings
    @Test("Round-trip: write → read → write → read preserves protection")
    func roundTripProtection() throws {
        // Write original
        var writer1 = WorkbookWriter()
        _ = writer1.addSheet(named: "Original")
        writer1.protectWorkbook(password: "test", options: .strict)
        
        let data1 = try writer1.buildData()
        let url1 = try createTempFile(data: data1)
        defer { try? FileManager.default.removeItem(at: url1) }
        
        // Read original
        let workbook1 = try Workbook.open(url: url1)
        let prot1 = workbook1.protection
        
        #expect(prot1?.structureProtected == true)
        #expect(prot1?.windowsProtected == true)
        #expect(prot1?.passwordHash != nil)
        
        // Write again with same protection
        var writer2 = WorkbookWriter()
        _ = writer2.addSheet(named: "Round 2")
        if let protection = prot1 {
            writer2.protectWorkbook(options: WorkbookProtectionOptions(
                structure: protection.structureProtected,
                windows: protection.windowsProtected
            ))
        }
        
        let data2 = try writer2.buildData()
        let url2 = try createTempFile(data: data2)
        defer { try? FileManager.default.removeItem(at: url2) }
        
        // Read again
        let workbook2 = try Workbook.open(url: url2)
        let prot2 = workbook2.protection
        
        #expect(prot2?.structureProtected == prot1?.structureProtected)
        #expect(prot2?.windowsProtected == prot1?.windowsProtected)
    }
    
    /// Test: unprotected workbook has no protection element
    @Test("Unprotected workbook has no protection")
    func unprotectedWorkbook() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        // Don't call protectWorkbook()
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        #expect(workbook.protection == nil)
    }
    
    /// Test: multiple sheets with workbook protection
    @Test("Multiple sheets with workbook protection")
    func multipleSheetWithProtection() throws {
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Sheet1")
        _ = writer.addSheet(named: "Sheet2")
        _ = writer.addSheet(named: "Sheet3")
        writer.protectWorkbook(options: .structureOnly)
        
        // Also protect individual sheets
        writer.modifySheet(at: 0) { sheet in
            sheet.protectSheet(password: "sheet1pwd")
        }
        writer.modifySheet(at: 1) { sheet in
            sheet.protectSheet(password: "sheet2pwd", options: .readonly)
        }
        
        let data = try writer.buildData()
        let url = try createTempFile(data: data)
        defer { try? FileManager.default.removeItem(at: url) }
        
        let workbook = try Workbook.open(url: url)
        
        // Check workbook protection
        #expect(workbook.protection?.structureProtected == true)
        
        // Check sheet protection
        if let sheet1 = try workbook.sheet(at: 0) {
            #expect(sheet1.protection != nil)
            #expect(sheet1.protection?.content == true)
        }
        
        if let sheet2 = try workbook.sheet(at: 1) {
            #expect(sheet2.protection != nil)
            #expect(sheet2.protection?.content == true)
        }
    }
}

// MARK: - Helper

private func createTempFile(data: Data) throws -> URL {
    let tempDir = FileManager.default.temporaryDirectory
    let filename = "test_\(UUID().uuidString).xlsx"
    let url = tempDir.appendingPathComponent(filename)
    try data.write(to: url)
    return url
}

