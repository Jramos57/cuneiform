import Foundation
import Testing
@testable import Cuneiform

@Suite struct SheetProtectionParserTests {
    @Test func parseUnprotectedSheet() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
        </worksheet>
        """.data(using: .utf8)!

        let data = try WorksheetParser.parse(data: xml)
        #expect(data.protection == nil)
    }

    @Test func parseProtectedSheetNoPassword() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <sheetProtection sheet="1" content="1"/>
        </worksheet>
        """.data(using: .utf8)!

        let data = try WorksheetParser.parse(data: xml)
        #expect(data.protection != nil)
        let prot = data.protection!
        #expect(prot.sheet == true)
        #expect(prot.content == true)
        #expect(prot.passwordHash == nil)
    }

    @Test func parseProtectedSheetWithPassword() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <sheetProtection sheet="1" content="1" password="ABC123"/>
        </worksheet>
        """.data(using: .utf8)!

        let data = try WorksheetParser.parse(data: xml)
        #expect(data.protection != nil)
        let prot = data.protection!
        #expect(prot.passwordHash == "ABC123")
    }

    @Test func parseProtectionFlags() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <sheetProtection sheet="1" content="1" formatCells="1" insertRows="1" deleteColumns="1" sort="1"/>
        </worksheet>
        """.data(using: .utf8)!

        let data = try WorksheetParser.parse(data: xml)
        let prot = data.protection!
        
        // Explicit true flags
        #expect(prot.sheet == true)
        #expect(prot.content == true)
        #expect(prot.formatCells == false)  // inverted: "1" means disabled
        #expect(prot.insertRows == false)
        #expect(prot.deleteColumns == false)
        #expect(prot.sort == false)
        
        // Default true flags (not specified)
        #expect(prot.formatColumns == true)
        #expect(prot.formatRows == true)
    }

    @Test func parseStrictProtection() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <sheetProtection sheet="1" content="1" formatCells="1" formatColumns="1" formatRows="1" insertColumns="1" insertRows="1" insertHyperlinks="1" deleteColumns="1" deleteRows="1" sort="1" autoFilter="1" pivotTables="1"/>
        </worksheet>
        """.data(using: .utf8)!

        let data = try WorksheetParser.parse(data: xml)
        let prot = data.protection!
        
        // All operations blocked except select (defaults)
        #expect(prot.formatCells == false)
        #expect(prot.formatColumns == false)
        #expect(prot.formatRows == false)
        #expect(prot.insertColumns == false)
        #expect(prot.insertRows == false)
        #expect(prot.insertHyperlinks == false)
        #expect(prot.deleteColumns == false)
        #expect(prot.deleteRows == false)
        #expect(prot.sort == false)
        #expect(prot.autoFilter == false)
        #expect(prot.pivotTables == false)
        
        // Selection allowed
        #expect(prot.selectLockedCells == true)
        #expect(prot.selectUnlockedCells == true)
    }

    @Test func sheetExposeProtection() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <sheetProtection sheet="1" content="1" password="pwd123"/>
        </worksheet>
        """.data(using: .utf8)!

        let data = try WorksheetParser.parse(data: xml)
        let sheet = Sheet(data: data, sharedStrings: .empty, styles: .empty)
        
        #expect(sheet.protection != nil)
        #expect(sheet.protection!.sheet == true)
        #expect(sheet.protection!.passwordHash == "pwd123")
    }
}
