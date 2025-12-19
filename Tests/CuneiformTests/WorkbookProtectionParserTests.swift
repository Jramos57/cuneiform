import Testing
@testable import Cuneiform

struct WorkbookProtectionParserTests {
    @Test
    func parseUnprotectedWorkbook() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """
        
        let data = xml.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: data)
        
        #expect(info.protection == nil)
    }

    @Test
    func parseProtectedWorkbookStructureOnly() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <workbookProtection sheet="1"/>
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """
        
        let data = xml.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: data)
        
        #expect(info.protection != nil)
        #expect(info.protection?.structureProtected == true)
        #expect(info.protection?.windowsProtected == false)
        #expect(info.protection?.passwordHash == nil)
    }

    @Test
    func parseProtectedWorkbookWithPassword() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <workbookProtection sheet="1" windows="1" password="7B2D8A6C"/>
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """
        
        let data = xml.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: data)
        
        #expect(info.protection != nil)
        #expect(info.protection?.structureProtected == true)
        #expect(info.protection?.windowsProtected == true)
        #expect(info.protection?.passwordHash == "7B2D8A6C")
    }

    @Test
    func parseProtectedWorkbookWindowsOnly() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <workbookProtection windows="1"/>
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """
        
        let data = xml.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: data)
        
        #expect(info.protection != nil)
        #expect(info.protection?.structureProtected == false)
        #expect(info.protection?.windowsProtected == true)
    }

    @Test
    func parseProtectedWorkbookExplicitlyUnprotected() throws {
        let xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <workbookProtection sheet="0" windows="0"/>
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """
        
        let data = xml.data(using: .utf8)!
        let info = try WorkbookParser.parse(data: data)
        
        #expect(info.protection != nil)
        #expect(info.protection?.structureProtected == false)
        #expect(info.protection?.windowsProtected == false)
    }

    @Test
    func workbookProtectionEquatable() throws {
        let p1 = WorkbookProtection(structureProtected: true, windowsProtected: false, passwordHash: "ABC123")
        let p2 = WorkbookProtection(structureProtected: true, windowsProtected: false, passwordHash: "ABC123")
        let p3 = WorkbookProtection(structureProtected: false, windowsProtected: false, passwordHash: "ABC123")

        #expect(p1 == p2)
        #expect(p1 != p3)
    }
}
