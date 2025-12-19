import Testing
import Foundation
@testable import Cuneiform

@Suite("Pivot Table Parser Tests")
struct PivotTableParserTests {
    
    /// Test: parse minimal pivot table XML
    @Test("Parse minimal pivot table")
    func parseMinimalPivotTable() throws {
        // Test WITH namespace
        let xml = """
        <?xml version="1.0"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Test" cacheId="0">
            <location ref="A1:B2"/>
        </pivotTableDefinition>
        """
        
        let data = xml.data(using: .utf8)!
        let pt = try PivotTableParser.parse(data: data)
        
        #expect(pt.name == "Test")
        #expect(pt.cacheId == 0)
        #expect(pt.location == "A1:B2")
    }
    
    /// Test: parse basic pivot table XML
    @Test("Parse basic pivot table")
    func parseBasicPivotTable() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            name="SalesPivot" cacheId="0" autoFormatId="4096" useAutoFormatting="1"
            dataCaption="Data" itemPrintTitles="1">
            <location ref="A1:D10" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"/>
            <rowFields count="1"/>
            <colFields count="1"/>
            <dataFields count="1"/>
        </pivotTableDefinition>
        """
        
        let data = xml.data(using: .utf8)!
        let pt = try PivotTableParser.parse(data: data)
        
        #expect(pt.name == "SalesPivot")
        #expect(pt.cacheId == 0)
        #expect(pt.location == "A1:D10")
        #expect(pt.useAutoFormatting == true)
        #expect(pt.rowFieldCount == 1)
        #expect(pt.colFieldCount == 1)
        #expect(pt.dataFieldCount == 1)
    }
    
    /// Test: parse pivot table with multiple fields
    @Test("Parse pivot table with multiple row/col fields")
    func parseMultiFieldPivotTable() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            name="DetailedPivot" cacheId="5" useAutoFormatting="0" dataCaption="Value">
            <location ref="B5:G50" firstHeaderRow="1" firstDataRow="1" firstDataCol="2"/>
            <rowFields count="2"/>
            <colFields count="1"/>
            <dataFields count="1"/>
        </pivotTableDefinition>
        """
        
        let data = xml.data(using: .utf8)!
        let pt = try PivotTableParser.parse(data: data)
        
        #expect(pt.name == "DetailedPivot")
        #expect(pt.cacheId == 5)
        #expect(pt.useAutoFormatting == false)
        #expect(pt.rowFieldCount == 2)
        #expect(pt.colFieldCount == 1)
        #expect(pt.dataFieldCount == 1)
        #expect(pt.location == "B5:G50")
    }
    
    /// Test: parse pivot table with multiple data fields
    @Test("Parse pivot table with multiple data fields")
    func parseMultiDataFieldPivotTable() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            name="ComparisonPivot" cacheId="2">
            <location ref="C3:H25"/>
            <rowFields count="2"/>
            <dataFields count="2"/>
        </pivotTableDefinition>
        """
        
        let data = xml.data(using: .utf8)!
        let pt = try PivotTableParser.parse(data: data)
        
        #expect(pt.dataFieldCount == 2)
        #expect(pt.rowFieldCount == 2)
    }
    
    /// Test: parse pivot table with high cache ID
    @Test("Parse pivot table with high cache ID")
    func parsePivotTableWithHighCacheId() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            name="LargePivot" cacheId="999">
            <location ref="A1:Z100"/>
        </pivotTableDefinition>
        """
        
        let data = xml.data(using: .utf8)!
        let pt = try PivotTableParser.parse(data: data)
        
        #expect(pt.cacheId == 999)
        #expect(pt.name == "LargePivot")
    }
    
    /// Test: pivot table equatable
    @Test("Pivot table data is equatable")
    func pivotTableEquatable() throws {
        let pt1 = PivotTableData(
            name: "Test",
            cacheId: 1,
            location: "A1:B2",
            fieldNames: ["Field1", "Field2"],
            rowFieldCount: 1,
            colFieldCount: 0,
            dataFieldCount: 1,
            useAutoFormatting: true
        )
        
        let pt2 = PivotTableData(
            name: "Test",
            cacheId: 1,
            location: "A1:B2",
            fieldNames: ["Field1", "Field2"],
            rowFieldCount: 1,
            colFieldCount: 0,
            dataFieldCount: 1,
            useAutoFormatting: true
        )
        
        #expect(pt1 == pt2)
    }
    
    /// Test: invalid XML throws error
    @Test("Invalid pivot table XML throws error")
    func invalidXMLThrowsError() throws {
        let xml = "<?xml version=\"1.0\"?><invalid></invalid>"
        let data = xml.data(using: .utf8)!
        
        #expect(throws: CuneiformError.self) {
            try PivotTableParser.parse(data: data)
        }
    }
}

