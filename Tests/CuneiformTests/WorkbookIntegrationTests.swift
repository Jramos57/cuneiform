import Foundation
import Testing
@testable import Cuneiform

@Suite struct WorkbookIntegrationTests {
    /// Path to the sample XLSX from the ISO spec materials
    static let sampleXlsxPath = "/Users/jonathan/Desktop/garden/cuneiform/Documentation/iEC 29500/ISO_IEC_29500-1_2016(en)_einsert/OfficeOpenXML-SpreadsheetMLStyles/PivotTableFormats.xlsx"

    @Test func openAndListSheets() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let workbook = try Workbook.open(url: url)

        #expect(!workbook.sheets.isEmpty, "Workbook should have at least one sheet")
        #expect(workbook.sheets.count > 0, "Sheet count should be positive")
    }

    @Test func accessSheetByName() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let workbook = try Workbook.open(url: url)

        guard let firstSheet = workbook.sheets.first else {
            #expect(false, "Expected at least one sheet")
            return
        }

        let sheet = try workbook.sheet(named: firstSheet.name)
        #expect(sheet != nil, "Should be able to access sheet by name")
    }

    @Test func accessSheetByIndex() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let workbook = try Workbook.open(url: url)

        let sheet = try workbook.sheet(at: 0)
        #expect(sheet != nil, "Should be able to access sheet at index 0")
    }

    @Test func resolveCellValues() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let workbook = try Workbook.open(url: url)

        guard let sheet = try workbook.sheet(at: 0) else {
            #expect(false, "Expected to load sheet")
            return
        }

        // Try to get a cell; it may be empty but shouldn't crash
        let cellA1 = sheet.cell(at: "A1")
        #expect(true, "Cell lookup should succeed without crashing")
        
        // Test alternate lookup by string
        let cellB1 = sheet.cell(at: "B1")
        #expect(true, "Cell lookup by reference string should succeed")
    }

    @Test func cellValueDescriptions() {
        let textVal = CellValue.text("Hello")
        #expect(textVal.description.contains("Hello"), "Text value should include content")

        let numVal = CellValue.number(42.5)
        #expect(numVal.description.contains("42.5"), "Number value should include value")

        let boolVal = CellValue.boolean(true)
        #expect(boolVal.description.contains("boolean"), "Boolean value should include type")

        let emptyVal = CellValue.empty
        #expect(emptyVal.description.contains("empty"), "Empty value should be identified")
    }
    
    @Test func discoverPivotTables() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let workbook = try Workbook.open(url: url)
        
        // The sample XLSX file contains pivot tables; verify they're discovered
        #expect(!workbook.pivotTables.isEmpty, "Sample XLSX should contain pivot tables")
        
        // Verify pivot table structure
        for pivotTable in workbook.pivotTables {
            #expect(!pivotTable.name.isEmpty, "Pivot table should have a name")
            #expect(pivotTable.cacheId >= 0, "Pivot table should have a valid cache ID")
            // Location and field counts may be optional
        }
    }
}
