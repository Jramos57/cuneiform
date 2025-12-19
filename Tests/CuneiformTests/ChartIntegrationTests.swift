import Testing
import Foundation
@testable import Cuneiform

struct ChartIntegrationTests {
    @Test
    func chartDiscoveryInitialization() throws {
        // Create a workbook writer
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Charts")
        
        // Write some basic data
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Category", to: "A1")
            sheet.writeText("Value", to: "B1")
            sheet.writeNumber(100, to: "B2")
            sheet.writeNumber(200, to: "B3")
        }
        
        // Save to temp file
        let tempDir = FileManager.default.temporaryDirectory
        let tempFile = tempDir.appendingPathComponent("chart-test-\(UUID().uuidString).xlsx")
        try writer.save(to: tempFile)
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        // Load and verify structure
        let workbook = try Workbook.open(url: tempFile)
        let sheet = try workbook.sheet(at: 0)
        
        #expect(sheet != nil)
        // Charts would be present if added, but for now we just verify empty array
        #expect(sheet!.charts.isEmpty)
    }

    @Test
    func sheetWithoutChartsHasEmptyChartsList() throws {
        // Create a simple workbook
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "NoCharts")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            sheet.writeText("Data", to: "A1")
        }
        
        // Save and load
        let tempDir = FileManager.default.temporaryDirectory
        let tempFile = tempDir.appendingPathComponent("no-charts-\(UUID().uuidString).xlsx")
        try writer.save(to: tempFile)
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        let workbook = try Workbook.open(url: tempFile)
        let sheet = try workbook.sheet(at: 0)
        
        #expect(sheet?.charts.count == 0)
    }

    @Test
    func multipleSheetChartDiscovery() throws {
        var writer = WorkbookWriter()
        
        // Sheet 1: no charts
        let sheet1 = writer.addSheet(named: "Sheet1")
        writer.modifySheet(at: sheet1) { sheet in
            sheet.writeText("A", to: "A1")
        }
        
        // Sheet 2: will have no charts (we're not adding actual charts yet)
        let sheet2 = writer.addSheet(named: "Sheet2")
        writer.modifySheet(at: sheet2) { sheet in
            sheet.writeText("B", to: "A1")
        }
        
        let tempDir = FileManager.default.temporaryDirectory
        let tempFile = tempDir.appendingPathComponent("multi-sheet-\(UUID().uuidString).xlsx")
        try writer.save(to: tempFile)
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        let workbook = try Workbook.open(url: tempFile)
        
        // Both sheets should load without errors
        let s1 = try workbook.sheet(at: 0)
        let s2 = try workbook.sheet(at: 1)
        
        #expect(s1?.charts.isEmpty == true)
        #expect(s2?.charts.isEmpty == true)
    }
}
