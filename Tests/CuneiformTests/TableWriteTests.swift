import Testing
import Foundation
@testable import Cuneiform

@Suite struct TableWriteTests {
    // MARK: - TableBuilder Tests
    
    @Test func buildSimpleTable() throws {
        var builder = TableBuilder(id: 1, displayName: "Sales", name: "SalesTable", ref: "A1:C10")
        builder.addColumn(id: 1, name: "Date")
        builder.addColumn(id: 2, name: "Product")
        builder.addColumn(id: 3, name: "Amount")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("id=\"1\""))
        #expect(xml.contains("name=\"SalesTable\""))
        #expect(xml.contains("displayName=\"Sales\""))
        #expect(xml.contains("ref=\"A1:C10\""))
        #expect(xml.contains("headerRowCount=\"1\""))
        #expect(xml.contains("totalsRowCount=\"0\""))
        #expect(xml.contains("<tableColumn id=\"1\" name=\"Date\"/>"))
        #expect(xml.contains("<tableColumn id=\"2\" name=\"Product\"/>"))
        #expect(xml.contains("<tableColumn id=\"3\" name=\"Amount\"/>"))
    }
    
    @Test func buildTableWithTotals() throws {
        var builder = TableBuilder(id: 2, displayName: "Analysis", name: "AnalysisTable", ref: "A1:D20")
        builder.setRowCounts(header: 1, totals: 1)
        builder.addColumn(id: 1, name: "Category")
        builder.addColumn(id: 2, name: "Count", totalsRowFunction: "count")
        builder.addColumn(id: 3, name: "Average", totalsRowFunction: "average")
        builder.addColumn(id: 4, name: "Total", totalsRowFunction: "sum")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("totalsRowCount=\"1\""))
        #expect(xml.contains("<tableColumn id=\"2\" name=\"Count\" totalsRowFunction=\"count\"/>"))
        #expect(xml.contains("<tableColumn id=\"3\" name=\"Average\" totalsRowFunction=\"average\"/>"))
        #expect(xml.contains("<tableColumn id=\"4\" name=\"Total\" totalsRowFunction=\"sum\"/>"))
    }
    
    @Test func buildTableWithAutoFilter() throws {
        var builder = TableBuilder(id: 1, displayName: "Filtered", name: "FilterTable", ref: "A1:B50")
        builder.setAutoFilter(ref: "A1:B50")
        builder.addColumn(id: 1, name: "Name")
        builder.addColumn(id: 2, name: "Score")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("<autoFilter ref=\"A1:B50\"/>"))
    }
    
    @Test func buildTableWithStyle() throws {
        var builder = TableBuilder(id: 1, displayName: "Styled", name: "StyleTable", ref: "A1:C5")
        builder.setTableStyle(name: "TableStyleMedium2", showFirstColumn: false, showLastColumn: true, showRowStripes: true, showColumnStripes: false)
        builder.addColumn(id: 1, name: "A")
        builder.addColumn(id: 2, name: "B")
        builder.addColumn(id: 3, name: "C")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("name=\"TableStyleMedium2\""))
        #expect(xml.contains("showFirstColumn=\"0\""))
        #expect(xml.contains("showLastColumn=\"1\""))
        #expect(xml.contains("showRowStripes=\"1\""))
        #expect(xml.contains("showColumnStripes=\"0\""))
    }
    
    @Test func buildTableWithSpecialCharacters() throws {
        var builder = TableBuilder(id: 1, displayName: "Special & Chars", name: "Special<>\"", ref: "A1:B2")
        builder.addColumn(id: 1, name: "Col & Name")
        builder.addColumn(id: 2, name: "Test < Value")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        // Should be XML-escaped
        #expect(xml.contains("&amp;"))
        #expect(xml.contains("&lt;"))
        #expect(xml.contains("&gt;"))
        #expect(xml.contains("&quot;"))
    }
    
    // MARK: - WorksheetBuilder Integration Tests
    
    @Test func worksheetWithTable() throws {
        var builder = WorksheetBuilder()
        builder.addCell(at: CellReference(column: "A", row: 1), value: .text("Name"))
        builder.addCell(at: CellReference(column: "B", row: 1), value: .text("Score"))
        builder.addCell(at: CellReference(column: "A", row: 2), value: .text("Alice"))
        builder.addCell(at: CellReference(column: "B", row: 2), value: .number(95))
        
        var tableBuilder = TableBuilder(id: 1, displayName: "Scores", name: "ScoreTable", ref: "A1:B2")
        tableBuilder.addColumn(id: 1, name: "Name")
        tableBuilder.addColumn(id: 2, name: "Score")
        builder.addTable(tableBuilder)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        // Should have both cell data and table parts
        #expect(xml.contains("<sheetData>"))
        #expect(xml.contains("<tableParts count=\"1\">"))
        #expect(xml.contains("<tablePart r:id=\"rIdTable1\"/>"))
    }
    
    @Test func worksheetWithMultipleTables() throws {
        var builder = WorksheetBuilder()
        
        // First table
        var table1 = TableBuilder(id: 1, displayName: "Sales", name: "SalesTable", ref: "A1:B5")
        table1.addColumn(id: 1, name: "Date")
        table1.addColumn(id: 2, name: "Amount")
        builder.addTable(table1)
        
        // Second table
        var table2 = TableBuilder(id: 2, displayName: "Expenses", name: "ExpenseTable", ref: "D1:E5")
        table2.addColumn(id: 1, name: "Category")
        table2.addColumn(id: 2, name: "Cost")
        builder.addTable(table2)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("<tableParts count=\"2\">"))
        #expect(xml.contains("<tablePart r:id=\"rIdTable1\"/>"))
        #expect(xml.contains("<tablePart r:id=\"rIdTable2\"/>"))
    }
    
    // MARK: - SheetWriter Integration Tests
    
    @Test func sheetWriterAddTable() throws {
        let tempFile = FileManager.default.temporaryDirectory.appendingPathComponent("test_table.xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Data")
        
        writer.modifySheet(at: 0) { sheet in
            sheet.writeText("Product", to: "A1")
            sheet.writeText("Sales", to: "B1")
            sheet.writeNumber(100, to: "A2")
            sheet.writeNumber(1500, to: "B2")
            
            // Add table
            sheet.addTable(
                name: "SalesData",
                displayName: "Sales Data",
                ref: "A1:B2",
                columns: [(id: 1, name: "Product", totalsFunction: nil), (id: 2, name: "Sales", totalsFunction: "sum")],
                headerRowCount: 1,
                totalsRowCount: 1
            )
        }
        
        try writer.save(to: tempFile)
        
        // Verify file was created
        #expect(FileManager.default.fileExists(atPath: tempFile.path))
        
        // Can read it back
        let workbook = try Workbook.open(url: tempFile)
        #expect(workbook.tables.count > 0)
        
        let table = workbook.tables.first
        #expect(table?.name == "SalesData")
        #expect(table?.displayName == "Sales Data")
        #expect(table?.ref == "A1:B2")
        #expect(table?.columns.count == 2)
    }
    
    @Test func tableRoundTrip() throws {
        let tempFile = FileManager.default.temporaryDirectory.appendingPathComponent("table_roundtrip.xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        // Write a table
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Inventory")
        
        writer.modifySheet(at: 0) { sheet in
            sheet.writeText("SKU", to: "A1")
            sheet.writeText("Name", to: "B1")
            sheet.writeText("Qty", to: "C1")
            
            sheet.addTable(
                name: "InventoryList",
                ref: "A1:C1",
                columns: [
                    (id: 1, name: "SKU", totalsFunction: nil),
                    (id: 2, name: "Name", totalsFunction: nil),
                    (id: 3, name: "Qty", totalsFunction: "sum")
                ],
                headerRowCount: 1,
                totalsRowCount: 1
            )
        }
        
        try writer.save(to: tempFile)
        
        // Read it back
        let workbook = try Workbook.open(url: tempFile)
        let table = workbook.tables.first
        
        #expect(table?.name == "InventoryList")
        #expect(table?.ref == "A1:C1")
        #expect(table?.headerRowCount == 1)
        #expect(table?.totalsRowCount == 1)
        #expect(table?.columns.count == 3)
        #expect(table?.columns[0].name == "SKU")
        #expect(table?.columns[1].name == "Name")
        #expect(table?.columns[2].name == "Qty")
        #expect(table?.columns[2].totalsRowFunction == "sum")
    }
    
    @Test func multiTableWorksheet() throws {
        let tempFile = FileManager.default.temporaryDirectory.appendingPathComponent("multi_table.xlsx")
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        _ = writer.addSheet(named: "Report")
        
        writer.modifySheet(at: 0) { sheet in
            // First table
            sheet.writeText("Region", to: "A1")
            sheet.writeText("Revenue", to: "B1")
            sheet.addTable(name: "RegionalRevenue", ref: "A1:B10", columns: [(1, "Region", nil), (2, "Revenue", "sum")])
            
            // Second table
            sheet.writeText("Product", to: "D1")
            sheet.writeText("Units", to: "E1")
            sheet.addTable(name: "ProductUnits", ref: "D1:E10", columns: [(1, "Product", nil), (2, "Units", "sum")])
        }
        
        try writer.save(to: tempFile)
        
        let workbook = try Workbook.open(url: tempFile)
        #expect(workbook.tables.count >= 2)
    }
}
