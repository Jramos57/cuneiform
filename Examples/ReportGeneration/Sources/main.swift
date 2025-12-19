import Foundation
import Cuneiform

/// Example: Generating Structured Reports in Excel
///
/// This example demonstrates:
/// - Creating multi-sheet workbooks
/// - Writing structured data with headers
/// - Using formulas for calculations
/// - Creating summary sheets
/// - Organizing data across multiple sheets

@main
struct ReportGenerationExample {
    static func main() throws {
        print("📝 Cuneiform Report Generation Example\n")
        
        // Create a comprehensive quarterly report
        let reportURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("quarterly-report.xlsx")
        
        try generateQuarterlyReport(to: reportURL)
        
        print("✅ Report generated: \(reportURL.lastPathComponent)")
        print("📊 File size: \(try fileSize(at: reportURL)) KB\n")
        
        // Verify the report by reading it back
        try verifyReport(at: reportURL)
        
        print("\n✅ Report generation complete!")
        print("📁 Report saved to: \(reportURL.path)")
    }
    
    /// Generate a quarterly sales report with multiple sheets
    static func generateQuarterlyReport(to url: URL) throws {
        print("📋 Generating quarterly report...")
        
        var writer = WorkbookWriter()
        
        // Sheet 1: Executive Summary
        print("   Creating Executive Summary...")
        let summaryIndex = writer.addSheet(named: "Executive Summary")
        writer.modifySheet(at: summaryIndex) { sheet in
            // Title
            sheet.writeText("Q4 2024 Sales Report", to: "A1")
            sheet.writeText("Generated: \(formattedDate())", to: "A2")
            
            // Summary metrics
            sheet.writeText("Metric", to: "A4")
            sheet.writeText("Value", to: "B4")
            
            sheet.writeText("Total Revenue", to: "A5")
            sheet.writeFormula("SUM('Regional Sales'!B2:B5)", cachedValue: 458750, to: "B5")
            
            sheet.writeText("Total Units", to: "A6")
            sheet.writeFormula("SUM('Regional Sales'!C2:C5)", cachedValue: 12450, to: "B6")
            
            sheet.writeText("Average Unit Price", to: "A7")
            sheet.writeFormula("B5/B6", cachedValue: 36.85, to: "B7")
            
            sheet.writeText("Best Region", to: "A8")
            sheet.writeText("West", to: "B8")
        }
        
        // Sheet 2: Regional Sales
        print("   Creating Regional Sales...")
        let regionalIndex = writer.addSheet(named: "Regional Sales")
        writer.modifySheet(at: regionalIndex) { sheet in
            // Headers
            sheet.writeText("Region", to: "A1")
            sheet.writeText("Revenue", to: "B1")
            sheet.writeText("Units", to: "C1")
            sheet.writeText("Avg Price", to: "D1")
            
            // Data
            let regions = [
                ("North", 125000.0, 3200),
                ("South", 98750.0, 2850),
                ("East", 112500.0, 3100),
                ("West", 122500.0, 3300)
            ]
            
            for (index, (region, revenue, units)) in regions.enumerated() {
                let row = index + 2
                sheet.writeText(region, to: CellReference(column: "A", row: row))
                sheet.writeNumber(revenue, to: CellReference(column: "B", row: row))
                sheet.writeNumber(Double(units), to: CellReference(column: "C", row: row))
                
                // Calculate average price with formula
                let avgFormula = "B\(row)/C\(row)"
                sheet.writeFormula(avgFormula, cachedValue: revenue / Double(units), 
                                 to: CellReference(column: "D", row: row))
            }
            
            // Totals
            let totalRow = regions.count + 2
            sheet.writeText("TOTAL", to: CellReference(column: "A", row: totalRow))
            sheet.writeFormula("SUM(B2:B5)", cachedValue: 458750, 
                             to: CellReference(column: "B", row: totalRow))
            sheet.writeFormula("SUM(C2:C5)", cachedValue: 12450,
                             to: CellReference(column: "C", row: totalRow))
        }
        
        // Sheet 3: Product Performance
        print("   Creating Product Performance...")
        let productIndex = writer.addSheet(named: "Product Performance")
        writer.modifySheet(at: productIndex) { sheet in
            // Headers
            sheet.writeText("Product", to: "A1")
            sheet.writeText("Category", to: "B1")
            sheet.writeText("Q3 Sales", to: "C1")
            sheet.writeText("Q4 Sales", to: "D1")
            sheet.writeText("Growth %", to: "E1")
            
            // Sample products
            let products = [
                ("Widget Pro", "Hardware", 45000.0, 52000.0),
                ("Widget Lite", "Hardware", 32000.0, 38000.0),
                ("Service Plan A", "Services", 28000.0, 31000.0),
                ("Service Plan B", "Services", 18000.0, 22000.0),
                ("Accessory Kit", "Accessories", 15000.0, 17000.0)
            ]
            
            for (index, (product, category, q3, q4)) in products.enumerated() {
                let row = index + 2
                sheet.writeText(product, to: CellReference(column: "A", row: row))
                sheet.writeText(category, to: CellReference(column: "B", row: row))
                sheet.writeNumber(q3, to: CellReference(column: "C", row: row))
                sheet.writeNumber(q4, to: CellReference(column: "D", row: row))
                
                // Growth calculation
                let growthFormula = "((D\(row)-C\(row))/C\(row))*100"
                let growth = ((q4 - q3) / q3) * 100
                sheet.writeFormula(growthFormula, cachedValue: growth,
                                 to: CellReference(column: "E", row: row))
            }
        }
        
        // Sheet 4: Monthly Breakdown
        print("   Creating Monthly Breakdown...")
        let monthlyIndex = writer.addSheet(named: "Monthly Breakdown")
        writer.modifySheet(at: monthlyIndex) { sheet in
            // Headers
            sheet.writeText("Month", to: "A1")
            sheet.writeText("Revenue", to: "B1")
            sheet.writeText("% of Quarter", to: "C1")
            
            let months = [
                ("October", 148250.0),
                ("November", 156000.0),
                ("December", 154500.0)
            ]
            
            let quarterTotal = months.reduce(0) { $0 + $1.1 }
            
            for (index, (month, revenue)) in months.enumerated() {
                let row = index + 2
                sheet.writeText(month, to: CellReference(column: "A", row: row))
                sheet.writeNumber(revenue, to: CellReference(column: "B", row: row))
                
                let pctFormula = "(B\(row)/\(quarterTotal))*100"
                let pct = (revenue / quarterTotal) * 100
                sheet.writeFormula(pctFormula, cachedValue: pct,
                                 to: CellReference(column: "C", row: row))
            }
            
            // Total
            sheet.writeText("TOTAL", to: "A5")
            sheet.writeFormula("SUM(B2:B4)", cachedValue: quarterTotal, to: "B5")
            sheet.writeNumber(100, to: "C5")
        }
        
        print("   Saving workbook...")
        try writer.save(to: url)
    }
    
    /// Verify the generated report by reading it back
    static func verifyReport(at url: URL) throws {
        print("🔍 Verifying generated report...")
        
        let workbook = try Workbook.open(url: url)
        
        print("   Sheets: \(workbook.sheets.map(\.name).joined(separator: ", "))")
        
        // Check Executive Summary
        if let summary = try workbook.sheet(named: "Executive Summary") {
            if case .text(let title) = summary.cell(at: "A1") {
                print("   ✓ Title: \(title)")
            }
        }
        
        // Check Regional Sales totals
        if let regional = try workbook.sheet(named: "Regional Sales") {
            let totalRow = 6
            if case .number(let revenue) = regional.cell(at: CellReference(column: "B", row: totalRow)) {
                print("   ✓ Total Revenue: $\(String(format: "%.2f", revenue))")
            }
        }
        
        // Check sheet count
        print("   ✓ Created \(workbook.sheets.count) sheets")
    }
    
    // MARK: - Helper Functions
    
    static func formattedDate() -> String {
        let formatter = DateFormatter()
        formatter.dateStyle = .medium
        return formatter.string(from: Date())
    }
    
    static func fileSize(at url: URL) throws -> Int {
        let attributes = try FileManager.default.attributesOfItem(atPath: url.path)
        let size = attributes[.size] as? Int ?? 0
        return size / 1024  // Convert to KB
    }
}
