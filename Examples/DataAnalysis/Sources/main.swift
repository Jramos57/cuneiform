import Foundation
import Cuneiform

/// Example: Analyzing Sales Data from an Excel File
///
/// This example demonstrates:
/// - Reading an Excel file
/// - Extracting numerical data
/// - Computing statistics (sum, average, min, max)
/// - Filtering data by criteria
/// - Finding specific values

@main
struct DataAnalysisExample {
    static func main() throws {
        print("📊 Cuneiform Data Analysis Example\n")
        
        // Create sample sales data
        let sampleFile = try createSampleSalesData()
        defer { try? FileManager.default.removeItem(at: sampleFile) }
        
        // Open the workbook
        let workbook = try Workbook.open(url: sampleFile)
        guard let sheet = try workbook.sheet(named: "Sales") else {
            print("❌ Could not find 'Sales' sheet")
            return
        }
        
        print("✅ Opened workbook with \(workbook.sheets.count) sheet(s)")
        print("📄 Sheet: \(sheet.dimension ?? "no dimension")\n")
        
        // Extract sales data from column B (assuming header in row 1)
        print("📈 Extracting sales data from column B...")
        let salesColumn = sheet.column("B")
        
        var salesValues: [Double] = []
        for (row, value) in salesColumn where row > 1 {  // Skip header
            if case .number(let amount) = value {
                salesValues.append(amount)
            }
        }
        
        print("   Found \(salesValues.count) sales records\n")
        
        // Compute statistics
        if !salesValues.isEmpty {
            let total = salesValues.reduce(0, +)
            let average = total / Double(salesValues.count)
            let min = salesValues.min() ?? 0
            let max = salesValues.max() ?? 0
            
            print("📊 Sales Statistics:")
            print("   Total Sales: $\(String(format: "%.2f", total))")
            print("   Average Sale: $\(String(format: "%.2f", average))")
            print("   Minimum Sale: $\(String(format: "%.2f", min))")
            print("   Maximum Sale: $\(String(format: "%.2f", max))\n")
        }
        
        // Find high-value sales (> $500)
        print("💰 High-Value Sales (> $500):")
        let highValueSales = sheet.findAll { ref, value in
            if case .number(let amount) = value, amount > 500 {
                return true
            }
            return false
        }
        
        for (ref, value) in highValueSales {
            if case .number(let amount) = value {
                // Get the product name from column A
                let productRef = CellReference(column: "A", row: ref.row)
                if case .text(let product) = sheet.cell(at: productRef) {
                    print("   \(product): $\(String(format: "%.2f", amount))")
                }
            }
        }
        print()
        
        // Filter rows by region (assuming column C contains region)
        print("🌎 Sales by Region (West):")
        let westSales = sheet.rows { cells in
            cells.contains { cell in
                cell.reference.column == "C" && cell.value == .text("West")
            }
        }
        
        var westTotal = 0.0
        for row in westSales {
            if let salesCell = row.first(where: { $0.reference.column == "B" }),
               case .number(let amount) = salesCell.value {
                westTotal += amount
                
                if let productCell = row.first(where: { $0.reference.column == "A" }),
                   case .text(let product) = productCell.value {
                    print("   \(product): $\(String(format: "%.2f", amount))")
                }
            }
        }
        print("   West Total: $\(String(format: "%.2f", westTotal))\n")
        
        // Demonstrate range queries
        print("📋 Processing range A1:C5:")
        let range = sheet.range("A1:C5")
        
        print("   \(range.count) cells in range")
        for (ref, value) in range where ref.row == 1 {  // Header row
            if case .text(let header) = value {
                print("   Header: \(ref) = \(header)")
            }
        }
        print()
        
        // Streaming large datasets efficiently
        print("🔄 Streaming all rows (memory-efficient):")
        var rowCount = 0
        for row in sheet.rows() {
            rowCount += 1
        }
        print("   Processed \(rowCount) rows using lazy iteration\n")
        
        print("✅ Analysis complete!")
    }
    
    /// Create a sample Excel file with sales data
    static func createSampleSalesData() throws -> URL {
        var writer = WorkbookWriter()
        let sheetIndex = writer.addSheet(named: "Sales")
        
        writer.modifySheet(at: sheetIndex) { sheet in
            // Headers
            sheet.writeText("Product", to: "A1")
            sheet.writeText("Sales", to: "B1")
            sheet.writeText("Region", to: "C1")
            
            // Sample data
            let products = [
                ("Widget A", 450.50, "East"),
                ("Widget B", 675.25, "West"),
                ("Widget C", 325.00, "East"),
                ("Widget D", 890.75, "West"),
                ("Widget E", 550.00, "North"),
                ("Widget F", 425.50, "South"),
                ("Widget G", 725.00, "West"),
                ("Widget H", 310.25, "East"),
                ("Widget I", 625.50, "North"),
                ("Widget J", 475.75, "South")
            ]
            
            for (index, (product, sales, region)) in products.enumerated() {
                let row = index + 2
                sheet.writeText(product, to: CellReference(column: "A", row: row))
                sheet.writeNumber(sales, to: CellReference(column: "B", row: row))
                sheet.writeText(region, to: CellReference(column: "C", row: row))
            }
            
            // Add a total formula
            sheet.writeFormula("SUM(B2:B11)", cachedValue: products.reduce(0) { $0 + $1.1 }, to: "B12")
            sheet.writeText("TOTAL", to: "A12")
        }
        
        let tempDir = FileManager.default.temporaryDirectory
        let fileURL = tempDir.appendingPathComponent("sales-example.xlsx")
        try writer.save(to: fileURL)
        
        return fileURL
    }
}
