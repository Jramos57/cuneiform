# Report Generation Example

This example demonstrates how to use Cuneiform to create structured, multi-sheet Excel reports.

## What This Example Shows

- Creating workbooks with multiple sheets
- Writing headers and structured data
- Using formulas for calculations and aggregations
- Cross-sheet formulas (referencing other sheets)
- Creating summary sheets with totals
- Organizing complex data across sheets

## Running the Example

```bash
cd Examples/ReportGeneration
swift run
```

## Example Output

```
📝 Cuneiform Report Generation Example

📋 Generating quarterly report...
   Creating Executive Summary...
   Creating Regional Sales...
   Creating Product Performance...
   Creating Monthly Breakdown...
   Saving workbook...
✅ Report generated: quarterly-report.xlsx
📊 File size: 12 KB

🔍 Verifying generated report...
   Sheets: Executive Summary, Regional Sales, Product Performance, Monthly Breakdown
   ✓ Title: Q4 2024 Sales Report
   ✓ Total Revenue: $458750.00
   ✓ Created 4 sheets

✅ Report generation complete!
📁 Report saved to: /tmp/quarterly-report.xlsx
```

## Report Structure

The generated report contains 4 sheets:

### 1. Executive Summary
- Report title and date
- Key metrics with formulas referencing other sheets
- Summary statistics

### 2. Regional Sales
- Sales data by region
- Revenue and units sold
- Calculated average prices using formulas
- Totals with SUM formulas

### 3. Product Performance
- Product-level sales data
- Quarter-over-quarter comparison
- Growth percentage calculations

### 4. Monthly Breakdown
- Month-by-month revenue
- Percentage of quarter calculations
- Quarterly totals

## Key Techniques

### Creating Multiple Sheets

```swift
var writer = WorkbookWriter()
let sheet1 = writer.addSheet(named: "Summary")
let sheet2 = writer.addSheet(named: "Details")
```

### Using Formulas

```swift
// Simple formula
sheet.writeFormula("SUM(B2:B10)", to: "B11")

// Cross-sheet formula
sheet.writeFormula("SUM('Regional Sales'!B2:B5)", to: "B5")

// Calculated field
let avgFormula = "B\(row)/C\(row)"
sheet.writeFormula(avgFormula, cachedValue: revenue / units, to: ref)
```

### Structured Data Layout

```swift
// Headers
sheet.writeText("Product", to: "A1")
sheet.writeText("Revenue", to: "B1")

// Data rows
for (index, item) in data.enumerated() {
    let row = index + 2
    sheet.writeText(item.name, to: CellReference(column: "A", row: row))
    sheet.writeNumber(item.value, to: CellReference(column: "B", row: row))
}

// Totals row
sheet.writeFormula("SUM(B2:B\(data.count + 1))", to: "B\(data.count + 2)")
```

### Verification

```swift
// Read back the generated file
let workbook = try Workbook.open(url: reportURL)

// Verify content
if let sheet = try workbook.sheet(named: "Summary") {
    let value = sheet.cell(at: "B5")
    // Check value matches expectations
}
```

## Use Cases

This pattern is useful for:
- Financial reports
- Sales dashboards
- Analytics exports
- Data warehouse extracts
- Automated reporting pipelines
- Business intelligence exports
