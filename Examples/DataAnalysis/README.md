# Data Analysis Example

This example demonstrates how to use Cuneiform for analyzing data in Excel files.

## What This Example Shows

- Opening and reading Excel files
- Extracting data from specific columns
- Computing statistics (sum, average, min, max)
- Finding cells matching criteria
- Filtering rows by conditions
- Processing ranges
- Memory-efficient streaming iteration

## Running the Example

```bash
cd Examples/DataAnalysis
swift run
```

## Example Output

```
📊 Cuneiform Data Analysis Example

✅ Opened workbook with 1 sheet(s)
📄 Sheet: no dimension

📈 Extracting sales data from column B...
   Found 10 sales records

📊 Sales Statistics:
   Total Sales: $5455.50
   Average Sale: $545.55
   Minimum Sale: $310.25
   Maximum Sale: $890.75

💰 High-Value Sales (> $500):
   Widget B: $675.25
   Widget D: $890.75
   Widget E: $550.00
   Widget G: $725.00
   Widget I: $625.50

🌎 Sales by Region (West):
   Widget B: $675.25
   Widget D: $890.75
   Widget G: $725.00
   West Total: $2291.00

📋 Processing range A1:C5:
   15 cells in range
   Header: A1 = Product
   Header: B1 = Sales
   Header: C1 = Region

🔄 Streaming all rows (memory-efficient):
   Processed 12 rows using lazy iteration

✅ Analysis complete!
```

## Key Concepts

### Column Access

```swift
let salesColumn = sheet.column("B")
for (row, value) in salesColumn where row > 1 {
    if case .number(let amount) = value {
        salesValues.append(amount)
    }
}
```

### Finding Cells

```swift
let highValueSales = sheet.findAll { ref, value in
    if case .number(let amount) = value, amount > 500 {
        return true
    }
    return false
}
```

### Row Filtering

```swift
let westSales = sheet.rows { cells in
    cells.contains { cell in
        cell.reference.column == "C" && cell.value == .text("West")
    }
}
```

### Streaming Iteration

```swift
for row in sheet.rows() {
    // Process each row lazily
    for (ref, value) in row {
        process(ref, value)
    }
}
```
