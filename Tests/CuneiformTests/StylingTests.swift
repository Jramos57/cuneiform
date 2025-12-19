import Foundation
import Testing

@testable import Cuneiform

@Suite("Styling and Formatting")
struct StylingTests {
    
    // MARK: - StylesBuilder Tests
    
    @Test("StylesBuilder: Create default styles")
    func defaultStyles() {
        var styles = StylesBuilder()
        let data = styles.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("<styleSheet"))
        #expect(xml.contains("</styleSheet>"))
        #expect(xml.contains("<fonts count="))
        #expect(xml.contains("<fills count="))
    }
    
    @Test("StylesBuilder: Add number format")
    func addNumberFormat() {
        var styles = StylesBuilder()
        let dateIdx = styles.addNumberFormat("yyyy-mm-dd")
        let currencyIdx = styles.addNumberFormat("\"$\"#,##0.00")
        
        #expect(dateIdx == 164)
        #expect(currencyIdx == 165)
        
        // Deduplication
        let dateIdx2 = styles.addNumberFormat("yyyy-mm-dd")
        #expect(dateIdx2 == dateIdx)
    }
    
    @Test("StylesBuilder: Add font")
    func addFont() {
        var styles = StylesBuilder()
        let boldIdx = styles.addFont(StylesBuilder.Font(bold: true))
        let italicIdx = styles.addFont(StylesBuilder.Font(italic: true))
        
        #expect(boldIdx > 0)
        #expect(italicIdx > 0)
        #expect(boldIdx != italicIdx)
        
        // Deduplication
        let boldIdx2 = styles.addFont(StylesBuilder.Font(bold: true))
        #expect(boldIdx2 == boldIdx)
    }
    
    @Test("StylesBuilder: Add fill")
    func addFill() {
        var styles = StylesBuilder()
        let redIdx = styles.addFill(StylesBuilder.Fill(patternType: "solid", fgColor: "FFFF0000"))
        let blueIdx = styles.addFill(StylesBuilder.Fill(patternType: "solid", fgColor: "FF0000FF"))
        
        #expect(redIdx > 1)  // At least 2 default fills
        #expect(blueIdx > redIdx)
        
        // Deduplication
        let redIdx2 = styles.addFill(StylesBuilder.Fill(patternType: "solid", fgColor: "FFFF0000"))
        #expect(redIdx2 == redIdx)
    }
    
    @Test("StylesBuilder: Add border")
    func addBorder() {
        var styles = StylesBuilder()
        let border = StylesBuilder.Border(
            left: ("medium", "FF000000"),
            bottom: ("medium", "FF000000")
        )
        let borderIdx = styles.addBorder(border)
        
        #expect(borderIdx > 0)
    }
    
    @Test("StylesBuilder: Add cell format")
    func addCellFormat() {
        var styles = StylesBuilder()
        let format = StylesBuilder.CellFormat(
            numFmtId: 14,  // Built-in date format
            fontId: 0,
            fillId: 0,
            borderId: 0
        )
        let formatIdx = styles.addCellFormat(format)
        
        #expect(formatIdx > 0)
    }
    
    @Test("StylesBuilder: Convenience date format")
    func dateFormatConvenience() {
        var styles = StylesBuilder()
        let dateIdx = styles.addDateFormat("dd/mm/yyyy")
        
        #expect(dateIdx > 0)
        
        let xml = String(data: styles.build(), encoding: .utf8)!
        #expect(xml.contains("dd/mm/yyyy"))
    }
    
    // MARK: - Write + Style Tests
    
    @Test("WorkbookWriter: Write styled text")
    func writeStyledText() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("styled_\(UUID().uuidString).xlsx")
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Styled")
        
        // Add styles
        var boldIdx = 0
        writer.style { styles in
            let boldFont = styles.addFont(StylesBuilder.Font(bold: true))
            let boldFormat = StylesBuilder.CellFormat(fontId: boldFont)
            boldIdx = styles.addCellFormat(boldFormat)
        }
        
        writer.modifySheet(at: sheetIdx) { sheet in
            sheet.writeText("Normal", to: CellReference(column: "A", row: 1))
            sheet.writeText("Bold", to: CellReference(column: "B", row: 1), styleIndex: boldIdx)
        }
        
        try writer.save(to: tempFile)
        
        // Verify file exists and has correct size
        let attrs = try FileManager.default.attributesOfItem(atPath: tempFile.path)
        let fileSize = attrs[.size] as! Int
        #expect(fileSize > 1000)  // Should have some content
    }
    
    @Test("WorkbookWriter: Write with date format")
    func writeDateFormat() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("dates_\(UUID().uuidString).xlsx")
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Dates")
        
        writer.style { styles in
            _ = styles.addDateFormat("yyyy-mm-dd")
        }
        
        writer.modifySheet(at: sheetIdx) { sheet in
            sheet.writeNumber(44927, to: CellReference(column: "A", row: 1), styleIndex: 1)  // Excel serial for a date
            sheet.writeNumber(44928, to: CellReference(column: "A", row: 2), styleIndex: 1)
        }
        
        try writer.save(to: tempFile)
        
        // Verify file was created
        #expect(FileManager.default.fileExists(atPath: tempFile.path))
        
        // Round-trip: read back
        let workbook = try Workbook.open(url: tempFile)
        let maybeSheet = try workbook.sheet(named: "Dates")
        let sheet = try #require(maybeSheet)
        
        let cell1 = try #require(sheet.cell(at: CellReference(column: "A", row: 1)))
        switch cell1 {
        case .date(let s):
            #expect(Double(s) == 44927)
        default:
            #expect(false)
        }
    }
    
    @Test("WorkbookWriter: Write with colors and borders")
    func writeColorsAndBorders() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("styled_cells_\(UUID().uuidString).xlsx")
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Colors")
        
        var redIdx = 0
        var boldBlueIdx = 0
        writer.style { styles in
            // Red background
            let redFill = styles.addFill(StylesBuilder.Fill(patternType: "solid", fgColor: "FFFF0000"))
            let redFormat = StylesBuilder.CellFormat(fillId: redFill)
            redIdx = styles.addCellFormat(redFormat)
            
            // Bold + Blue background
            let blueFill = styles.addFill(StylesBuilder.Fill(patternType: "solid", fgColor: "FF0000FF"))
            let boldFont = styles.addFont(StylesBuilder.Font(bold: true))
            let boldBlueFormat = StylesBuilder.CellFormat(fontId: boldFont, fillId: blueFill)
            boldBlueIdx = styles.addCellFormat(boldBlueFormat)
        }
        
        writer.modifySheet(at: sheetIdx) { sheet in
            sheet.writeText("Red Background", to: CellReference(column: "A", row: 1), styleIndex: redIdx)
            sheet.writeText("Bold on Blue", to: CellReference(column: "A", row: 2), styleIndex: boldBlueIdx)
        }
        
        try writer.save(to: tempFile)
        
        #expect(FileManager.default.fileExists(atPath: tempFile.path))
        
        // Verify round-trip
        let workbook = try Workbook.open(url: tempFile)
        let maybeSheet = try workbook.sheet(named: "Colors")
        let sheet = try #require(maybeSheet)
        
        let cell1 = try #require(sheet.cell(at: CellReference(column: "A", row: 1)))
        switch cell1 {
        case .text(let s): #expect(s == "Red Background")
        default: #expect(false)
        }
        
        let cell2 = try #require(sheet.cell(at: CellReference(column: "A", row: 2)))
        switch cell2 {
        case .text(let s): #expect(s == "Bold on Blue")
        default: #expect(false)
        }
    }
    
    @Test("WorkbookWriter: Multiple styles in single sheet")
    func multipleStylesSheet() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("multi_style_\(UUID().uuidString).xlsx")
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "MultiStyle")
        
        var dateStyleIdx = 0
        var currencyStyleIdx = 0
        var headerStyleIdx = 0
        writer.style { styles in
            dateStyleIdx = styles.addDateFormat("yyyy-mm-dd")
            let currencyFmtId = styles.addNumberFormat("\"$\"#,##0.00")
            currencyStyleIdx = styles.addCellFormat(StylesBuilder.CellFormat(numFmtId: currencyFmtId))
            let boldFont = styles.addFont(StylesBuilder.Font(bold: true))
            headerStyleIdx = styles.addCellFormat(StylesBuilder.CellFormat(fontId: boldFont))
        }
        
        writer.modifySheet(at: sheetIdx) { sheet in
            // Header in bold
            sheet.writeText("Date", to: CellReference(column: "A", row: 1), styleIndex: headerStyleIdx)
            sheet.writeText("Amount", to: CellReference(column: "B", row: 1), styleIndex: headerStyleIdx)
            
            // Data with formats
            sheet.writeNumber(44927, to: CellReference(column: "A", row: 2), styleIndex: dateStyleIdx)
            sheet.writeNumber(1234.56, to: CellReference(column: "B", row: 2), styleIndex: currencyStyleIdx)
            
            sheet.writeNumber(44928, to: CellReference(column: "A", row: 3), styleIndex: dateStyleIdx)
            sheet.writeNumber(2345.67, to: CellReference(column: "B", row: 3), styleIndex: currencyStyleIdx)
        }
        
        try writer.save(to: tempFile)
        
        // Verify
        let workbook = try Workbook.open(url: tempFile)
        let maybeSheet = try workbook.sheet(named: "MultiStyle")
        let sheet = try #require(maybeSheet)
        
        let dateCell = try #require(sheet.cell(at: CellReference(column: "A", row: 2)))
        switch dateCell {
        case .date(let s): #expect(Double(s) == 44927)
        default: #expect(false)
        }
        
        let amountCell = try #require(sheet.cell(at: CellReference(column: "B", row: 2)))
        switch amountCell {
        case .number(let n): #expect(n == 1234.56)
        default: #expect(false)
        }
    }
    
    @Test("StylesParser: Parse existing styles with date formats")
    func parseExistingStyles() throws {
        // Use a test file if available, or create minimal styles XML
        let stylesXML = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1">
        <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
        </numFmts>
        <fonts count="1"><font/></fonts>
        <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
        </cellXfs>
        </styleSheet>
        """
        
        let data = stylesXML.data(using: .utf8)!
        let styles = try StylesParser.parse(data: data)
        
        #expect(styles.cellFormats.count == 2)
        #expect(styles.numberFormats.count == 1)
        
        // Check date format detection
        let isDate = styles.isDateFormat(styleIndex: 1)
        #expect(isDate)
    }
    
    @Test("Round-trip: Write, read, verify date format")
    func roundTripDateFormat() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("roundtrip_date_\(UUID().uuidString).xlsx")
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        // Write
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Dates")
        
        writer.style { styles in
            _ = styles.addDateFormat("mm/dd/yyyy")
        }
        
        writer.modifySheet(at: sheetIdx) { sheet in
            sheet.writeNumber(45000, to: CellReference(column: "A", row: 1), styleIndex: 1)
            sheet.writeText("Header", to: CellReference(column: "B", row: 1))
            sheet.writeNumber(45001, to: CellReference(column: "A", row: 2), styleIndex: 1)
        }
        
        try writer.save(to: tempFile)
        
        // Read back
        let workbook = try Workbook.open(url: tempFile)
        let maybeSheet = try workbook.sheet(named: "Dates")
        let sheet = try #require(maybeSheet)
        
        let cell1 = try #require(sheet.cell(at: CellReference(column: "A", row: 1)))
        switch cell1 {
        case .date(let s): #expect(Double(s) == 45000)
        default: #expect(false)
        }
        
        let header = try #require(sheet.cell(at: CellReference(column: "B", row: 1)))
        switch header {
        case .text(let s): #expect(s == "Header")
        default: #expect(false)
        }
        
        let cell2 = try #require(sheet.cell(at: CellReference(column: "A", row: 2)))
        switch cell2 {
        case .date(let s): #expect(Double(s) == 45001)
        default: #expect(false)
        }
    }
    
    @Test("Styling: Large dataset with mixed formats")
    func largeDatasetMixedFormats() throws {
        let tempFile = FileManager.default.temporaryDirectory
            .appendingPathComponent("large_styled_\(UUID().uuidString).xlsx")
        
        defer { try? FileManager.default.removeItem(at: tempFile) }
        
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Data")
        
        var dateStyleIdx = 0
        var percentStyleIdx = 0
        var headerStyleIdx = 0
        writer.style { styles in
            dateStyleIdx = styles.addDateFormat("yyyy-mm-dd")
            let percentFmtId = styles.addNumberFormat("#0.00%")
            percentStyleIdx = styles.addCellFormat(StylesBuilder.CellFormat(numFmtId: percentFmtId))
            let boldFont = styles.addFont(StylesBuilder.Font(bold: true, size: 12))
            headerStyleIdx = styles.addCellFormat(StylesBuilder.CellFormat(fontId: boldFont))
        }
        
        writer.modifySheet(at: sheetIdx) { sheet in
            // Headers (bold, size 12)
            sheet.writeText("Date", to: CellReference(column: "A", row: 1), styleIndex: headerStyleIdx)
            sheet.writeText("Percentage", to: CellReference(column: "B", row: 1), styleIndex: headerStyleIdx)
            
            // 100 rows of data
            for i in 2...101 {
                let dateSerial = Double(44927 + (i - 2))
                let percentage = Double(i) * 0.01
                
                sheet.writeNumber(dateSerial, to: CellReference(column: "A", row: i), styleIndex: dateStyleIdx)
                sheet.writeNumber(percentage, to: CellReference(column: "B", row: i), styleIndex: percentStyleIdx)
            }
        }
        
        try writer.save(to: tempFile)
        
        // Verify
        let workbook = try Workbook.open(url: tempFile)
        let maybeSheet = try workbook.sheet(named: "Data")
        let sheet = try #require(maybeSheet)
        
        let cell = try #require(sheet.cell(at: CellReference(column: "A", row: 50)))
        switch cell {
        case .date(let s): #expect(Double(s) == 44975)
        default: #expect(false)
        }
    }
}
