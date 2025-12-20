import Testing
import Foundation
@testable import Cuneiform

@Suite("Shared Strings Optimization (Phase 4.6)")
struct SharedStringsOptimizationTests {
    
    // MARK: - Builder Tests
    
    @Test func buildSimpleStrings() {
        var builder = SharedStringsBuilder()
        let idx1 = builder.addString("Hello")
        let idx2 = builder.addString("World")
        let idx3 = builder.addString("Hello")  // Duplicate
        
        #expect(idx1 == 0)
        #expect(idx2 == 1)
        #expect(idx3 == 0)  // Should return same index as first "Hello"
        #expect(builder.count == 2)
    }
    
    @Test func buildRichText() {
        var builder = SharedStringsBuilder()
        let runs: RichText = [
            TextRun(text: "Hello", bold: true),
            TextRun(text: " World", italic: true)
        ]
        let idx = builder.addRichText(runs)
        #expect(idx == 0)
        #expect(builder.count == 1)
    }
    
    @Test func buildMixedContent() {
        var builder = SharedStringsBuilder()
        let plainIdx = builder.addString("Plain")
        
        let richRuns: RichText = [
            TextRun(text: "Rich", bold: true)
        ]
        let richIdx = builder.addRichText(richRuns)
        
        #expect(plainIdx == 0)
        #expect(richIdx == 1)
        #expect(builder.count == 2)
    }
    
    @Test func buildXMLSimple() throws {
        var builder = SharedStringsBuilder()
        builder.addString("Hello")
        builder.addString("World")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("<si><t>Hello</t></si>"))
        #expect(xml.contains("<si><t>World</t></si>"))
        #expect(xml.contains("count=\"2\""))
        #expect(xml.contains("uniqueCount=\"2\""))
    }
    
    @Test func buildXMLWithRichText() throws {
        var builder = SharedStringsBuilder()
        
        let runs: RichText = [
            TextRun(text: "Bold", color: "FF0000", bold: true),
            TextRun(text: " Normal")
        ]
        builder.addRichText(runs)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("<r>"))
        #expect(xml.contains("<rPr>"))
        #expect(xml.contains("<b/>"))
        #expect(xml.contains("rgb=\"FF0000\""))
        #expect(xml.contains("<t>Bold</t>"))
    }
    
    @Test func buildXMLWithRichTextMultipleRuns() throws {
        var builder = SharedStringsBuilder()
        
        let runs: RichText = [
            TextRun(text: "Bold", bold: true),
            TextRun(text: " Italic", italic: true),
            TextRun(text: " Normal")
        ]
        builder.addRichText(runs)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        // Should have 3 <r> elements
        let runCount = xml.components(separatedBy: "<r>").count - 1
        #expect(runCount == 3)
    }
    
    @Test func buildXMLWithEscaping() throws {
        var builder = SharedStringsBuilder()
        builder.addString("Text with <special> & \"characters\"")
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("&lt;special&gt;"))
        #expect(xml.contains("&amp;"))
        #expect(xml.contains("&quot;"))
    }
    
    @Test func buildXMLRichTextWithThemeColor() throws {
        var builder = SharedStringsBuilder()
        
        let runs: RichText = [
            TextRun(text: "Themed", themeColor: 1)
        ]
        builder.addRichText(runs)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("theme=\"1\""))
    }
    
    @Test func buildXMLRichTextWithUnderlineAndStrike() throws {
        var builder = SharedStringsBuilder()
        
        let runs: RichText = [
            TextRun(text: "Special", underline: "double", strikethrough: true)
        ]
        builder.addRichText(runs)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("u val=\"double\""))
        #expect(xml.contains("<strike/>"))
    }
    
    @Test func buildXMLRichTextVerticalAlign() throws {
        var builder = SharedStringsBuilder()
        
        let runs: RichText = [
            TextRun(text: "Super", verticalAlign: "superscript")
        ]
        builder.addRichText(runs)
        
        let data = builder.build()
        let xml = String(data: data, encoding: .utf8)!
        
        #expect(xml.contains("vertAlign val=\"superscript\""))
    }
    
    // MARK: - Integration Tests
    
    @Test func roundTripPlainStrings() throws {
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Sheet1")
        
        var sheet = writer.sheet(at: sheetIdx)!
        sheet.writeText("Hello", to: "A1")
        sheet.writeText("World", to: "A2")
        sheet.writeText("Hello", to: "A3")  // Duplicate
        
        writer.modifySheet(at: sheetIdx) { $0 = sheet }
        
        let data = try writer.buildData()
        
        // Parse back
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "Sheet1")!
        
        if case .text(let val) = readSheet.cell(at: "A1") {
            #expect(val == "Hello")
        } else {
            let actual = String(describing: readSheet.cell(at: "A1") ?? .empty)
            Issue.record("Expected text value at A1; got: \(actual)")
        }
        
        if case .text(let val) = readSheet.cell(at: "A2") {
            #expect(val == "World")
        } else {
            let actual = String(describing: readSheet.cell(at: "A2") ?? .empty)
            Issue.record("Expected text value at A2; got: \(actual)")
        }
        
        if case .text(let val) = readSheet.cell(at: "A3") {
            #expect(val == "Hello")
        } else {
            let actual = String(describing: readSheet.cell(at: "A3") ?? .empty)
            Issue.record("Expected text value at A3; got: \(actual)")
        }
    }
    
    @Test func workbookWriterBuildsSharedStrings() throws {
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Test")
        
        var sheet = writer.sheet(at: sheetIdx)!
        sheet.writeText("String1", to: CellReference(column: "A", row: 1))
        sheet.writeText("String2", to: CellReference(column: "A", row: 2))
        
        writer.modifySheet(at: sheetIdx) { $0 = sheet }
        
        let data = try writer.buildData()
            let pkg = try OPCPackage.open(data: data)
            #expect(pkg.partExists(.sharedStrings) == true)
            // Read the first worksheet part directly via well-known path
            let sheet1 = try pkg.readPart(PartPath("/xl/worksheets/sheet1.xml"))
            let xml = String(data: sheet1, encoding: .utf8)!
            #expect(xml.contains("t=\"s\""))
            // Inspect sharedStrings.xml directly
            let ssData = try pkg.readPart(.sharedStrings)
            let ssXml = String(data: ssData, encoding: .utf8)!
            #expect(ssXml.contains("<sst"))
            let parsedSS = try SharedStringsParser.parse(data: ssData)
            #expect(parsedSS.count > 0)
        
        // Verify we can read back the written strings
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "Test")!
        
        if case .text(let val) = readSheet.cell(at: "A1") {
            #expect(val == "String1")
        } else {
            let actual = String(describing: readSheet.cell(at: "A1") ?? .empty)
            Issue.record("Expected text value at A1; got: \(actual)")
        }
        
        if case .text(let val) = readSheet.cell(at: "A2") {
            #expect(val == "String2")
        } else {
            let actual = String(describing: readSheet.cell(at: "A2") ?? .empty)
            Issue.record("Expected text value at A2; got: \(actual)")
        }
    }
    
    @Test func largeFileStringOptimization() throws {
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Data")
        
        var sheet = writer.sheet(at: sheetIdx)!
        
        // Write 100 cells with repeated strings
        for row in 1...100 {
            let colRef = CellReference(column: "A", row: row)
            if row % 3 == 0 {
                sheet.writeText("Repeated", to: colRef)
            } else {
                sheet.writeText("Value\(row)", to: colRef)
            }
        }
        
        writer.modifySheet(at: sheetIdx) { $0 = sheet }
        
        let data = try writer.buildData()
        
        // Verify we can read back the data
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "Data")!
        
        // Verify first repeated string
        if case .text(let val) = readSheet.cell(at: "A3") {
            #expect(val == "Repeated")
        } else {
            let actual = String(describing: readSheet.cell(at: "A3") ?? .empty)
            Issue.record("Expected text value at A3; got: \(actual)")
        }
        
        // Verify that we have many rows (bulk read verification)
        #expect(readSheet.rowCount > 50)
    }
    
    @Test func emptyWorkbookNoSharedStrings() throws {
        var writer = WorkbookWriter()
        writer.addSheet(named: "Empty")
        
        let data = try writer.buildData()
        
        // Verify sharedStrings.xml doesn't exist if no strings were added
        let pkg = try OPCPackage.open(data: data)
        let hasSharedStrings = pkg.partExists(.sharedStrings)
        #expect(hasSharedStrings == false)
    }
}
