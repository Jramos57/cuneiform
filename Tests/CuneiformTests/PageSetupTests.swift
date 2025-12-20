import Testing
import Foundation
@testable import Cuneiform

@Suite("Page Setup Tests (Phase 4.7)")
struct PageSetupTests {
    
    // MARK: - Page Setup Type Tests
    
    @Test func createDefaultPageSetup() {
        let setup = PageSetup.default
        #expect(setup.orientation == .portrait)
        #expect(setup.paperSize == .letter)
        #expect(setup.scale == 100)
        #expect(setup.fitToPages == nil)
        #expect(setup.printQuality == 300)
        #expect(setup.firstPageNumber == nil)
    }
    
    @Test func createLandscapePageSetup() {
        let setup = PageSetup(orientation: .landscape)
        #expect(setup.orientation == .landscape)
    }
    
    @Test func createFitToPageSetup() {
        let fitTo = FitToPages(width: 2, height: 3)
        let setup = PageSetup(scale: nil, fitToPages: fitTo)
        #expect(setup.fitToPages?.width == 2)
        #expect(setup.fitToPages?.height == 3)
        #expect(setup.scale == nil)
    }
    
    @Test func customPaperSizes() {
        let a4 = PageSetup(paperSize: .a4)
        #expect(a4.paperSize == .a4)
        #expect(a4.paperSize.rawValue == 9)
        
        let legal = PageSetup(paperSize: .legal)
        #expect(legal.paperSize == .legal)
        #expect(legal.paperSize.rawValue == 5)
    }
    
    @Test func customMargins() {
        let margins = PageSetup.Margins(
            left: 1.0, right: 1.5,
            top: 0.8, bottom: 0.8,
            header: 0.3, footer: 0.3
        )
        let setup = PageSetup(margins: margins)
        #expect(setup.margins.left == 1.0)
        #expect(setup.margins.right == 1.5)
        #expect(setup.margins.top == 0.8)
        #expect(setup.margins.bottom == 0.8)
        #expect(setup.margins.header == 0.3)
        #expect(setup.margins.footer == 0.3)
    }
    
    @Test func pageSetupEquatable() {
        let setup1 = PageSetup(orientation: .landscape)
        let setup2 = PageSetup(orientation: .landscape)
        let setup3 = PageSetup(orientation: .portrait)
        
        #expect(setup1 == setup2)
        #expect(setup1 != setup3)
    }
    
    // MARK: - Parser Tests
    
    @Test func parseBasicPageSetup() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
            <pageSetup orientation="landscape" paperSize="9"/>
            <pageMargins left="0.75" right="0.75" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>
        </worksheet>
        """
        
        let data = xml.data(using: .utf8)!
        let worksheet = try WorksheetParser.parse(data: data)
        
        #expect(worksheet.pageSetup != nil)
        #expect(worksheet.pageSetup?.orientation == .landscape)
        #expect(worksheet.pageSetup?.paperSize == .a4)
    }
    
    @Test func parsePageMarginsOnly() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
            <pageMargins left="1.0" right="1.5" top="0.8" bottom="0.8" header="0.3" footer="0.3"/>
        </worksheet>
        """
        
        let data = xml.data(using: .utf8)!
        let worksheet = try WorksheetParser.parse(data: data)
        
        #expect(worksheet.pageSetup != nil)
        #expect(worksheet.pageSetup?.margins.left == 1.0)
        #expect(worksheet.pageSetup?.margins.right == 1.5)
        #expect(worksheet.pageSetup?.margins.header == 0.3)
        #expect(worksheet.pageSetup?.margins.footer == 0.3)
    }
    
    @Test func parseFitToPages() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
            <pageSetup fitToWidth="2" fitToHeight="3"/>
        </worksheet>
        """
        
        let data = xml.data(using: .utf8)!
        let worksheet = try WorksheetParser.parse(data: data)
        
        #expect(worksheet.pageSetup?.fitToPages?.width == 2)
        #expect(worksheet.pageSetup?.fitToPages?.height == 3)
    }
    
    @Test func parseScaleOnly() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
            <pageSetup scale="50"/>
        </worksheet>
        """
        
        let data = xml.data(using: .utf8)!
        let worksheet = try WorksheetParser.parse(data: data)
        
        #expect(worksheet.pageSetup?.scale == 50)
        #expect(worksheet.pageSetup?.fitToPages == nil)
    }
    
    @Test func parseFirstPageNumber() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
            <pageSetup firstPageNumber="5"/>
        </worksheet>
        """
        
        let data = xml.data(using: .utf8)!
        let worksheet = try WorksheetParser.parse(data: data)
        
        #expect(worksheet.pageSetup?.firstPageNumber == 5)
    }
    
    // MARK: - API Tests
    
    @Test func sheetPageSetupAccessor() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/>
            <pageSetup orientation="landscape" paperSize="5"/>
        </worksheet>
        """
        
        let data = xml.data(using: .utf8)!
        let worksheet = try WorksheetParser.parse(data: data)
        
        let sheet = Sheet(
            data: worksheet,
            sharedStrings: SharedStrings.empty,
            styles: StylesInfo(numberFormats: [:], fonts: [], fills: [], borders: [], cellFormats: [])
        )
        #expect(sheet.pageSetup != nil)
        #expect(sheet.pageSetup?.orientation == .landscape)
        #expect(sheet.pageSetup?.paperSize == .legal)
    }
    
    // MARK: - Write-side Tests
    
    @Test func writePageSetup() throws {
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Test")
        
        var sheet = writer.sheet(at: sheetIdx)!
        let pageSetup = PageSetup(
            orientation: .landscape,
            paperSize: .a4,
            scale: 75,
            margins: PageSetup.Margins(left: 1.0, right: 1.0, top: 0.8, bottom: 0.8, header: 0.4, footer: 0.4)
        )
        sheet.setPageSetup(pageSetup)
        sheet.writeText("Test", to: CellReference(column: "A", row: 1))
        
        writer.modifySheet(at: sheetIdx) { $0 = sheet }
        
        let data = try writer.buildData()
        
        // Verify by reading back
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "Test")!
        
        #expect(readSheet.pageSetup != nil)
        #expect(readSheet.pageSetup?.orientation == .landscape)
        #expect(readSheet.pageSetup?.paperSize == .a4)
        #expect(readSheet.pageSetup?.scale == 75)
        
        // Cleanup
        try FileManager.default.removeItem(at: tempURL)
    }
    
    @Test func writeFitToPages() throws {
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "FitTest")
        
        var sheet = writer.sheet(at: sheetIdx)!
        let fitTo = FitToPages(width: 1, height: 2)
        let pageSetup = PageSetup(scale: nil, fitToPages: fitTo)
        sheet.setPageSetup(pageSetup)
        sheet.writeText("Fit", to: CellReference(column: "A", row: 1))
        
        writer.modifySheet(at: sheetIdx) { $0 = sheet }
        
        let data = try writer.buildData()
        
        // Verify
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "FitTest")!
        
        #expect(readSheet.pageSetup?.fitToPages?.width == 1)
        #expect(readSheet.pageSetup?.fitToPages?.height == 2)
        #expect(readSheet.pageSetup?.scale == nil)
        
        // Cleanup
        try FileManager.default.removeItem(at: tempURL)
    }
    
    @Test func multipleSheetPageSetups() throws {
        var writer = WorkbookWriter()
        let sheet1Idx = writer.addSheet(named: "Landscape")
        let sheet2Idx = writer.addSheet(named: "Portrait")
        
        // Sheet 1: Landscape
        var sheet1 = writer.sheet(at: sheet1Idx)!
        sheet1.setPageSetup(PageSetup(orientation: .landscape))
        sheet1.writeText("A", to: CellReference(column: "A", row: 1))
        writer.modifySheet(at: sheet1Idx) { $0 = sheet1 }
        
        // Sheet 2: Portrait (default)
        var sheet2 = writer.sheet(at: sheet2Idx)!
        sheet2.setPageSetup(PageSetup(orientation: .portrait))
        sheet2.writeText("B", to: CellReference(column: "A", row: 1))
        writer.modifySheet(at: sheet2Idx) { $0 = sheet2 }
        
        let data = try writer.buildData()
        
        // Verify
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet1 = try workbook.sheet(named: "Landscape")!
        let readSheet2 = try workbook.sheet(named: "Portrait")!
        
        #expect(readSheet1.pageSetup?.orientation == .landscape)
        #expect(readSheet2.pageSetup?.orientation == .portrait)
        
        // Cleanup
        try FileManager.default.removeItem(at: tempURL)
    }
    
    @Test func pageSetupWithProtection() throws {
        var writer = WorkbookWriter()
        let sheetIdx = writer.addSheet(named: "Protected")
        
        var sheet = writer.sheet(at: sheetIdx)!
        sheet.setPageSetup(PageSetup(orientation: .landscape))
        var options = SheetProtectionOptions()
        options.deleteColumns = false
        sheet.protectSheet(options: options)
        sheet.writeText("Protected", to: CellReference(column: "A", row: 1))
        
        writer.modifySheet(at: sheetIdx) { $0 = sheet }
        
        let data = try writer.buildData()
        
        // Verify both are present
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "Protected")!
        
        #expect(readSheet.pageSetup != nil)
        #expect(readSheet.pageSetup?.orientation == .landscape)
        #expect(readSheet.protection != nil)
        
        // Cleanup
        try FileManager.default.removeItem(at: tempURL)
    }
    
    @Test func marginDefaults() {
        let margins = PageSetup.Margins()
        #expect(margins.left == 0.75)
        #expect(margins.right == 0.75)
        #expect(margins.top == 1.0)
        #expect(margins.bottom == 1.0)
        #expect(margins.header == 0.5)
        #expect(margins.footer == 0.5)
    }
    
    @Test func paperSizeValues() {
        #expect(PageSetup.PaperSize.letter.rawValue == 1)
        #expect(PageSetup.PaperSize.a4.rawValue == 9)
        #expect(PageSetup.PaperSize.legal.rawValue == 5)
        #expect(PageSetup.PaperSize.a3.rawValue == 8)
    }
}
