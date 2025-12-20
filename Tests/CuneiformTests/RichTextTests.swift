import Testing
import Foundation
@testable import Cuneiform

@Suite("Rich Text Support (Phase 4.5)")
struct RichTextTests {
    
    // MARK: - TextRun Tests
    
    @Test func createPlainTextRun() {
        let run = TextRun.plain("Hello")
        #expect(run.text == "Hello")
        #expect(run.bold == false)
        #expect(run.italic == false)
        #expect(run.fontName == nil)
        #expect(run.plainText == "Hello")
    }
    
    @Test func createFormattedTextRun() {
        let run = TextRun(
            text: "Bold Red",
            fontName: "Calibri",
            fontSize: 12.0,
            color: "FF0000",
            bold: true,
            italic: false
        )
        #expect(run.text == "Bold Red")
        #expect(run.bold == true)
        #expect(run.fontName == "Calibri")
        #expect(run.fontSize == 12.0)
        #expect(run.color == "FF0000")
    }
    
    @Test func textRunEquality() {
        let run1 = TextRun.plain("text")
        let run2 = TextRun.plain("text")
        let run3 = TextRun.plain("other")
        
        #expect(run1 == run2)
        #expect(run1 != run3)
    }
    
    // MARK: - RichText Tests
    
    @Test func richTextPlainTextConcatenation() {
        let runs: RichText = [
            .plain("Hello"),
            .plain(" "),
            .plain("World")
        ]
        #expect(runs.plainText == "Hello World")
    }
    
    @Test func richTextHasFormattingDetection() {
        let plain: RichText = [.plain("text"), .plain("more")]
        #expect(plain.hasFormatting == false)
        
        let formatted: RichText = [
            .plain("text"),
            TextRun(text: "bold", bold: true)
        ]
        #expect(formatted.hasFormatting == true)
    }
    
    @Test func richTextWithMultipleFormats() {
        let runs: RichText = [
            TextRun(text: "Bold ", fontName: "Calibri", bold: true),
            TextRun(text: "Italic", fontSize: 14.0, italic: true),
            TextRun(text: " Normal")
        ]
        #expect(runs.plainText == "Bold Italic Normal")
        #expect(runs.hasFormatting == true)
    }
    
    // MARK: - SharedString Entry Tests
    
    @Test func sharedStringEntryPlain() {
        let entry: SharedStringEntry = .plain("Simple text")
        #expect(entry.plainText == "Simple text")
    }
    
    @Test func sharedStringEntryRich() {
        let runs: RichText = [.plain("Part"), TextRun(text: " Bold", bold: true)]
        let entry: SharedStringEntry = .rich(runs)
        #expect(entry.plainText == "Part Bold")
    }
    
    @Test func sharedStringEntryEquality() {
        let entry1: SharedStringEntry = .plain("text")
        let entry2: SharedStringEntry = .plain("text")
        #expect(entry1 == entry2)
    }
    
    // MARK: - CellValue Rich Text
    
    @Test func cellValueRichText() {
        let runs: RichText = [
            TextRun(text: "Hello", bold: true),
            TextRun(text: " World", italic: true)
        ]
        let cellValue = CellValue.richText(runs)
        
        if case .richText(let retreivedRuns) = cellValue {
            #expect(retreivedRuns.count == 2)
            #expect(retreivedRuns.plainText == "Hello World")
        } else {
            Issue.record("Expected richText case")
        }
    }
    
    @Test func cellValueDescription() {
        let runs: RichText = [.plain("A longer text that should be truncated in description")]
        let cellValue = CellValue.richText(runs)
        let desc = cellValue.description
        #expect(desc.contains("richText"))
        #expect(desc.contains("..."))
    }
    
    // MARK: - Parser Tests (Rich Text from XML)
    
    @Test func parseSimpleRichText() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr><b/><color rgb="FF0000"/></rPr>
              <t>Bold</t>
            </r>
            <r>
              <t> Normal</t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        
        #expect(sharedStrings.count == 1)
        #expect(sharedStrings[0] == "Bold Normal")
        
        // Check that rich text was parsed
        let richText = sharedStrings.richText(at: 0)
        #expect(richText != nil)
        #expect(richText?.count == 2)
        #expect(richText?[0].text == "Bold")
        #expect(richText?[0].bold == true)
        #expect(richText?[0].color == "FF0000")
        #expect(richText?[1].text == " Normal")
    }
    
    @Test func parseRichTextWithMultipleFormats() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr>
                <b/>
                <i/>
                <rFont val="Calibri"/>
                <sz val="12"/>
                <color rgb="FF0000"/>
              </rPr>
              <t>Formatted</t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        let richText = sharedStrings.richText(at: 0)
        
        #expect(richText?.count == 1)
        let run = richText?[0]
        #expect(run?.text == "Formatted")
        #expect(run?.bold == true)
        #expect(run?.italic == true)
        #expect(run?.fontName == "Calibri")
        #expect(run?.fontSize == 12.0)
        #expect(run?.color == "FF0000")
    }
    
    @Test func parseRichTextWithUnderlineAndStrikethrough() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr>
                <u val="double"/>
                <strike/>
              </rPr>
              <t>Text</t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        let richText = sharedStrings.richText(at: 0)
        
        let run = richText?[0]
        #expect(run?.underline == "double")
        #expect(run?.strikethrough == true)
    }
    
    @Test func parseRichTextWithThemeColor() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr>
                <color theme="1"/>
              </rPr>
              <t>Themed</t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        let richText = sharedStrings.richText(at: 0)
        
        let run = richText?[0]
        #expect(run?.themeColor == 1)
    }
    
    @Test func parsePlainTextStillWorks() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <t>Plain text</t>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        
        #expect(sharedStrings[0] == "Plain text")
        // Plain text should return nil for richText
        #expect(sharedStrings.richText(at: 0) == nil)
    }
    
    @Test func parseMixedPlainAndRichText() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si><t>Plain</t></si>
          <si>
            <r><rPr><b/></rPr><t>Bold</t></r>
          </si>
          <si><t>Another plain</t></si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        
        #expect(sharedStrings.count == 3)
        #expect(sharedStrings[0] == "Plain")
        #expect(sharedStrings.richText(at: 0) == nil)
        
        #expect(sharedStrings[1] == "Bold")
        #expect(sharedStrings.richText(at: 1) != nil)
        
        #expect(sharedStrings[2] == "Another plain")
        #expect(sharedStrings.richText(at: 2) == nil)
    }
    
    @Test func parseVerticalAlign() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr>
                <vertAlign val="superscript"/>
              </rPr>
              <t>Super</t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        let richText = sharedStrings.richText(at: 0)
        
        #expect(richText?[0].verticalAlign == "superscript")
    }
    
    // MARK: - Round-Trip Tests
    
    @Test func roundTripRichText() throws {
        // Create rich text, write it, parse it back
        let originalRuns: RichText = [
            TextRun(text: "Hello ", color: "FF0000", bold: true),
            TextRun(text: "World", fontSize: 14.0, italic: true)
        ]
        
        // Parse from sample XML
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr><b/><color rgb="FF0000"/></rPr>
              <t>Hello </t>
            </r>
            <r>
              <rPr><i/><sz val="14"/></rPr>
              <t>World</t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        let parsedRuns = sharedStrings.richText(at: 0)
        
        #expect(parsedRuns?.plainText == "Hello World")
        #expect(parsedRuns?[0].bold == true)
        #expect(parsedRuns?[0].color == "FF0000")
        #expect(parsedRuns?[1].italic == true)
        #expect(parsedRuns?[1].fontSize == 14.0)
    }
    
    @Test func emptyRichText() throws {
        let xml = """
        <?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <si>
            <r>
              <rPr><b/></rPr>
              <t></t>
            </r>
          </si>
        </sst>
        """
        
        let data = xml.data(using: .utf8)!
        let sharedStrings = try SharedStringsParser.parse(data: data)
        
        #expect(sharedStrings[0] == "")
        #expect(sharedStrings.richText(at: 0) != nil)
    }
}
