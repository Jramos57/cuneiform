import Testing
@testable import Cuneiform
import Foundation

@Suite("Phase 4.1: Full Styles Support")
struct Phase41StylesTests {
    
    // MARK: - CellFont Tests
    
    @Test("Parse fonts with bold, italic, underline, strike")
    func parseFonts() throws {
        let xml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="4">
            <font/>
            <font><b/><sz val="12"/></font>
            <font><i/><u/></font>
            <font><strike/><color rgb="FFFF0000"/></font>
          </fonts>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        
        #expect(styles.fonts.count == 4)
        #expect(styles.fonts[1].bold)
        #expect(styles.fonts[1].size == 12)
        #expect(styles.fonts[2].italic)
        #expect(styles.fonts[2].underline)
        #expect(styles.fonts[3].strike)
        #expect(styles.fonts[3].color == .rgb("FFFF0000"))
    }
    
    @Test("Parse fonts with theme colors")
    func parseFontsWithThemeColor() throws {
        let xml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1">
            <font><color theme="1"/></font>
          </fonts>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        
        #expect(styles.fonts.count == 1)
        #expect(styles.fonts[0].color == .theme(1))
    }
    
    // MARK: - CellFill Tests
    
    @Test("Parse fills with patterns and colors")
    func parseFills() throws {
        let xml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="3">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
            <fill><patternFill patternType="medGray"><fgColor rgb="FFFF0000"/><bgColor rgb="FF000000"/></patternFill></fill>
          </fills>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        
        #expect(styles.fills.count == 3)
        #expect(styles.fills[0].pattern == .none)
        #expect(styles.fills[1].pattern == .solid)
        #expect(styles.fills[1].foregroundColor == .rgb("FFFFFF00"))
        #expect(styles.fills[2].pattern == .medGray)
        #expect(styles.fills[2].foregroundColor == .rgb("FFFF0000"))
        #expect(styles.fills[2].backgroundColor == .rgb("FF000000"))
    }
    
    // MARK: - CellBorder Tests
    
    @Test("Parse borders with all sides")
    func parseBorders() throws {
        let xml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <borders count="1">
            <border>
              <left style="thin"/>
              <right style="medium"/>
              <top style="thick"/>
              <bottom style="dashed"/>
              <diagonal style="dotted"/>
            </border>
          </borders>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        
        #expect(styles.borders.count == 1)
        let border = styles.borders[0]
        #expect(border.left?.style == .thin)
        #expect(border.right?.style == .medium)
        #expect(border.top?.style == .thick)
        #expect(border.bottom?.style == .dashed)
        #expect(border.diagonal?.style == .dotted)
    }
    
    // MARK: - CellAlignment Tests
    
    @Test("Parse alignment with all properties")
    func parseAlignment() throws {
        let xml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0">
              <alignment horizontal="center" vertical="distributed" wrapText="1" textRotation="45" indent="2"/>
            </xf>
          </cellXfs>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        
        #expect(styles.cellFormats.count == 1)
        let alignment = styles.cellFormats[0].alignment
        #expect(alignment?.horizontal == .center)
        #expect(alignment?.vertical == .distributed)
        #expect(alignment?.wrapText == true)
        #expect(alignment?.textRotation == 45)
        #expect(alignment?.indent == 2)
    }
    
    // MARK: - CellStyle Query Tests
    
    @Test("Get complete cell style from format record")
    func getCellStyle() throws {
        let xml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="1">
            <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
          </numFmts>
          <fonts count="2">
            <font/>
            <font><b/><sz val="14"/></font>
          </fonts>
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <borders count="2">
            <border><left/><right/><top/><bottom/><diagonal/></border>
            <border><left style="thin"/><right/><top/><bottom/><diagonal/></border>
          </borders>
          <cellXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            <xf numFmtId="164" fontId="1" fillId="1" borderId="1"><alignment horizontal="center"/></xf>
          </cellXfs>
        </styleSheet>
        """.data(using: .utf8)!
        let styles = try StylesParser.parse(data: xml)
        
        // Query style index 1 (the custom one, which references border 1)
        let style = styles.cellStyle(forStyleIndex: 1)
        #expect(style != nil)
        #expect(style?.numberFormat?.id == 164)
        #expect(style?.numberFormat?.isDateFormat == true)
        #expect(style?.font?.bold == true)
        #expect(style?.font?.size == 14)
        #expect(style?.fill?.pattern == .solid)
        #expect(style?.fill?.foregroundColor == .rgb("FFFFFF00"))
        #expect(style?.border?.left?.style == .thin)
        #expect(style?.alignment?.horizontal == .center)
    }
    
    // MARK: - StylesBuilder Tests
    
    @Test("Build and parse CellStyle round-trip")
    func roundTripCellStyle() throws {
        var builder = StylesBuilder()
        
        // Create a complex style
        let font = CellFont(name: nil, size: 14, bold: true, italic: false, underline: false, strike: false, color: .rgb("FFFF0000"))
        let fill = CellFill(pattern: .solid, foregroundColor: .rgb("FFFFFF00"))
        let borderSide = CellBorderSide(style: .medium, color: .rgb("FF000000"))
        let border = CellBorder(left: borderSide, right: borderSide, top: borderSide, bottom: borderSide)
        let alignment = CellAlignment(horizontal: .center, vertical: .center, wrapText: true)
        
        let style = CellStyle(
            font: font,
            fill: fill,
            border: border,
            alignment: alignment
        )
        
        let styleIndex = builder.addCellStyle(style)
        
        // Build and parse back
        let xmlData = builder.build()
        let parsed = try StylesParser.parse(data: xmlData)
        
        // Verify we can retrieve the style
        let retrievedStyle = parsed.cellStyle(forStyleIndex: styleIndex)
        #expect(retrievedStyle != nil)
        #expect(retrievedStyle?.font?.bold == true)
        #expect(retrievedStyle?.font?.size == 14)
        #expect(retrievedStyle?.fill?.pattern == .solid)
        #expect(retrievedStyle?.alignment?.horizontal == .center)
        #expect(retrievedStyle?.alignment?.wrapText == true)
    }
    
    @Test("Multiple styles with deduplication")
    func multipleStylesWithDedup() throws {
        var builder = StylesBuilder()
        
        let boldFont = CellFont(bold: true)
        let yellowFill = CellFill(pattern: .solid, foregroundColor: .rgb("FFFFFF00"))
        
        // Add different styles
        let idx1 = builder.addCellStyle(CellStyle(font: boldFont))
        let idx2 = builder.addCellStyle(CellStyle(fill: yellowFill))
        
        // Both should be added with different indices
        #expect(idx1 != idx2)
        
        // The indices should be >= 1 (after the default format at 0)
        #expect(idx1 >= 1)
        #expect(idx2 >= 1)
    }
}
