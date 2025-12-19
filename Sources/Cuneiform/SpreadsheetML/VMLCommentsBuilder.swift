import Foundation

/// Builds a minimal VML drawing for worksheet comments so they display in Excel UI.
struct VMLCommentsBuilder {
    /// A single comment anchor description
    struct Anchor {
        let column: Int
        let row: Int
    }

    private let entries: [CommentsBuilder.Entry]

    init(entries: [CommentsBuilder.Entry]) {
        self.entries = entries
    }

    /// Build VML drawing data for the provided comments.
    func build() -> Data {
        var xml = """
        <?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
        <xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">
          <o:shapelayout v:ext=\"edit\">
            <o:idmap v:ext=\"edit\" data=\"1\"/>
          </o:shapelayout>
          <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">
            <v:stroke joinstyle=\"miter\"/>
            <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>
          </v:shapetype>
        """

        var shapeIndex = 1025
        for entry in entries {
            let anchor = anchorFor(entry.reference)
            xml += shapeXML(id: shapeIndex, anchor: anchor)
            shapeIndex += 1
        }

        xml += "\n</xml>\n"
        return xml.data(using: .utf8)!
    }

    // MARK: - Helpers

    private func anchorFor(_ ref: CellReference) -> Anchor {
        // VML anchor rows/cols are zero-based; keep a small offset.
        Anchor(column: ref.columnIndex, row: ref.row - 1)
    }

    private func shapeXML(id: Int, anchor: Anchor) -> String {
        let anchorString = vmlAnchor(column: anchor.column, row: anchor.row)
        return """
          <v:shape id=\"_x0000_s\(id)\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:108pt;margin-top:80pt;width:108pt;height:59.25pt;z-index:\(id - 1000);visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">
            <v:fill color2=\"#ffffe1\"/>
            <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>
            <v:path o:connecttype=\"none\"/>
            <v:textbox style=\"mso-direction-alt:auto\">
              <div style=\"text-align:left\"></div>
            </v:textbox>
            <x:ClientData ObjectType=\"Note\">
              <x:MoveWithCells/>
              <x:SizeWithCells/>
              <x:Anchor>\(anchorString)</x:Anchor>
              <x:AutoFill>False</x:AutoFill>
              <x:Row>\(anchor.row)</x:Row>
              <x:Column>\(anchor.column)</x:Column>
            </x:ClientData>
          </v:shape>
        """
    }

    /// Basic anchor: [startCol, dx1, startRow, dy1, endCol, dx2, endRow, dy2]
    private func vmlAnchor(column: Int, row: Int) -> String {
        // Roughly anchor to the cell with a small offset
        let startCol = column
        let startRow = row
        let endCol = column + 2
        let endRow = row + 3
        return "\(startCol), 15, \(startRow), 0, \(endCol), 15, \(endRow), 4"
    }
}
