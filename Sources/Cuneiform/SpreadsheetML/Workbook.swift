import Foundation

/// A high-level view of an .xlsx workbook.
public struct Workbook: Sendable {
    private let package: OPCPackage
    private let workbookInfo: WorkbookInfo
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    private let definedNames: [DefinedName]

    /// All sheets in the workbook
    public var sheets: [SheetInfo] { workbookInfo.sheets }

    /// Open an .xlsx file from a URL
    public static func open(url: URL) throws(CuneiformError) -> Workbook {
        let pkg: OPCPackage
        do {
            pkg = try OPCPackage.open(url: url)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.invalidPackageStructure(reason: error.localizedDescription)
        }

        // Parse workbook
        let wbData: Data
        do {
            wbData = try pkg.readPart(.workbook)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: PartPath.workbook.value, detail: error.localizedDescription)
        }
        let wb = try WorkbookParser.parse(data: wbData)

        // Parse shared strings
        let ss: SharedStrings
        if pkg.partExists(.sharedStrings) {
            let ssData: Data
            do {
                ssData = try pkg.readPart(.sharedStrings)
            } catch let err as CuneiformError {
                throw err
            } catch {
                throw CuneiformError.malformedXML(part: PartPath.sharedStrings.value, detail: error.localizedDescription)
            }
            ss = try SharedStringsParser.parse(data: ssData)
        } else {
            ss = .empty
        }

        // Parse styles
        let st: StylesInfo
        if pkg.partExists(.styles) {
            let stData: Data
            do {
                stData = try pkg.readPart(.styles)
            } catch let err as CuneiformError {
                throw err
            } catch {
                throw CuneiformError.malformedXML(part: PartPath.styles.value, detail: error.localizedDescription)
            }
            st = try StylesParser.parse(data: stData)
        } else {
            st = .empty
        }

        return Workbook(package: pkg, workbookInfo: wb, sharedStrings: ss, styles: st)
    }

    /// Get a sheet by name
    public func sheet(named: String) throws(CuneiformError) -> Sheet? {
        guard let sheetInfo = workbookInfo.sheet(named: named) else { return nil }
        return try loadSheet(sheetInfo)
    }

    /// Get a sheet by index (0-based)
    public func sheet(at index: Int) throws(CuneiformError) -> Sheet? {
        guard index >= 0, index < sheets.count else { return nil }
        return try loadSheet(sheets[index])
    }

    /// Load a sheet by SheetInfo
    private func loadSheet(_ sheetInfo: SheetInfo) throws(CuneiformError) -> Sheet {
        let wbRels: Relationships
        do {
            var mutablePkg = package
            wbRels = try mutablePkg.relationships(for: .workbook)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: PartPath.workbookRelationships.value, detail: error.localizedDescription)
        }
        
        guard let rel = wbRels[sheetInfo.relationshipId] else {
            throw CuneiformError.missingRequiredElement(
                element: "Relationship",
                inPart: "Sheet '\(sheetInfo.name)' with ID '\(sheetInfo.relationshipId)'"
            )
        }

        let sheetPath = rel.resolveTarget(relativeTo: .workbook)
        let sheetData: Data
        do {
            var mutablePkg = package
            sheetData = try mutablePkg.readPart(sheetPath)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: sheetPath.value, detail: error.localizedDescription)
        }
        let worksheet = try WorksheetParser.parse(data: sheetData)

        // Parse comments via worksheet relationships (if present)
        let comments: [Comment]
        do {
            var mutablePkg = package
            let wsRels = try mutablePkg.relationships(for: sheetPath)
            let commentRels = wsRels[.comments]
            if commentRels.isEmpty {
                comments = []
            } else {
                var acc: [Comment] = []
                for cRel in commentRels {
                    let cPath = cRel.resolveTarget(relativeTo: sheetPath)
                    let cData = try mutablePkg.readPart(cPath)
                    let parsed = try CommentsParser.parse(data: cData)
                    acc.append(contentsOf: parsed.comments)
                }
                comments = acc
            }
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: sheetPath.relationshipsPath.value, detail: error.localizedDescription)
        }

        return Sheet(data: worksheet, sharedStrings: sharedStrings, styles: styles, comments: comments)
    }

    /// Workbook-level defined names
    public var definedNamesList: [DefinedName] { definedNames }

    /// Get a workbook-level defined name by name (case-sensitive)
    public func definedName(_ name: String) -> DefinedName? {
        definedNames.first { $0.name == name }
    }

    /// Parse a defined name's reference into (sheet, range) components if it is of the form "Sheet!$A$1:$B$10"
    public func definedNameRange(_ name: String) -> (sheet: String, range: String)? {
        guard let dn = definedName(name) else { return nil }
        // Expected patterns like Sheet1!$A$1:$A$10 or 'My Sheet'!$A$1:$B$2
        let refers = dn.refersTo.trimmingCharacters(in: .whitespacesAndNewlines)
        if refers.isEmpty { return nil }
        // If sheet has quotes, split on last !
        if let bangIndex = refers.lastIndex(of: "!") {
            let sheetName = String(refers[..<bangIndex]).trimmingCharacters(in: CharacterSet(charactersIn: "'"))
            let rangePart = String(refers[refers.index(after: bangIndex)...])
            return (sheet: sheetName, range: rangePart)
        }
        return nil
    }

    init(package: OPCPackage, workbookInfo: WorkbookInfo, sharedStrings: SharedStrings, styles: StylesInfo) {
        self.package = package
        self.workbookInfo = workbookInfo
        self.sharedStrings = sharedStrings
        self.styles = styles
        self.definedNames = workbookInfo.definedNames
    }
}
