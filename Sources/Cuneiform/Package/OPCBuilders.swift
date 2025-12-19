import Foundation

/// Builder for [Content_Types].xml
public struct ContentTypesBuilder {
    private var defaults: [String: String] = [:]  // extension -> contentType
    private var overrides: [String: String] = [:] // partName -> contentType
    
    public init() {
        // Common defaults
        defaults["rels"] = "application/vnd.openxmlformats-package.relationships+xml"
        defaults["xml"] = "application/xml"
    }
    
    /// Add a default content type for an extension
    public mutating func addDefault(extension ext: String, contentType: String) {
        defaults[ext] = contentType
    }
    
    /// Add an override content type for a specific part
    public mutating func addOverride(partName: String, contentType: String) {
        overrides[partName] = contentType
    }
    
    /// Build the XML data
    public func build() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        """
        
        // Add defaults
        for (ext, ct) in defaults.sorted(by: { $0.key < $1.key }) {
            xml += """
            
            <Default Extension="\(ext)" ContentType="\(ct)"/>
            """
        }
        
        // Add overrides
        for (part, ct) in overrides.sorted(by: { $0.key < $1.key }) {
            xml += """
            
            <Override PartName="\(part)" ContentType="\(ct)"/>
            """
        }
        
        xml += """
        
        </Types>
        """
        
        return xml.data(using: .utf8)!
    }
}

/// Builder for .rels relationship files
public struct RelationshipsBuilder {
    private var relationships: [(id: String, type: String, target: String, targetMode: String?)] = []
    private var nextId = 1
    
    public init() {}
    
    /// Add a relationship
    @discardableResult
    public mutating func addRelationship(type: String, target: String, id: String? = nil, targetMode: String? = nil) -> String {
        let relId = id ?? "rId\(nextId)"
        nextId += 1
        relationships.append((relId, type, target, targetMode))
        return relId
    }
    
    /// Build the XML data
    public func build() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        
        for (id, type, target, targetMode) in relationships {
            xml += """
            
            <Relationship Id="\(id)" Type="\(type)" Target="\(target)"\(targetMode != nil ? " TargetMode=\"\(targetMode!)\"" : "")/>
            """
        }
        
        xml += """
        
        </Relationships>
        """
        
        return xml.data(using: .utf8)!
    }
}
