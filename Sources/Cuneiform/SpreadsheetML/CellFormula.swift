/// A cell formula as extracted from SpreadsheetML.
public struct CellFormula: Sendable, Hashable {
    /// The formula string (e.g., "=SUM(A1:A10)")
    public let formula: String

    /// Whether the formula is an array formula
    public let isArrayFormula: Bool

    /// Create a cell formula
    public init(_ formula: String, isArrayFormula: Bool = false) {
        self.formula = formula
        self.isArrayFormula = isArrayFormula
    }

    /// Human-readable description
    public var description: String {
        isArrayFormula ? "{=\(formula)}" : formula
    }
}

extension CellFormula: CustomStringConvertible {
    public var debugDescription: String { description }
}
