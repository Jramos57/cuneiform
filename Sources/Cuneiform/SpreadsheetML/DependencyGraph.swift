import Foundation

/// Dependency graph for formula recalculation
public struct DependencyGraph: Sendable {
    /// Node representing a cell with formulas
    private struct Node: Hashable {
        let ref: CellReference
    }
    
    /// Adjacency list: cell -> cells it depends on
    private var dependencies: [CellReference: Set<CellReference>]
    
    /// Reverse adjacency: cell -> cells that depend on it
    private var dependents: [CellReference: Set<CellReference>]
    
    public init() {
        self.dependencies = [:]
        self.dependents = [:]
    }
    
    /// Add a formula dependency: formula in `cell` depends on `references`
    public mutating func addFormula(at cell: CellReference, dependsOn references: [CellReference]) {
        dependencies[cell] = Set(references)
        
        for ref in references {
            dependents[ref, default: []].insert(cell)
        }
    }
    
    /// Remove all dependencies for a cell (e.g., when formula is deleted)
    public mutating func removeFormula(at cell: CellReference) {
        guard let deps = dependencies[cell] else { return }
        
        for ref in deps {
            dependents[ref]?.remove(cell)
        }
        
        dependencies.removeValue(forKey: cell)
    }
    
    /// Get all cells that directly depend on the given cell
    public func directDependents(of cell: CellReference) -> Set<CellReference> {
        return dependents[cell] ?? []
    }
    
    /// Get all cells that the given cell directly depends on
    public func directDependencies(of cell: CellReference) -> Set<CellReference> {
        return dependencies[cell] ?? []
    }
    
    /// Calculate recalculation order for cells affected by changes to `changedCells`
    /// Returns cells in topological order (dependencies before dependents)
    /// Throws if circular reference detected
    public func recalculationOrder(for changedCells: Set<CellReference>) throws -> [CellReference] {
        // Collect all affected cells (transitive closure of dependents)
        var affected = Set<CellReference>()
        var queue = Array(changedCells)
        
        while !queue.isEmpty {
            let cell = queue.removeFirst()
            for dependent in directDependents(of: cell) {
                if !affected.contains(dependent) {
                    affected.insert(dependent)
                    queue.append(dependent)
                }
            }
        }
        
        // Topological sort of affected cells
        return try topologicalSort(cells: affected)
    }
    
    /// Detect if there's a circular reference involving any of the given cells
    public func hasCircularReference(involving cells: Set<CellReference>) -> Bool {
        do {
            _ = try topologicalSort(cells: cells)
            return false
        } catch {
            return true
        }
    }
    
    // MARK: - Private Helpers
    
    private func topologicalSort(cells: Set<CellReference>) throws -> [CellReference] {
        var result: [CellReference] = []
        var visited = Set<CellReference>()
        var visiting = Set<CellReference>()
        
        func visit(_ cell: CellReference) throws {
            if visited.contains(cell) {
                return
            }
            
            if visiting.contains(cell) {
                throw DependencyError.circularReference(cell: cell)
            }
            
            visiting.insert(cell)
            
            // Visit dependencies first
            for dep in directDependencies(of: cell) {
                if cells.contains(dep) {
                    try visit(dep)
                }
            }
            
            visiting.remove(cell)
            visited.insert(cell)
            result.append(cell)
        }
        
        for cell in cells {
            try visit(cell)
        }
        
        return result
    }
}

/// Errors related to dependency graph operations
public enum DependencyError: Error, Sendable {
    case circularReference(cell: CellReference)
}

/// Calculator for worksheet formulas with dependency tracking
public struct WorksheetCalculator: Sendable {
    private var graph: DependencyGraph
    private let cellResolver: @Sendable (CellReference) -> CellValue?
    private let evaluator: FormulaEvaluator
    
    /// Create calculator with cell resolver
    public init(cellResolver: @escaping @Sendable (CellReference) -> CellValue?) {
        self.graph = DependencyGraph()
        self.cellResolver = cellResolver
        self.evaluator = FormulaEvaluator(cellResolver: cellResolver)
    }
    
    /// Register a formula cell
    public mutating func registerFormula(at cell: CellReference, references: [CellReference]) {
        graph.addFormula(at: cell, dependsOn: references)
    }
    
    /// Remove a formula cell
    public mutating func removeFormula(at cell: CellReference) {
        graph.removeFormula(at: cell)
    }
    
    /// Calculate cells that need recalculation when given cells change
    /// Returns cells in correct evaluation order
    public func recalculationOrder(for changedCells: Set<CellReference>) throws -> [CellReference] {
        return try graph.recalculationOrder(for: changedCells)
    }
    
    /// Evaluate a parsed formula expression
    public func evaluate(_ expression: FormulaExpression) throws -> FormulaValue {
        return try evaluator.evaluate(expression)
    }
    
    /// Check for circular references
    public func hasCircularReference(involving cells: Set<CellReference>) -> Bool {
        return graph.hasCircularReference(involving: cells)
    }
    
    /// Extract cell references from a formula expression
    public static func extractReferences(from expression: FormulaExpression) -> [CellReference] {
        var refs: [CellReference] = []
        
        func extract(_ expr: FormulaExpression) {
            switch expr {
            case .cellRef(let ref):
                refs.append(ref)
            case .range(let start, let end):
                // Add both corners of range
                refs.append(start)
                refs.append(end)
            case .binaryOp(_, let left, let right):
                extract(left)
                extract(right)
            case .functionCall(_, let args):
                for arg in args {
                    extract(arg)
                }
            default:
                break
            }
        }
        
        extract(expression)
        return refs
    }
}
