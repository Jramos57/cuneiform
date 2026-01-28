import Foundation
import Testing
@testable import Cuneiform

@Suite("Dependency Graph & Recalculation Tests (Phase 5)")
struct DependencyGraphTests {
    
    // MARK: - Basic Dependency Tracking
    
    @Test func addSingleDependency() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        graph.addFormula(at: a1, dependsOn: [b1])
        
        let deps = graph.directDependencies(of: a1)
        #expect(deps.contains(b1))
        #expect(deps.count == 1)
    }
    
    @Test func addMultipleDependencies() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        let d1 = CellReference("D1")
        
        graph.addFormula(at: a1, dependsOn: [b1, c1, d1])
        
        let deps = graph.directDependencies(of: a1)
        #expect(deps.count == 3)
        #expect(deps.contains(b1))
        #expect(deps.contains(c1))
        #expect(deps.contains(d1))
    }
    
    @Test func trackDependents() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        
        // A1 and C1 both depend on B1
        graph.addFormula(at: a1, dependsOn: [b1])
        graph.addFormula(at: c1, dependsOn: [b1])
        
        let dependents = graph.directDependents(of: b1)
        #expect(dependents.count == 2)
        #expect(dependents.contains(a1))
        #expect(dependents.contains(c1))
    }
    
    @Test func removeFormula() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        graph.addFormula(at: a1, dependsOn: [b1])
        #expect(graph.directDependencies(of: a1).count == 1)
        
        graph.removeFormula(at: a1)
        #expect(graph.directDependencies(of: a1).isEmpty)
        #expect(graph.directDependents(of: b1).isEmpty)
    }
    
    // MARK: - Recalculation Order
    
    @Test func simpleRecalculationOrder() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        // A1 = B1 + 1
        graph.addFormula(at: a1, dependsOn: [b1])
        
        let order = try graph.recalculationOrder(for: [b1])
        #expect(order == [a1])
    }
    
    @Test func chainedRecalculationOrder() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        
        // A1 = B1 + 1
        // B1 = C1 + 1
        graph.addFormula(at: a1, dependsOn: [b1])
        graph.addFormula(at: b1, dependsOn: [c1])
        
        let order = try graph.recalculationOrder(for: [c1])
        // Should calculate B1 before A1
        #expect(order.count == 2)
        if let b1Index = order.firstIndex(of: b1),
           let a1Index = order.firstIndex(of: a1) {
            #expect(b1Index < a1Index)
        }
    }
    
    @Test func diamondRecalculationOrder() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        let d1 = CellReference("D1")
        
        // Diamond pattern:
        //     D1
        //    /  \
        //   B1  C1
        //    \  /
        //     A1
        graph.addFormula(at: d1, dependsOn: [b1, c1])
        graph.addFormula(at: b1, dependsOn: [a1])
        graph.addFormula(at: c1, dependsOn: [a1])
        
        let order = try graph.recalculationOrder(for: [a1])
        #expect(order.count == 3)
        
        // B1 and C1 must come before D1
        if let b1Index = order.firstIndex(of: b1),
           let c1Index = order.firstIndex(of: c1),
           let d1Index = order.firstIndex(of: d1) {
            #expect(b1Index < d1Index)
            #expect(c1Index < d1Index)
        }
    }
    
    // MARK: - Circular Reference Detection
    
    @Test func detectSimpleCircularReference() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        // A1 = B1 + 1
        // B1 = A1 + 1  (circular!)
        graph.addFormula(at: a1, dependsOn: [b1])
        graph.addFormula(at: b1, dependsOn: [a1])
        
        #expect(graph.hasCircularReference(involving: [a1, b1]))
    }
    
    @Test func detectIndirectCircularReference() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        
        // A1 = B1 + 1
        // B1 = C1 + 1
        // C1 = A1 + 1  (indirect circular!)
        graph.addFormula(at: a1, dependsOn: [b1])
        graph.addFormula(at: b1, dependsOn: [c1])
        graph.addFormula(at: c1, dependsOn: [a1])
        
        #expect(graph.hasCircularReference(involving: [a1, b1, c1]))
    }
    
    @Test func noCircularReferenceInAcyclicGraph() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        
        // A1 = B1 + C1 (no cycle)
        graph.addFormula(at: a1, dependsOn: [b1, c1])
        
        #expect(!graph.hasCircularReference(involving: [a1, b1, c1]))
    }
    
    // MARK: - WorksheetCalculator
    
    @Test func calculatorBasicSetup() throws {
        let cells: [String: CellValue] = [
            "B1": .number(10)
        ]
        
        var calculator = WorksheetCalculator(cellResolver: { ref in
            cells[ref.description]
        })
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        calculator.registerFormula(at: a1, references: [b1])
        
        let order = try calculator.recalculationOrder(for: [b1])
        #expect(order == [a1])
    }
    
    @Test func calculatorEvaluateFormula() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "B1": .number(20)
        ]
        
        let calculator = WorksheetCalculator(cellResolver: { ref in
            cells[ref.description]
        })
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        let expr = FormulaExpression.binaryOp(.add, .cellRef(a1), .cellRef(b1))
        let result = try calculator.evaluate(expr)
        #expect(result == .number(30))
    }
    
    @Test func extractReferencesFromExpression() throws {
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        let c1 = CellReference("C1")
        
        // SUM(A1:B1) + C1
        let sumExpr = FormulaExpression.functionCall("SUM", [.range(a1, b1)])
        let expr = FormulaExpression.binaryOp(.add, sumExpr, .cellRef(c1))
        
        let refs = WorksheetCalculator.extractReferences(from: expr)
        #expect(refs.count >= 3)
        #expect(refs.contains(a1))
        #expect(refs.contains(b1))
        #expect(refs.contains(c1))
    }
    
    // MARK: - Integration: Complex Scenarios
    
    @Test func complexDependencyChain() throws {
        var graph = DependencyGraph()
        
        let a1 = CellReference("A1")
        let a2 = CellReference("A2")
        let a3 = CellReference("A3")
        let a4 = CellReference("A4")
        let a5 = CellReference("A5")
        
        // A5 = SUM(A1:A4)
        // A4 = A3 * 2
        // A3 = A2 + 10
        // A2 = A1 + 5
        // A1 = 10 (constant)
        
        graph.addFormula(at: a5, dependsOn: [a1, a2, a3, a4])
        graph.addFormula(at: a4, dependsOn: [a3])
        graph.addFormula(at: a3, dependsOn: [a2])
        graph.addFormula(at: a2, dependsOn: [a1])
        
        let order = try graph.recalculationOrder(for: [a1])
        
        // Should calculate in order: A2, A3, A4, A5
        #expect(order.count == 4)
        if let a2Index = order.firstIndex(of: a2),
           let a3Index = order.firstIndex(of: a3),
           let a4Index = order.firstIndex(of: a4),
           let a5Index = order.firstIndex(of: a5) {
            #expect(a2Index < a3Index)
            #expect(a3Index < a4Index)
            #expect(a4Index < a5Index)
        }
    }
}
