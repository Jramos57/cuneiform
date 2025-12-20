import Foundation
import Testing
@testable import Cuneiform

@Suite("Formula Functions - Extended (Phase 5)")
struct FormulaFunctionsExtendedTests {
    
    // MARK: - Helper
    
    func makeTestEvaluator(cells: [String: CellValue]) -> FormulaEvaluator {
        FormulaEvaluator { ref in
            let key = ref.description
            return cells[key]
        }
    }
    
    // MARK: - Logical Functions
    
    @Test func evaluateAND_allTrue() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.boolean(true), .boolean(true), .boolean(true)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateAND_oneFalse() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.boolean(true), .boolean(false), .boolean(true)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateAND_withNumbers() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.number(1), .number(5), .number(-3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true)) // All non-zero = true
    }
    
    @Test func evaluateAND_withZero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.number(1), .number(0), .number(5)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false)) // Zero = false
    }
    
    @Test func evaluateOR_allFalse() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("OR", [.boolean(false), .boolean(false)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateOR_oneTrue() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("OR", [.boolean(false), .boolean(true), .boolean(false)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateNOT_true() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOT", [.boolean(true)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateNOT_false() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOT", [.boolean(false)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateNOT_number() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOT", [.number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true)) // 0 = false, NOT false = true
    }
    
    // MARK: - String Functions
    
    @Test func evaluateLEN_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEN", [.string("Hello")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateLEN_empty() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEN", [.string("")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    @Test func evaluateLEN_withSpaces() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEN", [.string("Hello World")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(11))
    }
    
    @Test func evaluateUPPER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("UPPER", [.string("hello world")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("HELLO WORLD"))
    }
    
    @Test func evaluateLOWER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LOWER", [.string("HELLO WORLD")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("hello world"))
    }
    
    @Test func evaluateCONCAT_strings() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCAT", [.string("Hello"), .string(" "), .string("World")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    @Test func evaluateCONCAT_mixed() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCAT", [.string("Value: "), .number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Value: 42.0"))
    }
    
    @Test func evaluateCONCAT_empty() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCAT", [])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string(""))
    }
    
    // MARK: - Math Functions
    
    @Test func evaluateROUND_noDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ROUND", [.number(3.7)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4))
    }
    
    @Test func evaluateROUND_twoDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ROUND", [.number(3.14159), .number(2)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3.14))
    }
    
    @Test func evaluateROUND_negativeDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ROUND", [.number(1234.5), .number(-2)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1200))
    }
    
    @Test func evaluateABS_positive() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ABS", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateABS_negative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ABS", [.number(-42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateABS_zero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ABS", [.number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    @Test func evaluateMEDIAN_oddCount() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(1), .number(3), .number(5), .number(7), .number(9)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateMEDIAN_evenCount() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(1), .number(2), .number(3), .number(4)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2.5))
    }
    
    @Test func evaluateMEDIAN_unsorted() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(5), .number(1), .number(9), .number(3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4)) // (3 + 5) / 2
    }
    
    @Test func evaluateMEDIAN_single() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    // MARK: - Integration Tests: Parser + New Functions
    
    @Test func parseAndEvaluateAND() throws {
        let cells: [String: CellValue] = [
            "A1": .boolean(true),
            "A2": .boolean(true),
            "A3": .boolean(false)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=AND(A1, A2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func parseAndEvaluateOR() throws {
        let cells: [String: CellValue] = [
            "A1": .boolean(false),
            "A2": .boolean(true),
            "A3": .boolean(false)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=OR(A1, A3, A2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func parseAndEvaluateUPPER() throws {
        let cells: [String: CellValue] = [
            "A1": .text("hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=UPPER(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("HELLO"))
    }
    
    @Test func parseAndEvaluateROUND() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3.14159)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=ROUND(A1, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3.14))
    }
    
    @Test func parseAndEvaluateMEDIAN() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MEDIAN(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(20))
    }
    
    // MARK: - Nested Function Tests
    
    @Test func nestedIF_AND() throws {
        let cells: [String: CellValue] = [
            "A1": .number(75),
            "B1": .number(80)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IF(AND(A1 >= 70, B1 >= 70), \"Pass\", \"Fail\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Pass"))
    }
    
    @Test func nestedROUND_SUM() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1.234),
            "A2": .number(2.567),
            "A3": .number(3.890)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=ROUND(SUM(A1:A3), 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(7.7))
    }
    
    @Test func complexNested() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "B1": .number(20),
            "C1": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IF(OR(A1 > 15, B1 > 15), ROUND(AVERAGE(A1:C1), 0), 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(20)) // B1 > 15, so ROUND(AVERAGE(10,20,30), 0) = 20
    }
}
