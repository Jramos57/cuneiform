import Foundation

/// Token types for formula parsing
public enum FormulaToken: Sendable, Equatable {
    case number(Double)
    case string(String)
    case cellRef(CellReference)
    case range(CellReference, CellReference)
    case function(String)
    case leftParen
    case rightParen
    case comma
    case plus
    case minus
    case multiply
    case divide
    case power
    case equals
    case notEquals
    case lessThan
    case greaterThan
    case lessThanOrEqual
    case greaterThanOrEqual
    case concat
}

/// Parsed formula expression tree
public indirect enum FormulaExpression: Sendable, Equatable {
    case number(Double)
    case string(String)
    case cellRef(CellReference)
    case range(CellReference, CellReference)
    case binaryOp(BinaryOperator, FormulaExpression, FormulaExpression)
    case functionCall(String, [FormulaExpression])
    case error(String)
    
    public enum BinaryOperator: String, Sendable, Equatable {
        case add = "+"
        case subtract = "-"
        case multiply = "*"
        case divide = "/"
        case power = "^"
        case equals = "="
        case notEquals = "<>"
        case lessThan = "<"
        case greaterThan = ">"
        case lessThanOrEqual = "<="
        case greaterThanOrEqual = ">="
        case concat = "&"
    }
}

/// Parses Excel formula strings into an expression tree
public struct FormulaParser {
    private let formula: String
    private var tokens: [FormulaToken] = []
    private var position = 0
    
    public init(_ formula: String) {
        self.formula = formula.trimmingCharacters(in: .whitespaces)
    }
    
    /// Parse the formula into an expression tree
    public func parse() throws -> FormulaExpression {
        var parser = self
        parser.tokenize()
        guard !parser.tokens.isEmpty else {
            throw FormulaError.emptyFormula
        }
        return try parser.parseExpression()
    }
    
    // MARK: - Tokenization
    
    private mutating func tokenize() {
        var chars = Array(formula)
        
        // Strip leading "=" if present
        if chars.first == "=" {
            chars.removeFirst()
        }
        
        var i = 0
        while i < chars.count {
            let ch = chars[i]
            
            // Skip whitespace
            if ch.isWhitespace {
                i += 1
                continue
            }
            
            // Numbers
            if ch.isNumber || (ch == "." && i + 1 < chars.count && chars[i + 1].isNumber) {
                let start = i
                while i < chars.count && (chars[i].isNumber || chars[i] == ".") {
                    i += 1
                }
                if let num = Double(String(chars[start..<i])) {
                    tokens.append(.number(num))
                }
                continue
            }
            
            // Strings (quoted)
            if ch == "\"" {
                i += 1
                let start = i
                while i < chars.count && chars[i] != "\"" {
                    i += 1
                }
                tokens.append(.string(String(chars[start..<i])))
                i += 1 // skip closing quote
                continue
            }
            
            // Cell references and function names
            if ch.isLetter || ch == "$" {
                let start = i
                
                // Check for range or cell reference
                var end = i
                while end < chars.count && (chars[end].isLetter || chars[end].isNumber || chars[end] == "$" || chars[end] == ".") {
                    end += 1
                }
                
                let token = String(chars[start..<end])
                
                // Check if followed by colon (range)
                if end < chars.count && chars[end] == ":" {
                    // Parse range like A1:B10
                    guard let startRef = CellReference(token) else {
                        i = end + 1
                        continue
                    }
                    
                    end += 1 // skip colon
                    let rangeEnd = end
                    while end < chars.count && (chars[end].isLetter || chars[end].isNumber || chars[end] == "$") {
                        end += 1
                    }
                    
                    if let endRef = CellReference(String(chars[rangeEnd..<end])) {
                        tokens.append(.range(startRef, endRef))
                        i = end
                        continue
                    }
                }
                
                // Check if followed by ( - it's a function
                if end < chars.count && chars[end] == "(" {
                    tokens.append(.function(token))
                    i = end
                    continue
                }
                
                // Try as cell reference
                if let cellRef = CellReference(token) {
                    tokens.append(.cellRef(cellRef))
                    i = end
                    continue
                }
                
                // Otherwise treat as function name that will be followed by (
                tokens.append(.function(token))
                i = end
                continue
            }
            
            // Operators and punctuation
            switch ch {
            case "(": tokens.append(.leftParen)
            case ")": tokens.append(.rightParen)
            case ",": tokens.append(.comma)
            case "+": tokens.append(.plus)
            case "-": tokens.append(.minus)
            case "*": tokens.append(.multiply)
            case "/": tokens.append(.divide)
            case "^": tokens.append(.power)
            case "&": tokens.append(.concat)
            case "=": tokens.append(.equals)
            case "<":
                if i + 1 < chars.count && chars[i + 1] == "=" {
                    tokens.append(.lessThanOrEqual)
                    i += 1
                } else if i + 1 < chars.count && chars[i + 1] == ">" {
                    tokens.append(.notEquals)
                    i += 1
                } else {
                    tokens.append(.lessThan)
                }
            case ">":
                if i + 1 < chars.count && chars[i + 1] == "=" {
                    tokens.append(.greaterThanOrEqual)
                    i += 1
                } else {
                    tokens.append(.greaterThan)
                }
            default:
                break
            }
            
            i += 1
        }
    }
    
    // MARK: - Expression Parsing
    
    private mutating func parseExpression() throws -> FormulaExpression {
        return try parseComparison()
    }
    
    private mutating func parseComparison() throws -> FormulaExpression {
        var left = try parseAdditive()
        
        while position < tokens.count {
            let token = tokens[position]
            let op: FormulaExpression.BinaryOperator
            
            switch token {
            case .equals: op = .equals
            case .notEquals: op = .notEquals
            case .lessThan: op = .lessThan
            case .greaterThan: op = .greaterThan
            case .lessThanOrEqual: op = .lessThanOrEqual
            case .greaterThanOrEqual: op = .greaterThanOrEqual
            default: return left
            }
            
            position += 1
            let right = try parseAdditive()
            left = .binaryOp(op, left, right)
        }
        
        return left
    }
    
    private mutating func parseAdditive() throws -> FormulaExpression {
        var left = try parseMultiplicative()
        
        while position < tokens.count {
            let token = tokens[position]
            let op: FormulaExpression.BinaryOperator
            
            switch token {
            case .plus: op = .add
            case .minus: op = .subtract
            case .concat: op = .concat
            default: return left
            }
            
            position += 1
            let right = try parseMultiplicative()
            left = .binaryOp(op, left, right)
        }
        
        return left
    }
    
    private mutating func parseMultiplicative() throws -> FormulaExpression {
        var left = try parsePower()
        
        while position < tokens.count {
            let token = tokens[position]
            let op: FormulaExpression.BinaryOperator
            
            switch token {
            case .multiply: op = .multiply
            case .divide: op = .divide
            default: return left
            }
            
            position += 1
            let right = try parsePower()
            left = .binaryOp(op, left, right)
        }
        
        return left
    }
    
    private mutating func parsePower() throws -> FormulaExpression {
        var left = try parsePrimary()
        
        while position < tokens.count {
            guard case .power = tokens[position] else { return left }
            position += 1
            let right = try parsePrimary()
            left = .binaryOp(.power, left, right)
        }
        
        return left
    }
    
    private mutating func parsePrimary() throws -> FormulaExpression {
        guard position < tokens.count else {
            throw FormulaError.unexpectedEndOfFormula
        }
        
        let token = tokens[position]
        position += 1
        
        switch token {
        case .number(let value):
            return .number(value)
            
        case .string(let value):
            return .string(value)
            
        case .cellRef(let ref):
            return .cellRef(ref)
            
        case .range(let start, let end):
            return .range(start, end)
            
        case .function(let name):
            guard position < tokens.count, case .leftParen = tokens[position] else {
                throw FormulaError.expectedLeftParen(function: name)
            }
            position += 1 // skip (
            
            var args: [FormulaExpression] = []
            
            // Empty argument list
            if position < tokens.count, case .rightParen = tokens[position] {
                position += 1
                return .functionCall(name, args)
            }
            
            // Parse arguments
            while true {
                let arg = try parseExpression()
                args.append(arg)
                
                guard position < tokens.count else {
                    throw FormulaError.unexpectedEndOfFormula
                }
                
                let next = tokens[position]
                if case .rightParen = next {
                    position += 1
                    break
                } else if case .comma = next {
                    position += 1
                    continue
                } else {
                    throw FormulaError.expectedCommaOrRightParen
                }
            }
            
            return .functionCall(name, args)
            
        case .leftParen:
            let expr = try parseExpression()
            guard position < tokens.count, case .rightParen = tokens[position] else {
                throw FormulaError.expectedRightParen
            }
            position += 1
            return expr
            
        case .minus:
            // Unary minus
            let expr = try parsePrimary()
            return .binaryOp(.multiply, .number(-1), expr)
            
        default:
            throw FormulaError.unexpectedToken(String(describing: token))
        }
    }
}

// MARK: - Formula Errors

public enum FormulaError: Error, Sendable, Equatable {
    case emptyFormula
    case unexpectedEndOfFormula
    case unexpectedToken(String)
    case expectedLeftParen(function: String)
    case expectedRightParen
    case expectedCommaOrRightParen
    case invalidCellReference(String)
    case invalidRange(String)
    case divisionByZero
    case invalidArgumentCount(function: String, expected: Int, got: Int)
    case invalidArgumentType(function: String, argument: Int)
    case circularReference
    case nameNotFound(String)
    case evaluationError(String)
}
