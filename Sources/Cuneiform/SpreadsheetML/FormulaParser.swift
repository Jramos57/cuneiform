import Foundation

/// Lexical tokens produced during formula tokenization.
///
/// The formula parser tokenizes Excel formula strings into a sequence of tokens before
/// constructing the abstract syntax tree (AST). Each token represents a discrete element
/// of the formula such as numbers, cell references, operators, or functions.
///
/// ## Topics
///
/// ### Literals
/// - ``number(_:)``
/// - ``string(_:)``
///
/// ### References
/// - ``cellRef(_:)``
/// - ``range(_:_:)``
///
/// ### Functions and Delimiters
/// - ``function(_:)``
/// - ``leftParen``
/// - ``rightParen``
/// - ``comma``
///
/// ### Arithmetic Operators
/// - ``plus``
/// - ``minus``
/// - ``multiply``
/// - ``divide``
/// - ``power``
///
/// ### Comparison Operators
/// - ``equals``
/// - ``notEquals``
/// - ``lessThan``
/// - ``greaterThan``
/// - ``lessThanOrEqual``
/// - ``greaterThanOrEqual``
///
/// ### Text Operators
/// - ``concat``
public enum FormulaToken: Sendable, Equatable {
    /// A numeric literal token.
    ///
    /// Examples: `42`, `3.14`, `.5`
    case number(Double)
    
    /// A text string literal token (without quotes).
    ///
    /// Example: For the formula `"Hello"`, this contains `Hello`
    case string(String)
    
    /// A cell reference token.
    ///
    /// Examples: `A1`, `$B$2`, `C$3`
    case cellRef(CellReference)
    
    /// A cell range token.
    ///
    /// Examples: `A1:B10`, `$A$1:$C$5`
    case range(CellReference, CellReference)
    
    /// A function name token.
    ///
    /// Examples: `SUM`, `AVERAGE`, `IF`
    case function(String)
    
    /// Left parenthesis token `(`.
    case leftParen
    
    /// Right parenthesis token `)`.
    case rightParen
    
    /// Comma token `,` (argument separator).
    case comma
    
    /// Addition operator token `+`.
    case plus
    
    /// Subtraction operator token `-`.
    case minus
    
    /// Multiplication operator token `*`.
    case multiply
    
    /// Division operator token `/`.
    case divide
    
    /// Exponentiation operator token `^`.
    case power
    
    /// Equality operator token `=`.
    case equals
    
    /// Inequality operator token `<>`.
    case notEquals
    
    /// Less than operator token `<`.
    case lessThan
    
    /// Greater than operator token `>`.
    case greaterThan
    
    /// Less than or equal operator token `<=`.
    case lessThanOrEqual
    
    /// Greater than or equal operator token `>=`.
    case greaterThanOrEqual
    
    /// String concatenation operator token `&`.
    case concat
}

/// Abstract syntax tree (AST) representation of a parsed Excel formula.
///
/// After tokenization, the formula parser constructs an abstract syntax tree that represents
/// the logical structure of the formula. This tree is then evaluated by ``FormulaEvaluator``
/// to produce the final result.
///
/// The AST uses an indirect enum to represent nested expressions, supporting arbitrary
/// depth for complex formulas. Each node in the tree represents either a terminal value
/// (number, string, cell reference) or an operation (binary operator, function call).
///
/// ## Parsing Example
///
/// The formula `=SUM(A1:A3) + B1 * 2` produces this AST structure:
///
/// ```
/// .binaryOp(.add,
///   .functionCall("SUM", [.range(A1, A3)]),
///   .binaryOp(.multiply,
///     .cellRef(B1),
///     .number(2.0)
///   )
/// )
/// ```
///
/// ## Usage
///
/// Parse a formula string into an AST:
///
/// ```swift
/// let parser = FormulaParser("=A1 + B2")
/// let ast = try parser.parse()
///
/// // AST structure:
/// // .binaryOp(.add,
/// //   .cellRef(CellReference("A1")),
/// //   .cellRef(CellReference("B2"))
/// // )
/// ```
///
/// ## Topics
///
/// ### Terminal Values
/// - ``number(_:)``
/// - ``string(_:)``
/// - ``cellRef(_:)``
/// - ``range(_:_:)``
///
/// ### Operations
/// - ``binaryOp(_:_:_:)``
/// - ``functionCall(_:_:)``
/// - ``error(_:)``
///
/// ### Operators
/// - ``BinaryOperator``
///
/// ## See Also
/// - ``FormulaParser``
/// - ``FormulaEvaluator``
public indirect enum FormulaExpression: Sendable, Equatable {
    /// A numeric literal value.
    case number(Double)
    
    /// A text string literal.
    case string(String)
    
    /// A reference to a single cell (e.g., `A1`, `$B$2`).
    case cellRef(CellReference)
    
    /// A range of cells (e.g., `A1:B10`).
    case range(CellReference, CellReference)
    
    /// A binary operation applied to two expressions.
    ///
    /// - Parameters:
    ///   - operator: The binary operator to apply
    ///   - left: The left-hand expression
    ///   - right: The right-hand expression
    case binaryOp(BinaryOperator, FormulaExpression, FormulaExpression)
    
    /// A function call with arguments.
    ///
    /// - Parameters:
    ///   - name: The function name (case-insensitive, e.g., "SUM", "AVERAGE")
    ///   - arguments: Array of argument expressions
    case functionCall(String, [FormulaExpression])
    
    /// A parse error with a descriptive message.
    case error(String)
    
    /// Binary operators supported in Excel formulas.
    ///
    /// These operators follow Excel's precedence rules:
    /// 1. Power (`^`)
    /// 2. Multiplication and Division (`*`, `/`)
    /// 3. Addition and Subtraction (`+`, `-`)
    /// 4. String Concatenation (`&`)
    /// 5. Comparison (`=`, `<>`, `<`, `>`, `<=`, `>=`)
    ///
    /// ## Topics
    ///
    /// ### Arithmetic Operators
    /// - ``add``
    /// - ``subtract``
    /// - ``multiply``
    /// - ``divide``
    /// - ``power``
    ///
    /// ### Comparison Operators
    /// - ``equals``
    /// - ``notEquals``
    /// - ``lessThan``
    /// - ``greaterThan``
    /// - ``lessThanOrEqual``
    /// - ``greaterThanOrEqual``
    ///
    /// ### Text Operators
    /// - ``concat``
    public enum BinaryOperator: String, Sendable, Equatable {
        /// Addition operator `+`.
        case add = "+"
        
        /// Subtraction operator `-`.
        case subtract = "-"
        
        /// Multiplication operator `*`.
        case multiply = "*"
        
        /// Division operator `/`.
        case divide = "/"
        
        /// Exponentiation operator `^`.
        case power = "^"
        
        /// Equality comparison operator `=`.
        case equals = "="
        
        /// Inequality comparison operator `<>`.
        case notEquals = "<>"
        
        /// Less than comparison operator `<`.
        case lessThan = "<"
        
        /// Greater than comparison operator `>`.
        case greaterThan = ">"
        
        /// Less than or equal comparison operator `<=`.
        case lessThanOrEqual = "<="
        
        /// Greater than or equal comparison operator `>=`.
        case greaterThanOrEqual = ">="
        
        /// String concatenation operator `&`.
        case concat = "&"
    }
}

/// Parses Excel formula strings into an abstract syntax tree (AST).
///
/// `FormulaParser` implements a recursive descent parser for Excel formulas, converting
/// textual formula strings into structured ``FormulaExpression`` trees. The parser handles
/// the complete Excel formula syntax including operators, functions, cell references, ranges,
/// and literals.
///
/// ## Parsing Process
///
/// The parser operates in two phases:
///
/// 1. **Tokenization**: The formula string is scanned and broken into tokens (numbers,
///    operators, cell references, etc.)
/// 2. **Parsing**: Tokens are assembled into an AST using recursive descent with
///    operator precedence
///
/// ## Operator Precedence
///
/// The parser implements Excel's operator precedence (highest to lowest):
///
/// 1. Power (`^`)
/// 2. Multiplication and Division (`*`, `/`)
/// 3. Addition and Subtraction (`+`, `-`)
/// 4. String Concatenation (`&`)
/// 5. Comparison operators (`=`, `<>`, `<`, `>`, `<=`, `>=`)
///
/// ## Supported Syntax
///
/// ### Literals
/// - **Numbers**: `42`, `3.14`, `.5`
/// - **Strings**: `"Hello"`, `"Revenue"`
///
/// ### Cell References
/// - **Relative**: `A1`, `B2`
/// - **Absolute**: `$A$1`, `$B1`, `A$1`
/// - **Ranges**: `A1:B10`, `$A$1:$C$5`
///
/// ### Operators
/// - **Arithmetic**: `+`, `-`, `*`, `/`, `^`
/// - **Comparison**: `=`, `<>`, `<`, `>`, `<=`, `>=`
/// - **Text**: `&` (concatenation)
///
/// ### Functions
/// - **Any Excel function**: `SUM()`, `AVERAGE()`, `IF()`, etc.
/// - **Nested functions**: `SUM(IF(A1:A10>0, A1:A10, 0))`
///
/// ## Usage Examples
///
/// ### Simple Arithmetic
///
/// ```swift
/// let parser = FormulaParser("=2 + 3 * 4")
/// let ast = try parser.parse()
/// // Result: .binaryOp(.add, .number(2.0),
/// //           .binaryOp(.multiply, .number(3.0), .number(4.0)))
/// ```
///
/// ### Cell References
///
/// ```swift
/// let parser = FormulaParser("=A1 + B2")
/// let ast = try parser.parse()
/// // Result: .binaryOp(.add, .cellRef(A1), .cellRef(B2))
/// ```
///
/// ### Function Calls
///
/// ```swift
/// let parser = FormulaParser("=SUM(A1:A10)")
/// let ast = try parser.parse()
/// // Result: .functionCall("SUM", [.range(A1, A10)])
/// ```
///
/// ### Complex Nested Formula
///
/// ```swift
/// let parser = FormulaParser("=IF(SUM(A1:A5) > 100, \"High\", \"Low\")")
/// let ast = try parser.parse()
/// // Result: .functionCall("IF", [
/// //   .binaryOp(.greaterThan,
/// //     .functionCall("SUM", [.range(A1, A5)]),
/// //     .number(100.0)
/// //   ),
/// //   .string("High"),
/// //   .string("Low")
/// // ])
/// ```
///
/// ## AST Structure
///
/// The parsed AST uses an indirect enum to support arbitrary nesting depth. Each node
/// represents either:
///
/// - **Terminal**: A leaf value (number, string, cell reference, range)
/// - **Operation**: A binary operator or function call with child expressions
///
/// ### Understanding the AST
///
/// For the formula `=A1 * 2 + B1`:
///
/// ```
/// .binaryOp(.add,                    // Top-level addition
///   .binaryOp(.multiply,              // Left side: A1 * 2
///     .cellRef(A1),
///     .number(2.0)
///   ),
///   .cellRef(B1)                      // Right side: B1
/// )
/// ```
///
/// ## Error Handling
///
/// The parser throws ``FormulaError`` for invalid syntax:
///
/// ```swift
/// let parser = FormulaParser("=SUM(A1:A10")  // Missing closing paren
/// do {
///     let ast = try parser.parse()
/// } catch {
///     // Throws: FormulaError.expectedRightParen
/// }
/// ```
///
/// Common errors include:
/// - ``FormulaError/emptyFormula``: Empty or whitespace-only formula
/// - ``FormulaError/unexpectedToken(_:)``: Invalid token in expression
/// - ``FormulaError/expectedRightParen``: Mismatched parentheses
/// - ``FormulaError/expectedCommaOrRightParen``: Invalid function argument separator
///
/// ## Thread Safety
///
/// `FormulaParser` is a value type and is thread-safe. Each parse operation creates
/// an independent copy of the parser state.
///
/// ## Topics
///
/// ### Creating a Parser
/// - ``init(_:)``
///
/// ### Parsing Formulas
/// - ``parse()``
///
/// ## See Also
/// - ``FormulaExpression``: The AST produced by parsing
/// - ``FormulaEvaluator``: Evaluates parsed expressions
/// - ``FormulaError``: Errors that can occur during parsing
/// - ``FormulaToken``: Tokens produced during tokenization
public struct FormulaParser {
    private let formula: String
    private var tokens: [FormulaToken] = []
    private var position = 0
    
    /// Creates a formula parser for the given formula string.
    ///
    /// The leading equals sign (`=`) is optional and will be stripped automatically
    /// during parsing. Whitespace at the beginning and end of the formula is also
    /// trimmed.
    ///
    /// - Parameter formula: The Excel formula string to parse
    ///
    /// ## Examples
    ///
    /// All of these are equivalent:
    ///
    /// ```swift
    /// let p1 = FormulaParser("=A1 + B1")
    /// let p2 = FormulaParser("A1 + B1")
    /// let p3 = FormulaParser("  =A1 + B1  ")
    /// ```
    public init(_ formula: String) {
        self.formula = formula.trimmingCharacters(in: .whitespaces)
    }
    
    /// Parses the formula string into an abstract syntax tree (AST).
    ///
    /// This method performs two-phase parsing:
    /// 1. **Tokenization**: Scans the formula string into tokens
    /// 2. **Expression parsing**: Builds an AST from the token stream
    ///
    /// The parser implements recursive descent with operator precedence to correctly
    /// handle complex expressions. The resulting AST can be evaluated using
    /// ``FormulaEvaluator/evaluate(_:)``.
    ///
    /// - Returns: A ``FormulaExpression`` representing the parsed formula
    /// - Throws: ``FormulaError`` if the formula contains syntax errors
    ///
    /// ## Examples
    ///
    /// ### Basic Arithmetic
    ///
    /// ```swift
    /// let parser = FormulaParser("=10 + 20 * 3")
    /// let ast = try parser.parse()
    /// // Result follows order of operations: 10 + (20 * 3)
    /// // AST: .binaryOp(.add,
    /// //        .number(10.0),
    /// //        .binaryOp(.multiply, .number(20.0), .number(3.0))
    /// //      )
    /// ```
    ///
    /// ### Cell References and Ranges
    ///
    /// ```swift
    /// let parser = FormulaParser("=SUM(A1:A10) + B5")
    /// let ast = try parser.parse()
    /// // AST: .binaryOp(.add,
    /// //        .functionCall("SUM", [.range(A1, A10)]),
    /// //        .cellRef(B5)
    /// //      )
    /// ```
    ///
    /// ### Comparison Operations
    ///
    /// ```swift
    /// let parser = FormulaParser("=A1 > 100")
    /// let ast = try parser.parse()
    /// // AST: .binaryOp(.greaterThan, .cellRef(A1), .number(100.0))
    /// ```
    ///
    /// ### Nested Functions
    ///
    /// ```swift
    /// let parser = FormulaParser("=AVERAGE(IF(A1:A5 > 0, A1:A5, 0))")
    /// let ast = try parser.parse()
    /// // AST: .functionCall("AVERAGE", [
    /// //        .functionCall("IF", [
    /// //          .binaryOp(.greaterThan, .range(A1, A5), .number(0.0)),
    /// //          .range(A1, A5),
    /// //          .number(0.0)
    /// //        ])
    /// //      ])
    /// ```
    ///
    /// ## Error Handling
    ///
    /// ```swift
    /// let parser = FormulaParser("=SUM(A1, A2")  // Missing closing paren
    /// do {
    ///     let ast = try parser.parse()
    /// } catch FormulaError.expectedRightParen {
    ///     print("Syntax error: missing closing parenthesis")
    /// }
    /// ```
    ///
    /// ## Common Errors
    ///
    /// - ``FormulaError/emptyFormula``: Formula is empty or contains only whitespace
    /// - ``FormulaError/unexpectedEndOfFormula``: Formula ends unexpectedly
    /// - ``FormulaError/unexpectedToken(_:)``: Encountered an unexpected token
    /// - ``FormulaError/expectedLeftParen(function:)``: Function name not followed by `(`
    /// - ``FormulaError/expectedRightParen``: Missing closing parenthesis
    /// - ``FormulaError/expectedCommaOrRightParen``: Invalid function argument syntax
    ///
    /// ## See Also
    /// - ``FormulaExpression``: The AST node types
    /// - ``FormulaEvaluator``: Evaluates the parsed AST
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

/// Errors that can occur during formula parsing and evaluation.
///
/// These errors represent both syntax errors during parsing and runtime errors during
/// evaluation. The error cases mirror Excel's error values (e.g., `#REF!`, `#DIV/0!`).
///
/// ## Topics
///
/// ### Parsing Errors
/// - ``emptyFormula``
/// - ``unexpectedEndOfFormula``
/// - ``unexpectedToken(_:)``
/// - ``expectedLeftParen(function:)``
/// - ``expectedRightParen``
/// - ``expectedCommaOrRightParen``
///
/// ### Reference Errors
/// - ``invalidCellReference(_:)``
/// - ``invalidRange(_:)``
/// - ``circularReference``
/// - ``nameNotFound(_:)``
///
/// ### Evaluation Errors
/// - ``divisionByZero``
/// - ``invalidArgumentCount(function:expected:got:)``
/// - ``invalidArgumentType(function:argument:)``
/// - ``evaluationError(_:)``
///
/// ## See Also
/// - ``FormulaParser``
/// - ``FormulaEvaluator``
public enum FormulaError: Error, Sendable, Equatable {
    /// The formula string is empty or contains only whitespace.
    case emptyFormula
    
    /// The formula ended unexpectedly (e.g., incomplete expression).
    case unexpectedEndOfFormula
    
    /// An unexpected or invalid token was encountered.
    ///
    /// - Parameter description: A description of the unexpected token
    case unexpectedToken(String)
    
    /// A function name was not followed by an opening parenthesis.
    ///
    /// - Parameter function: The name of the function
    case expectedLeftParen(function: String)
    
    /// Expected a closing parenthesis but found something else.
    case expectedRightParen
    
    /// Expected a comma or closing parenthesis in function arguments.
    case expectedCommaOrRightParen
    
    /// The cell reference string is invalid.
    ///
    /// - Parameter reference: The invalid reference string
    case invalidCellReference(String)
    
    /// The range string is invalid.
    ///
    /// - Parameter range: The invalid range string
    case invalidRange(String)
    
    /// Attempted to divide by zero.
    ///
    /// Corresponds to Excel's `#DIV/0!` error.
    case divisionByZero
    
    /// A function was called with the wrong number of arguments.
    ///
    /// - Parameters:
    ///   - function: The function name
    ///   - expected: The expected argument count
    ///   - got: The actual argument count
    case invalidArgumentCount(function: String, expected: Int, got: Int)
    
    /// A function argument has an invalid type.
    ///
    /// - Parameters:
    ///   - function: The function name
    ///   - argument: The argument index (0-based)
    case invalidArgumentType(function: String, argument: Int)
    
    /// A circular reference was detected.
    ///
    /// Corresponds to Excel's circular reference error.
    case circularReference
    
    /// A named range or function was not found.
    ///
    /// - Parameter name: The name that was not found
    ///
    /// Corresponds to Excel's `#NAME?` error.
    case nameNotFound(String)
    
    /// A general evaluation error occurred.
    ///
    /// - Parameter message: Description of the error
    case evaluationError(String)
}
