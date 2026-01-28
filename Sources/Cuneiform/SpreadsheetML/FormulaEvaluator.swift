import Foundation

/// Result of evaluating a formula expression
public enum FormulaValue: Sendable, Equatable {
    case number(Double)
    case string(String)
    case boolean(Bool)
    case error(String)
    case array([[FormulaValue]])
    
    public var asDouble: Double? {
        switch self {
        case .number(let n): return n
        case .boolean(let b): return b ? 1.0 : 0.0
        case .string(let s): return Double(s)
        default: return nil
        }
    }
    
    public var asString: String {
        switch self {
        case .number(let n): return String(n)
        case .string(let s): return s
        case .boolean(let b): return b ? "TRUE" : "FALSE"
        case .error(let e): return "#\(e)!"
        case .array: return "#ARRAY!"
        }
    }
    
    public var asBoolean: Bool? {
        switch self {
        case .boolean(let b): return b
        case .number(let n): return n != 0
        case .string(let s): return s.uppercased() == "TRUE"
        default: return nil
        }
    }
}

/// Evaluates parsed formula expressions in the context of a worksheet
public struct FormulaEvaluator: Sendable {
    private let cellResolver: @Sendable (CellReference) -> CellValue?
    
    /// Create evaluator with a cell resolver function
    public init(cellResolver: @escaping @Sendable (CellReference) -> CellValue?) {
        self.cellResolver = cellResolver
    }
    
    /// Evaluate a formula expression
    public func evaluate(_ expression: FormulaExpression) throws -> FormulaValue {
        switch expression {
        case .number(let value):
            return .number(value)
            
        case .string(let value):
            return .string(value)
            
        case .cellRef(let ref):
            guard let cell = cellResolver(ref) else {
                return .error("REF")
            }
            return cellValueToFormulaValue(cell)
            
        case .range(let start, let end):
            // Convert range to array
            var values: [[FormulaValue]] = []
            for row in start.row...end.row {
                var rowValues: [FormulaValue] = []
                // Iterate through column indices
                for col in start.columnIndex...end.columnIndex {
                    // Convert column index back to letter
                    let colLetter = columnLetterFrom(index: col)
                    let ref = CellReference(column: colLetter, row: row)
                    if let cell = cellResolver(ref) {
                        // Skip empty cells - matches Excel behavior where AVERAGE ignores empty cells
                        if case .empty = cell {
                            continue
                        }
                        rowValues.append(cellValueToFormulaValue(cell))
                    }
                    // Don't add anything for missing cells (skip them)
                }
                values.append(rowValues)
            }
            return .array(values)
            
        case .binaryOp(let op, let left, let right):
            return try evaluateBinaryOp(op, left, right)
            
        case .functionCall(let name, let args):
            return try evaluateFunction(name, args)
            
        case .error(let message):
            return .error(message)
        }
    }
    
    // MARK: - Binary Operations
    
    private func evaluateBinaryOp(_ op: FormulaExpression.BinaryOperator, _ left: FormulaExpression, _ right: FormulaExpression) throws -> FormulaValue {
        let leftVal = try evaluate(left)
        let rightVal = try evaluate(right)
        
        switch op {
        case .add:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(l + r)
            
        case .subtract:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(l - r)
            
        case .multiply:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(l * r)
            
        case .divide:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            guard r != 0 else {
                return .error("DIV/0")
            }
            return .number(l / r)
            
        case .power:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(pow(l, r))
            
        case .concat:
            return .string(leftVal.asString + rightVal.asString)
            
        case .equals:
            return .boolean(leftVal == rightVal)
            
        case .notEquals:
            return .boolean(leftVal != rightVal)
            
        case .lessThan:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l < r)
            
        case .greaterThan:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l > r)
            
        case .lessThanOrEqual:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l <= r)
            
        case .greaterThanOrEqual:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l >= r)
        }
    }
    
    // MARK: - Function Evaluation
    
    private func evaluateFunction(_ name: String, _ args: [FormulaExpression]) throws -> FormulaValue {
        let upperName = name.uppercased()
        
        switch upperName {
        case "SUM":
            return try evaluateSUM(args)
        case "AVERAGE":
            return try evaluateAVERAGE(args)
        case "IF":
            return try evaluateIF(args)
        case "VLOOKUP":
            return try evaluateVLOOKUP(args)
        case "HLOOKUP":
            return try evaluateHLOOKUP(args)
        case "INDEX":
            return try evaluateINDEX(args)
        case "MATCH":
            return try evaluateMATCH(args)
        case "XMATCH":
            return try evaluateXMATCH(args)
        case "OFFSET":
            return try evaluateOFFSET(args)
        case "INDIRECT":
            return try evaluateINDIRECT(args)
        case "CHOOSE":
            return try evaluateCHOOSE(args)
        case "CHOOSECOLS":
            return try evaluateCHOOSECOLS(args)
        case "CHOOSEROWS":
            return try evaluateCHOOSEROWS(args)
        case "TRANSPOSE":
            return try evaluateTRANSPOSE(args)
        case "ROWS":
            return try evaluateROWS(args)
        case "COLUMNS":
            return try evaluateCOLUMNS(args)
        case "MIN":
            return try evaluateMIN(args)
        case "MAX":
            return try evaluateMAX(args)
        case "COUNT":
            return try evaluateCOUNT(args)
        case "COUNTA":
            return try evaluateCOUNTA(args)
        case "AND":
            return try evaluateAND(args)
        case "OR":
            return try evaluateOR(args)
        case "NOT":
            return try evaluateNOT(args)
        case "LEN":
            return try evaluateLEN(args)
        case "UPPER":
            return try evaluateUPPER(args)
        case "LOWER":
            return try evaluateLOWER(args)
        case "CONCAT":
            return try evaluateCONCAT(args)
        case "ROUND":
            return try evaluateROUND(args)
        case "ABS":
            return try evaluateABS(args)
        case "MEDIAN":
            return try evaluateMEDIAN(args)
        // Date functions
        case "TODAY":
            return evaluateTODAY(args)
        case "NOW":
            return evaluateNOW(args)
        case "DATE":
            return try evaluateDATE(args)
        case "YEAR":
            return try evaluateYEAR(args)
        case "MONTH":
            return try evaluateMONTH(args)
        case "DAY":
            return try evaluateDAY(args)
        case "WEEKDAY":
            return try evaluateWEEKDAY(args)
        case "WEEKNUM":
            return try evaluateWEEKNUM(args)
        case "ISOWEEKNUM":
            return try evaluateISOWEEKNUM(args)
        case "EOMONTH":
            return try evaluateEOMONTH(args)
        case "EDATE":
            return try evaluateEDATE(args)
        case "NETWORKDAYS":
            return try evaluateNETWORKDAYS(args)
        case "WORKDAY":
            return try evaluateWORKDAY(args)
        case "DATEDIF":
            return try evaluateDATEDIF(args)
        case "YEARFRAC":
            return try evaluateYEARFRAC(args)
        case "TIME":
            return try evaluateTIME(args)
        case "TIMEVALUE":
            return try evaluateTIMEVALUE(args)
        case "HOUR":
            return try evaluateHOUR(args)
        case "MINUTE":
            return try evaluateMINUTE(args)
        case "SECOND":
            return try evaluateSECOND(args)
        // String functions
        case "LEFT":
            return try evaluateLEFT(args)
        case "RIGHT":
            return try evaluateRIGHT(args)
        case "MID":
            return try evaluateMID(args)
        case "TRIM":
            return try evaluateTRIM(args)
        case "PROPER":
            return try evaluatePROPER(args)
        case "CLEAN":
            return try evaluateCLEAN(args)
        case "CHAR":
            return try evaluateCHAR(args)
        case "CODE":
            return try evaluateCODE(args)
        case "EXACT":
            return try evaluateEXACT(args)
        case "REPLACE":
            return try evaluateREPLACE(args)
        case "REPT":
            return try evaluateREPT(args)
        case "VALUE":
            return try evaluateVALUE(args)
        case "TEXTBEFORE":
            return try evaluateTEXTBEFORE(args)
        case "TEXTAFTER":
            return try evaluateTEXTAFTER(args)
        case "TEXTSPLIT":
            return try evaluateTEXTSPLIT(args)
        // Conditional aggregates
        case "SUMIF":
            return try evaluateSUMIF(args)
        case "COUNTIF":
            return try evaluateCOUNTIF(args)
        case "IFERROR":
            return try evaluateIFERROR(args)
        // Additional math
        case "INT":
            return try evaluateINT(args)
        case "MOD":
            return try evaluateMOD(args)
        case "SQRT":
            return try evaluateSQRT(args)
        // Tier 2: Multi-criteria aggregates
        case "SUMIFS":
            return try evaluateSUMIFS(args)
        case "COUNTIFS":
            return try evaluateCOUNTIFS(args)
        case "AVERAGEIF":
            return try evaluateAVERAGEIF(args)
        // Tier 2: String search/manipulation
        case "FIND":
            return try evaluateFIND(args)
        case "SEARCH":
            return try evaluateSEARCH(args)
        case "SUBSTITUTE":
            return try evaluateSUBSTITUTE(args)
        case "CONCATENATE":
            return try evaluateCONCATENATE(args)
        case "TEXT":
            return try evaluateTEXT(args)
        // Tier 2: Type checking
        case "ISBLANK":
            return try evaluateISBLANK(args)
        case "ISNUMBER":
            return try evaluateISNUMBER(args)
        case "ISTEXT":
            return try evaluateISTEXT(args)
        case "ISERROR":
            return try evaluateISERROR(args)
        // Excel 365 high-priority functions
        case "XLOOKUP":
            return try evaluateXLOOKUP(args)
        case "TEXTJOIN":
            return try evaluateTEXTJOIN(args)
        case "IFS":
            return try evaluateIFS(args)
        case "SWITCH":
            return try evaluateSWITCH(args)
        case "MAXIFS":
            return try evaluateMAXIFS(args)
        case "MINIFS":
            return try evaluateMINIFS(args)
        case "AVERAGEIFS":
            return try evaluateAVERAGEIFS(args)
        // Statistical functions
        case "STDEV", "STDEV.S":
            return try evaluateSTDEV(args, sample: true)
        case "STDEV.P":
            return try evaluateSTDEV(args, sample: false)
        case "VAR", "VAR.S":
            return try evaluateVAR(args, sample: true)
        case "VAR.P":
            return try evaluateVAR(args, sample: false)
        case "PERCENTILE", "PERCENTILE.INC":
            return try evaluatePERCENTILE(args)
        case "QUARTILE", "QUARTILE.INC":
            return try evaluateQUARTILE(args)
        case "MODE", "MODE.SNGL":
            return try evaluateMODE(args)
        case "LARGE":
            return try evaluateLARGE(args)
        case "SMALL":
            return try evaluateSMALL(args)
        case "RANK", "RANK.EQ":
            return try evaluateRANK(args)
        case "CORREL":
            return try evaluateCORREL(args)
        // Math and trigonometric functions
        case "SIN":
            return try evaluateSIN(args)
        case "COS":
            return try evaluateCOS(args)
        case "TAN":
            return try evaluateTAN(args)
        case "ASIN":
            return try evaluateASIN(args)
        case "ACOS":
            return try evaluateACOS(args)
        case "ATAN":
            return try evaluateATAN(args)
        case "ATAN2":
            return try evaluateATAN2(args)
        case "PI":
            return evaluatePI(args)
        case "RADIANS":
            return try evaluateRADIANS(args)
        case "DEGREES":
            return try evaluateDEGREES(args)
        case "LOG":
            return try evaluateLOG(args)
        case "LOG10":
            return try evaluateLOG10(args)
        case "LN":
            return try evaluateLN(args)
        case "EXP":
            return try evaluateEXP(args)
        case "CEILING", "CEILING.MATH":
            return try evaluateCEILING(args)
        case "FLOOR", "FLOOR.MATH":
            return try evaluateFLOOR(args)
        case "TRUNC":
            return try evaluateTRUNC(args)
        case "SIGN":
            return try evaluateSIGN(args)
        case "FACT":
            return try evaluateFACT(args)
        case "SUMPRODUCT":
            return try evaluateSUMPRODUCT(args)
        case "GCD":
            return try evaluateGCD(args)
        case "LCM":
            return try evaluateLCM(args)
        // Financial functions
        case "PMT":
            return try evaluatePMT(args)
        case "PV":
            return try evaluatePV(args)
        case "FV":
            return try evaluateFV(args)
        case "RATE":
            return try evaluateRATE(args)
        case "NPER":
            return try evaluateNPER(args)
        case "IPMT":
            return try evaluateIPMT(args)
        case "PPMT":
            return try evaluatePPMT(args)
        case "NPV":
            return try evaluateNPV(args)
        case "IRR":
            return try evaluateIRR(args)
        case "XNPV":
            return try evaluateXNPV(args)
        case "XIRR":
            return try evaluateXIRR(args)
        default:
            return .error("NAME")
        }
    }
    
    // MARK: - Core Functions
    
    private func evaluateSUM(_ args: [FormulaExpression]) throws -> FormulaValue {
        var sum: Double = 0
        for arg in args {
            let val = try evaluate(arg)
            sum += flattenToNumbers(val).reduce(0, +)
        }
        return .number(sum)
    }
    
    private func evaluateAVERAGE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var sum: Double = 0
        var count = 0
        
        for arg in args {
            let val = try evaluate(arg)
            let numbers = flattenToNumbers(val)
            sum += numbers.reduce(0, +)
            count += numbers.count
        }
        
        guard count > 0 else {
            return .error("DIV/0")
        }
        
        return .number(sum / Double(count))
    }
    
    private func evaluateIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "IF", expected: 3, got: args.count)
        }
        
        let condition = try evaluate(args[0])
        guard let isTrue = condition.asBoolean else {
            return .error("VALUE")
        }
        
        if isTrue {
            return try evaluate(args[1])
        } else if args.count == 3 {
            return try evaluate(args[2])
        } else {
            return .boolean(false)
        }
    }
    
    private func evaluateVLOOKUP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "VLOOKUP", expected: 4, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let tableArray = try evaluate(args[1])
        let colIndex = try evaluate(args[2])
        let rangeLookup = args.count == 4 ? try evaluate(args[3]) : .boolean(true)
        
        guard case .array(let table) = tableArray else {
            return .error("VALUE")
        }
        
        guard let colNum = colIndex.asDouble, colNum >= 1, Int(colNum) <= (table.first?.count ?? 0) else {
            return .error("VALUE")
        }
        
        let col = Int(colNum) - 1
        let exactMatch = !(rangeLookup.asBoolean ?? true)
        
        // Search first column for lookup value
        for row in table {
            guard !row.isEmpty else { continue }
            
            if exactMatch {
                if row[0] == lookupValue {
                    return row[col]
                }
            } else {
                // Approximate match (assumes sorted)
                if let rowVal = row[0].asDouble, let lookupVal = lookupValue.asDouble {
                    if rowVal <= lookupVal {
                        return row[col]
                    }
                }
            }
        }
        
        return .error("N/A")
    }
    
    private func evaluateINDEX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "INDEX", expected: 3, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let rowNum = try evaluate(args[1])
        
        guard case .array(let array) = arrayVal else {
            return .error("VALUE")
        }
        
        guard let row = rowNum.asDouble, row >= 1, Int(row) <= array.count else {
            return .error("REF")
        }
        
        let rowIndex = Int(row) - 1
        
        if args.count == 3 {
            let colNum = try evaluate(args[2])
            guard let col = colNum.asDouble, col >= 1, Int(col) <= array[rowIndex].count else {
                return .error("REF")
            }
            let colIndex = Int(col) - 1
            return array[rowIndex][colIndex]
        } else {
            // Return entire row as array if no column specified
            return .array([array[rowIndex]])
        }
    }
    
    private func evaluateMATCH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "MATCH", expected: 3, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let lookupArray = try evaluate(args[1])
        let matchType = args.count == 3 ? try evaluate(args[2]) : .number(1)
        
        guard case .array(let array) = lookupArray else {
            return .error("VALUE")
        }
        
        let matchMode = Int(matchType.asDouble ?? 1)
        
        // Flatten array to 1D
        let values = array.flatMap { $0 }
        
        for (index, value) in values.enumerated() {
            if matchMode == 0 {
                // Exact match
                if value == lookupValue {
                    return .number(Double(index + 1))
                }
            } else if matchMode == 1 {
                // Less than or equal (assumes sorted ascending)
                if let valNum = value.asDouble, let lookupNum = lookupValue.asDouble {
                    if valNum <= lookupNum {
                        return .number(Double(index + 1))
                    }
                }
            }
        }
        
        return .error("N/A")
    }
    
    private func evaluateMIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard let min = numbers.min() else {
            return .error("VALUE")
        }
        
        return .number(min)
    }
    
    private func evaluateMAX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard let max = numbers.max() else {
            return .error("VALUE")
        }
        
        return .number(max)
    }
    
    private func evaluateCOUNT(_ args: [FormulaExpression]) throws -> FormulaValue {
        var count = 0
        for arg in args {
            let val = try evaluate(arg)
            count += flattenToNumbers(val).count
        }
        return .number(Double(count))
    }
    
    private func evaluateCOUNTA(_ args: [FormulaExpression]) throws -> FormulaValue {
        var count = 0
        for arg in args {
            let val = try evaluate(arg)
            count += flattenToAll(val).count
        }
        return .number(Double(count))
    }
    
    // MARK: - Logical Functions
    
    private func evaluateAND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        for arg in args {
            let val = try evaluate(arg)
            let bools = flattenToBooleans(val)
            for b in bools {
                if !b { return .boolean(false) }
            }
        }
        return .boolean(true)
    }
    
    private func evaluateOR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        for arg in args {
            let val = try evaluate(arg)
            let bools = flattenToBooleans(val)
            for b in bools {
                if b { return .boolean(true) }
            }
        }
        return .boolean(false)
    }
    
    private func evaluateNOT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "NOT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let bool = val.asBoolean else {
            return .error("VALUE")
        }
        return .boolean(!bool)
    }
    
    // MARK: - String Functions
    
    private func evaluateLEN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "LEN", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        return .number(Double(str.count))
    }
    
    private func evaluateUPPER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "UPPER", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        return .string(val.asString.uppercased())
    }
    
    private func evaluateLOWER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "LOWER", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        return .string(val.asString.lowercased())
    }
    
    private func evaluateCONCAT(_ args: [FormulaExpression]) throws -> FormulaValue {
        var result = ""
        for arg in args {
            let val = try evaluate(arg)
            result += val.asString
        }
        return .string(result)
    }
    
    // MARK: - Math Functions
    
    private func evaluateROUND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "ROUND", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        let digits = args.count == 2 ? Int(try evaluate(args[1]).asDouble ?? 0) : 0
        let multiplier = pow(10.0, Double(digits))
        return .number((num * multiplier).rounded() / multiplier)
    }
    
    private func evaluateABS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ABS", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        return .number(abs(num))
    }
    
    private func evaluateMEDIAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("VALUE")
        }
        
        let sorted = numbers.sorted()
        let count = sorted.count
        
        if count % 2 == 0 {
            return .number((sorted[count / 2 - 1] + sorted[count / 2]) / 2.0)
        } else {
            return .number(sorted[count / 2])
        }
    }
    
    // MARK: - Helpers
    
    private func cellValueToFormulaValue(_ cell: CellValue) -> FormulaValue {
        switch cell {
        case .empty:
            return .number(0)
        case .number(let n):
            return .number(n)
        case .text(let s):
            return .string(s)
        case .boolean(let b):
            return .boolean(b)
        case .date(let d):
            // Date is stored as string in ISO format; return as string
            return .string(d)
        case .error(let e):
            return .error(e)
        case .richText(let rt):
            return .string(rt.plainText)
        }
    }
    
    private func flattenToNumbers(_ value: FormulaValue) -> [Double] {
        switch value {
        case .number(let n):
            return [n]
        case .boolean(let b):
            return [b ? 1.0 : 0.0]
        case .array(let arr):
            return arr.flatMap { row in
                row.compactMap { $0.asDouble }
            }
        default:
            return []
        }
    }
    
    private func flattenToAll(_ value: FormulaValue) -> [FormulaValue] {
        switch value {
        case .array(let arr):
            return arr.flatMap { $0 }
        case .error:
            return []
        default:
            return [value]
        }
    }
    
    private func flattenToBooleans(_ value: FormulaValue) -> [Bool] {
        switch value {
        case .boolean(let b):
            return [b]
        case .number(let n):
            return [n != 0]
        case .array(let arr):
            return arr.flatMap { row in
                row.compactMap { $0.asBoolean }
            }
        default:
            return []
        }
    }
    
    // MARK: - Date Functions
    
    /// Returns today's date as an Excel serial number
    private func evaluateTODAY(_ args: [FormulaExpression]) -> FormulaValue {
        // Excel serial date: days since 1899-12-30 (accounting for the 1900 leap year bug)
        let now = Date()
        let calendar = Calendar(identifier: .gregorian)
        let components = calendar.dateComponents([.year, .month, .day], from: now)
        return dateToSerial(year: components.year!, month: components.month!, day: components.day!)
    }
    
    /// Returns current date and time as an Excel serial number
    private func evaluateNOW(_ args: [FormulaExpression]) -> FormulaValue {
        let now = Date()
        let calendar = Calendar(identifier: .gregorian)
        let components = calendar.dateComponents([.year, .month, .day, .hour, .minute, .second], from: now)
        
        // Get the date portion
        guard case .number(let datePart) = dateToSerial(year: components.year!, month: components.month!, day: components.day!) else {
            return .error("VALUE")
        }
        
        // Add time as fraction of day
        let timeFraction = Double(components.hour!) / 24.0 +
                          Double(components.minute!) / 1440.0 +
                          Double(components.second!) / 86400.0
        
        return .number(datePart + timeFraction)
    }
    
    /// Creates a date serial number from year, month, day
    private func evaluateDATE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "DATE", expected: 3, got: args.count)
        }
        
        let yearVal = try evaluate(args[0])
        let monthVal = try evaluate(args[1])
        let dayVal = try evaluate(args[2])
        
        guard let year = yearVal.asDouble,
              let month = monthVal.asDouble,
              let day = dayVal.asDouble else {
            return .error("VALUE")
        }
        
        return dateToSerial(year: Int(year), month: Int(month), day: Int(day))
    }
    
    /// Extracts year from a date serial number
    private func evaluateYEAR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "YEAR", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let serial = val.asDouble else {
            return .error("VALUE")
        }
        
        let (year, _, _) = serialToDate(serial: serial)
        return .number(Double(year))
    }
    
    /// Extracts month from a date serial number
    private func evaluateMONTH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "MONTH", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let serial = val.asDouble else {
            return .error("VALUE")
        }
        
        let (_, month, _) = serialToDate(serial: serial)
        return .number(Double(month))
    }
    
    /// Extracts day from a date serial number
    private func evaluateDAY(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "DAY", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let serial = val.asDouble else {
            return .error("VALUE")
        }
        
        let (_, _, day) = serialToDate(serial: serial)
        return .number(Double(day))
    }
    
    /// Convert year/month/day to Excel serial number
    private func dateToSerial(year: Int, month: Int, day: Int) -> FormulaValue {
        // Handle year values 0-99 as 1900-1999, 100-9999 as-is (Excel behavior)
        var adjustedYear = year
        if year >= 0 && year <= 99 {
            adjustedYear = 1900 + year
        }
        
        // Handle month overflow/underflow
        var adjMonth = month
        var adjYear = adjustedYear
        while adjMonth > 12 {
            adjMonth -= 12
            adjYear += 1
        }
        while adjMonth < 1 {
            adjMonth += 12
            adjYear -= 1
        }
        
        // Create date components
        var components = DateComponents()
        components.year = adjYear
        components.month = adjMonth
        components.day = day
        
        let calendar = Calendar(identifier: .gregorian)
        guard let date = calendar.date(from: components) else {
            return .error("VALUE")
        }
        
        // Excel epoch: 1899-12-30 (due to the 1900 leap year bug)
        var epochComponents = DateComponents()
        epochComponents.year = 1899
        epochComponents.month = 12
        epochComponents.day = 30
        
        guard let epoch = calendar.date(from: epochComponents) else {
            return .error("VALUE")
        }
        
        let days = calendar.dateComponents([.day], from: epoch, to: date).day!
        
        // Excel incorrectly treats 1900 as a leap year, so dates after Feb 28, 1900 need +1
        if days > 59 {
            return .number(Double(days + 1))
        }
        return .number(Double(days))
    }
    
    /// Convert Excel serial number to year/month/day
    private func serialToDate(serial: Double) -> (year: Int, month: Int, day: Int) {
        var adjustedSerial = Int(serial)
        
        // Account for Excel's 1900 leap year bug
        if adjustedSerial > 60 {
            adjustedSerial -= 1
        }
        
        var components = DateComponents()
        components.year = 1899
        components.month = 12
        components.day = 30
        
        let calendar = Calendar(identifier: .gregorian)
        guard let epoch = calendar.date(from: components),
              let date = calendar.date(byAdding: .day, value: adjustedSerial, to: epoch) else {
            return (1900, 1, 1)
        }
        
        let resultComponents = calendar.dateComponents([.year, .month, .day], from: date)
        return (resultComponents.year!, resultComponents.month!, resultComponents.day!)
    }
    
    // MARK: - Extended String Functions
    
    /// Returns leftmost characters from a string
    private func evaluateLEFT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "LEFT", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        let numChars: Int
        if args.count == 2 {
            guard let n = try evaluate(args[1]).asDouble, n >= 0 else {
                return .error("VALUE")
            }
            numChars = Int(n)
        } else {
            numChars = 1
        }
        
        if numChars >= str.count {
            return .string(str)
        }
        
        let endIndex = str.index(str.startIndex, offsetBy: numChars)
        return .string(String(str[..<endIndex]))
    }
    
    /// Returns rightmost characters from a string
    private func evaluateRIGHT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "RIGHT", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        let numChars: Int
        if args.count == 2 {
            guard let n = try evaluate(args[1]).asDouble, n >= 0 else {
                return .error("VALUE")
            }
            numChars = Int(n)
        } else {
            numChars = 1
        }
        
        if numChars >= str.count {
            return .string(str)
        }
        
        let startIndex = str.index(str.endIndex, offsetBy: -numChars)
        return .string(String(str[startIndex...]))
    }
    
    /// Returns characters from the middle of a string
    private func evaluateMID(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "MID", expected: 3, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        guard let startNum = try evaluate(args[1]).asDouble,
              let numChars = try evaluate(args[2]).asDouble,
              startNum >= 1, numChars >= 0 else {
            return .error("VALUE")
        }
        
        let startIndex = Int(startNum) - 1 // Excel is 1-based
        let length = Int(numChars)
        
        guard startIndex < str.count else {
            return .string("")
        }
        
        let start = str.index(str.startIndex, offsetBy: startIndex)
        let end = str.index(start, offsetBy: min(length, str.count - startIndex))
        return .string(String(str[start..<end]))
    }
    
    /// Removes leading and trailing spaces, and reduces internal spaces to single spaces
    private func evaluateTRIM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "TRIM", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        // Excel's TRIM removes leading/trailing spaces and collapses multiple spaces to single
        let components = str.split(separator: " ", omittingEmptySubsequences: true)
        return .string(components.joined(separator: " "))
    }
    
    // MARK: - Conditional Aggregate Functions
    
    /// Sum cells that match a criteria
    private func evaluateSUMIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "SUMIF", expected: 3, got: args.count)
        }
        
        let rangeVal = try evaluate(args[0])
        let criteriaVal = try evaluate(args[1])
        
        // If sum_range is provided, use it; otherwise use the criteria range
        let sumRangeVal: FormulaValue
        if args.count == 3 {
            sumRangeVal = try evaluate(args[2])
        } else {
            sumRangeVal = rangeVal
        }
        
        guard case .array(let criteriaRange) = rangeVal else {
            // Single cell
            if matchesCriteria(rangeVal, criteriaVal) {
                return sumRangeVal.asDouble.map { .number($0) } ?? .number(0)
            }
            return .number(0)
        }
        
        let sumRange: [[FormulaValue]]
        if case .array(let arr) = sumRangeVal {
            sumRange = arr
        } else {
            sumRange = [[sumRangeVal]]
        }
        
        var sum: Double = 0
        for (rowIdx, row) in criteriaRange.enumerated() {
            for (colIdx, cell) in row.enumerated() {
                if matchesCriteria(cell, criteriaVal) {
                    // Get corresponding value from sum range
                    if rowIdx < sumRange.count && colIdx < sumRange[rowIdx].count {
                        if let num = sumRange[rowIdx][colIdx].asDouble {
                            sum += num
                        }
                    }
                }
            }
        }
        
        return .number(sum)
    }
    
    /// Count cells that match a criteria
    private func evaluateCOUNTIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "COUNTIF", expected: 2, got: args.count)
        }
        
        let rangeVal = try evaluate(args[0])
        let criteriaVal = try evaluate(args[1])
        
        guard case .array(let range) = rangeVal else {
            // Single cell
            return .number(matchesCriteria(rangeVal, criteriaVal) ? 1 : 0)
        }
        
        var count = 0
        for row in range {
            for cell in row {
                if matchesCriteria(cell, criteriaVal) {
                    count += 1
                }
            }
        }
        
        return .number(Double(count))
    }
    
    /// Returns value if no error, otherwise returns alternate value
    private func evaluateIFERROR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "IFERROR", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        if case .error = val {
            return try evaluate(args[1])
        }
        
        return val
    }
    
    /// Helper to check if a value matches a criteria
    private func matchesCriteria(_ value: FormulaValue, _ criteria: FormulaValue) -> Bool {
        let criteriaStr = criteria.asString
        
        // Check for comparison operators in criteria string
        if criteriaStr.hasPrefix(">=") {
            guard let threshold = Double(String(criteriaStr.dropFirst(2))),
                  let val = value.asDouble else { return false }
            return val >= threshold
        } else if criteriaStr.hasPrefix("<=") {
            guard let threshold = Double(String(criteriaStr.dropFirst(2))),
                  let val = value.asDouble else { return false }
            return val <= threshold
        } else if criteriaStr.hasPrefix("<>") {
            let compareStr = String(criteriaStr.dropFirst(2))
            if let threshold = Double(compareStr), let val = value.asDouble {
                return val != threshold
            }
            return value.asString != compareStr
        } else if criteriaStr.hasPrefix(">") {
            guard let threshold = Double(String(criteriaStr.dropFirst(1))),
                  let val = value.asDouble else { return false }
            return val > threshold
        } else if criteriaStr.hasPrefix("<") {
            guard let threshold = Double(String(criteriaStr.dropFirst(1))),
                  let val = value.asDouble else { return false }
            return val < threshold
        } else if criteriaStr.hasPrefix("=") {
            let compareStr = String(criteriaStr.dropFirst(1))
            if let threshold = Double(compareStr), let val = value.asDouble {
                return val == threshold
            }
            return value.asString == compareStr
        } else if criteriaStr.contains("*") || criteriaStr.contains("?") {
            // Wildcard matching
            return wildcardMatch(value.asString, pattern: criteriaStr)
        } else {
            // Direct comparison
            if let criteriaNum = criteria.asDouble, let valNum = value.asDouble {
                return valNum == criteriaNum
            }
            return value.asString == criteriaStr
        }
    }
    
    /// Simple wildcard matching (* and ?)
    private func wildcardMatch(_ string: String, pattern: String) -> Bool {
        let regexPattern = "^" + NSRegularExpression.escapedPattern(for: pattern)
            .replacingOccurrences(of: "\\*", with: ".*")
            .replacingOccurrences(of: "\\?", with: ".") + "$"
        
        guard let regex = try? NSRegularExpression(pattern: regexPattern, options: .caseInsensitive) else {
            return false
        }
        
        let range = NSRange(string.startIndex..., in: string)
        return regex.firstMatch(in: string, options: [], range: range) != nil
    }
    
    // MARK: - Additional Math Functions
    
    /// Truncates a number to an integer
    private func evaluateINT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "INT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        // Excel's INT rounds toward negative infinity
        return .number(floor(num))
    }
    
    /// Returns the remainder after division
    private func evaluateMOD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "MOD", expected: 2, got: args.count)
        }
        
        let numVal = try evaluate(args[0])
        let divisorVal = try evaluate(args[1])
        
        guard let num = numVal.asDouble, let divisor = divisorVal.asDouble else {
            return .error("VALUE")
        }
        
        guard divisor != 0 else {
            return .error("DIV/0")
        }
        
        // Excel's MOD: result has the same sign as divisor
        let result = num - divisor * floor(num / divisor)
        return .number(result)
    }
    
    /// Returns the square root of a number
    private func evaluateSQRT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SQRT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        guard num >= 0 else {
            return .error("NUM")
        }
        
        return .number(sqrt(num))
    }
    
    // MARK: - Tier 2: Multi-Criteria Aggregate Functions
    
    /// SUMIFS - Sum cells that match multiple criteria
    private func evaluateSUMIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
        guard args.count >= 3 && args.count % 2 == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SUMIFS", expected: 3, got: args.count)
        }
        
        let sumRangeVal = try evaluate(args[0])
        let sumRange: [[FormulaValue]]
        if case .array(let arr) = sumRangeVal {
            sumRange = arr
        } else {
            sumRange = [[sumRangeVal]]
        }
        
        // Collect criteria pairs
        var criteriaPairs: [(range: [[FormulaValue]], criteria: FormulaValue)] = []
        var i = 1
        while i < args.count {
            let rangeVal = try evaluate(args[i])
            let criteriaVal = try evaluate(args[i + 1])
            
            let range: [[FormulaValue]]
            if case .array(let arr) = rangeVal {
                range = arr
            } else {
                range = [[rangeVal]]
            }
            criteriaPairs.append((range: range, criteria: criteriaVal))
            i += 2
        }
        
        var sum: Double = 0
        
        // Iterate through sum range
        for (rowIdx, row) in sumRange.enumerated() {
            for (colIdx, _) in row.enumerated() {
                // Check all criteria
                var allMatch = true
                for pair in criteriaPairs {
                    if rowIdx < pair.range.count && colIdx < pair.range[rowIdx].count {
                        if !matchesCriteria(pair.range[rowIdx][colIdx], pair.criteria) {
                            allMatch = false
                            break
                        }
                    } else {
                        allMatch = false
                        break
                    }
                }
                
                if allMatch {
                    if rowIdx < sumRange.count && colIdx < sumRange[rowIdx].count {
                        if let num = sumRange[rowIdx][colIdx].asDouble {
                            sum += num
                        }
                    }
                }
            }
        }
        
        return .number(sum)
    }
    
    /// COUNTIFS - Count cells that match multiple criteria
    private func evaluateCOUNTIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        // COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
        guard args.count >= 2 && args.count % 2 == 0 else {
            throw FormulaError.invalidArgumentCount(function: "COUNTIFS", expected: 2, got: args.count)
        }
        
        // Get first range to determine dimensions
        let firstRangeVal = try evaluate(args[0])
        let firstRange: [[FormulaValue]]
        if case .array(let arr) = firstRangeVal {
            firstRange = arr
        } else {
            firstRange = [[firstRangeVal]]
        }
        
        // Collect all criteria pairs
        var criteriaPairs: [(range: [[FormulaValue]], criteria: FormulaValue)] = []
        var i = 0
        while i < args.count {
            let rangeVal = try evaluate(args[i])
            let criteriaVal = try evaluate(args[i + 1])
            
            let range: [[FormulaValue]]
            if case .array(let arr) = rangeVal {
                range = arr
            } else {
                range = [[rangeVal]]
            }
            criteriaPairs.append((range: range, criteria: criteriaVal))
            i += 2
        }
        
        var count = 0
        
        // Iterate through first range dimensions
        for rowIdx in 0..<firstRange.count {
            for colIdx in 0..<(firstRange.first?.count ?? 0) {
                var allMatch = true
                for pair in criteriaPairs {
                    if rowIdx < pair.range.count && colIdx < pair.range[rowIdx].count {
                        if !matchesCriteria(pair.range[rowIdx][colIdx], pair.criteria) {
                            allMatch = false
                            break
                        }
                    } else {
                        allMatch = false
                        break
                    }
                }
                
                if allMatch {
                    count += 1
                }
            }
        }
        
        return .number(Double(count))
    }
    
    /// AVERAGEIF - Average cells that match a criteria
    private func evaluateAVERAGEIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "AVERAGEIF", expected: 3, got: args.count)
        }
        
        let rangeVal = try evaluate(args[0])
        let criteriaVal = try evaluate(args[1])
        
        // If average_range is provided, use it; otherwise use the criteria range
        let avgRangeVal: FormulaValue
        if args.count == 3 {
            avgRangeVal = try evaluate(args[2])
        } else {
            avgRangeVal = rangeVal
        }
        
        guard case .array(let criteriaRange) = rangeVal else {
            // Single cell
            if matchesCriteria(rangeVal, criteriaVal) {
                return avgRangeVal.asDouble.map { .number($0) } ?? .error("DIV/0")
            }
            return .error("DIV/0")
        }
        
        let avgRange: [[FormulaValue]]
        if case .array(let arr) = avgRangeVal {
            avgRange = arr
        } else {
            avgRange = [[avgRangeVal]]
        }
        
        var sum: Double = 0
        var count = 0
        
        for (rowIdx, row) in criteriaRange.enumerated() {
            for (colIdx, cell) in row.enumerated() {
                if matchesCriteria(cell, criteriaVal) {
                    if rowIdx < avgRange.count && colIdx < avgRange[rowIdx].count {
                        if let num = avgRange[rowIdx][colIdx].asDouble {
                            sum += num
                            count += 1
                        }
                    }
                }
            }
        }
        
        guard count > 0 else {
            return .error("DIV/0")
        }
        
        return .number(sum / Double(count))
    }
    
    // MARK: - Tier 2: String Search & Manipulation
    
    /// FIND - Find text within text (case-sensitive)
    private func evaluateFIND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "FIND", expected: 2, got: args.count)
        }
        
        let findTextVal = try evaluate(args[0])
        let withinTextVal = try evaluate(args[1])
        
        let findText = findTextVal.asString
        let withinText = withinTextVal.asString
        
        let startNum: Int
        if args.count == 3 {
            guard let n = try evaluate(args[2]).asDouble, n >= 1 else {
                return .error("VALUE")
            }
            startNum = Int(n)
        } else {
            startNum = 1
        }
        
        guard startNum <= withinText.count else {
            return .error("VALUE")
        }
        
        let searchStartIndex = withinText.index(withinText.startIndex, offsetBy: startNum - 1)
        let searchRange = searchStartIndex..<withinText.endIndex
        
        if let range = withinText.range(of: findText, range: searchRange) {
            let position = withinText.distance(from: withinText.startIndex, to: range.lowerBound) + 1
            return .number(Double(position))
        }
        
        return .error("VALUE")
    }
    
    /// SEARCH - Find text within text (case-insensitive, supports wildcards)
    private func evaluateSEARCH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "SEARCH", expected: 2, got: args.count)
        }
        
        let findTextVal = try evaluate(args[0])
        let withinTextVal = try evaluate(args[1])
        
        let findText = findTextVal.asString
        let withinText = withinTextVal.asString
        
        let startNum: Int
        if args.count == 3 {
            guard let n = try evaluate(args[2]).asDouble, n >= 1 else {
                return .error("VALUE")
            }
            startNum = Int(n)
        } else {
            startNum = 1
        }
        
        guard startNum <= withinText.count else {
            return .error("VALUE")
        }
        
        let searchStartIndex = withinText.index(withinText.startIndex, offsetBy: startNum - 1)
        let searchString = String(withinText[searchStartIndex...])
        
        // Handle wildcards in find text
        if findText.contains("*") || findText.contains("?") {
            let pattern = "^(.*?)(" + NSRegularExpression.escapedPattern(for: findText)
                .replacingOccurrences(of: "\\*", with: ".*")
                .replacingOccurrences(of: "\\?", with: ".") + ")"
            
            if let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
               let match = regex.firstMatch(in: searchString, options: [], range: NSRange(searchString.startIndex..., in: searchString)) {
                let position = startNum + match.range(at: 1).length
                return .number(Double(position))
            }
        } else {
            // Simple case-insensitive search
            if let range = searchString.range(of: findText, options: .caseInsensitive) {
                let position = startNum + searchString.distance(from: searchString.startIndex, to: range.lowerBound)
                return .number(Double(position))
            }
        }
        
        return .error("VALUE")
    }
    
    /// SUBSTITUTE - Replace occurrences of text
    private func evaluateSUBSTITUTE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "SUBSTITUTE", expected: 3, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        let oldTextVal = try evaluate(args[1])
        let newTextVal = try evaluate(args[2])
        
        var text = textVal.asString
        let oldText = oldTextVal.asString
        let newText = newTextVal.asString
        
        if args.count == 4 {
            // Replace only the nth occurrence
            guard let n = try evaluate(args[3]).asDouble, n >= 1 else {
                return .error("VALUE")
            }
            let instanceNum = Int(n)
            
            var occurrence = 0
            var searchStart = text.startIndex
            
            while let range = text.range(of: oldText, range: searchStart..<text.endIndex) {
                occurrence += 1
                if occurrence == instanceNum {
                    text.replaceSubrange(range, with: newText)
                    break
                }
                searchStart = range.upperBound
            }
        } else {
            // Replace all occurrences
            text = text.replacingOccurrences(of: oldText, with: newText)
        }
        
        return .string(text)
    }
    
    /// CONCATENATE - Join multiple text strings
    private func evaluateCONCATENATE(_ args: [FormulaExpression]) throws -> FormulaValue {
        var result = ""
        for arg in args {
            let val = try evaluate(arg)
            result += val.asString
        }
        return .string(result)
    }
    
    /// TEXT - Format a number as text with a format string
    private func evaluateTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "TEXT", expected: 2, got: args.count)
        }
        
        let valueArg = try evaluate(args[0])
        let formatArg = try evaluate(args[1])
        
        guard let num = valueArg.asDouble else {
            return .string(valueArg.asString)
        }
        
        let formatStr = formatArg.asString
        
        // Handle common Excel format patterns
        let result: String
        
        if formatStr.lowercased().contains("0.00%") || formatStr == "0%" {
            // Percentage format
            let percentage = num * 100
            let decimals = formatStr.contains(".00") ? 2 : 0
            result = String(format: "%.\(decimals)f%%", percentage)
        } else if formatStr.contains("$") || formatStr.lowercased().contains("#,##0") {
            // Currency/number format with thousands separator
            let formatter = NumberFormatter()
            formatter.numberStyle = .decimal
            formatter.groupingSeparator = ","
            
            // Count decimal places
            if let decimalPart = formatStr.split(separator: ".").last {
                formatter.minimumFractionDigits = decimalPart.filter { $0 == "0" }.count
                formatter.maximumFractionDigits = decimalPart.filter { $0 == "0" }.count
            } else {
                formatter.minimumFractionDigits = 0
                formatter.maximumFractionDigits = 0
            }
            
            let formatted = formatter.string(from: NSNumber(value: num)) ?? String(num)
            if formatStr.contains("$") {
                result = "$" + formatted
            } else {
                result = formatted
            }
        } else if formatStr.lowercased().contains("yyyy") || formatStr.lowercased().contains("mm") || formatStr.lowercased().contains("dd") {
            // Date format - convert serial to date string
            let (year, month, day) = serialToDate(serial: num)
            var output = formatStr
            output = output.replacingOccurrences(of: "yyyy", with: String(format: "%04d", year))
            output = output.replacingOccurrences(of: "YYYY", with: String(format: "%04d", year))
            output = output.replacingOccurrences(of: "yy", with: String(format: "%02d", year % 100))
            output = output.replacingOccurrences(of: "YY", with: String(format: "%02d", year % 100))
            output = output.replacingOccurrences(of: "mm", with: String(format: "%02d", month))
            output = output.replacingOccurrences(of: "MM", with: String(format: "%02d", month))
            output = output.replacingOccurrences(of: "m", with: String(month))
            output = output.replacingOccurrences(of: "M", with: String(month))
            output = output.replacingOccurrences(of: "dd", with: String(format: "%02d", day))
            output = output.replacingOccurrences(of: "DD", with: String(format: "%02d", day))
            output = output.replacingOccurrences(of: "d", with: String(day))
            output = output.replacingOccurrences(of: "D", with: String(day))
            result = output
        } else if formatStr.contains("0") || formatStr.contains("#") {
            // General number format
            let decimalCount = formatStr.split(separator: ".").last?.filter { $0 == "0" }.count ?? 0
            result = String(format: "%.\(decimalCount)f", num)
        } else {
            result = String(num)
        }
        
        return .string(result)
    }
    
    // MARK: - Tier 2: Type Checking Functions
    
    /// ISBLANK - Check if a cell is blank
    private func evaluateISBLANK(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISBLANK", expected: 1, got: args.count)
        }
        
        // For cell references, check if cell is empty
        if case .cellRef(let ref) = args[0] {
            if let value = cellResolver(ref) {
                switch value {
                case .empty:
                    return .boolean(true)
                default:
                    return .boolean(false)
                }
            }
            // Cell not found = blank
            return .boolean(true)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .string(let s):
            return .boolean(s.isEmpty)
        case .error:
            return .boolean(false)
        default:
            return .boolean(false)
        }
    }
    
    /// ISNUMBER - Check if a value is a number
    private func evaluateISNUMBER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISNUMBER", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .number:
            return .boolean(true)
        default:
            return .boolean(false)
        }
    }
    
    /// ISTEXT - Check if a value is text
    private func evaluateISTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISTEXT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .string:
            return .boolean(true)
        default:
            return .boolean(false)
        }
    }
    
    /// ISERROR - Check if a value is an error
    private func evaluateISERROR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISERROR", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .error:
            return .boolean(true)
        default:
            return .boolean(false)
        }
    }
    
    // MARK: - Excel 365 High-Priority Functions
    
    /// XLOOKUP - Modern lookup function (Excel 365)
    /// Syntax: XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
    private func evaluateXLOOKUP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 6 else {
            throw FormulaError.invalidArgumentCount(function: "XLOOKUP", expected: 3, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let lookupArray = try evaluate(args[1])
        let returnArray = try evaluate(args[2])
        let ifNotFound = args.count > 3 ? try evaluate(args[3]) : FormulaValue.error("N/A")
        // match_mode and search_mode not fully implemented yet (defaults: exact match, first-to-last)
        
        // Extract lookup array values
        guard case .array(let lookupRows) = lookupArray else {
            return .error("VALUE")
        }
        let lookupValues = lookupRows.flatMap { $0 }
        
        // Extract return array values
        guard case .array(let returnRows) = returnArray else {
            return .error("VALUE")
        }
        let returnValues = returnRows.flatMap { $0 }
        
        // Ensure arrays are same length
        guard lookupValues.count == returnValues.count else {
            return .error("VALUE")
        }
        
        // Find match
        for (index, value) in lookupValues.enumerated() {
            if value == lookupValue {
                return returnValues[index]
            }
        }
        
        // Not found
        return ifNotFound
    }
    
    /// TEXTJOIN - Join text with delimiter (Excel 365)
    /// Syntax: TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
    private func evaluateTEXTJOIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            throw FormulaError.invalidArgumentCount(function: "TEXTJOIN", expected: 3, got: args.count)
        }
        
        let delimiter = try evaluate(args[0]).asString
        let ignoreEmpty = try evaluate(args[1]).asBoolean ?? false
        
        var parts: [String] = []
        for i in 2..<args.count {
            let val = try evaluate(args[i])
            if case .array(let rows) = val {
                for row in rows {
                    for cell in row {
                        let str = cell.asString
                        if !ignoreEmpty || !str.isEmpty {
                            parts.append(str)
                        }
                    }
                }
            } else {
                let str = val.asString
                if !ignoreEmpty || !str.isEmpty {
                    parts.append(str)
                }
            }
        }
        
        return .string(parts.joined(separator: delimiter))
    }
    
    /// IFS - Multiple IF conditions (Excel 365)
    /// Syntax: IFS(condition1, value1, [condition2, value2], ...)
    private func evaluateIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count % 2 == 0 else {
            return .error("VALUE")
        }
        
        for i in stride(from: 0, to: args.count, by: 2) {
            let condition = try evaluate(args[i])
            if condition.asBoolean == true {
                return try evaluate(args[i + 1])
            }
        }
        
        // No condition matched
        return .error("N/A")
    }
    
    /// SWITCH - Match value and return result (Excel 365)
    /// Syntax: SWITCH(expression, value1, result1, [value2, result2], ..., [default])
    private func evaluateSWITCH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            throw FormulaError.invalidArgumentCount(function: "SWITCH", expected: 3, got: args.count)
        }
        
        let expression = try evaluate(args[0])
        
        // Check value/result pairs
        var i = 1
        while i < args.count - 1 {
            let value = try evaluate(args[i])
            if value == expression {
                return try evaluate(args[i + 1])
            }
            i += 2
        }
        
        // If odd number of args after expression, last is default
        if args.count % 2 == 0 {
            return try evaluate(args[args.count - 1])
        }
        
        // No match and no default
        return .error("N/A")
    }
    
    /// MAXIFS - Maximum value with multiple criteria
    /// Syntax: MAXIFS(max_range, criteria_range1, criterion1, [criteria_range2, criterion2], ...)
    private func evaluateMAXIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count % 2 == 1 else {
            return .error("VALUE")
        }
        
        let maxRange = try evaluate(args[0])
        guard case .array(let maxRows) = maxRange else {
            return .error("VALUE")
        }
        let maxValues = maxRows.flatMap { $0 }
        
        // Build list of matching indices
        var matchingIndices = Set(0..<maxValues.count)
        
        // Process each criteria pair
        for i in stride(from: 1, to: args.count, by: 2) {
            let criteriaRange = try evaluate(args[i])
            let criterion = try evaluate(args[i + 1])
            
            guard case .array(let criteriaRows) = criteriaRange else {
                return .error("VALUE")
            }
            let criteriaValues = criteriaRows.flatMap { $0 }
            
            guard criteriaValues.count == maxValues.count else {
                return .error("VALUE")
            }
            
            // Filter matching indices
            matchingIndices = matchingIndices.filter { index in
                matchesCriteria(criteriaValues[index], criterion)
            }
        }
        
        // Find max of matching values
        let matchingNumbers = matchingIndices.compactMap { maxValues[$0].asDouble }
        
        guard !matchingNumbers.isEmpty else {
            return .number(0) // Excel returns 0 if no matches
        }
        
        return .number(matchingNumbers.max() ?? 0)
    }
    
    /// MINIFS - Minimum value with multiple criteria
    /// Syntax: MINIFS(min_range, criteria_range1, criterion1, [criteria_range2, criterion2], ...)
    private func evaluateMINIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count % 2 == 1 else {
            return .error("VALUE")
        }
        
        let minRange = try evaluate(args[0])
        guard case .array(let minRows) = minRange else {
            return .error("VALUE")
        }
        let minValues = minRows.flatMap { $0 }
        
        // Build list of matching indices
        var matchingIndices = Set(0..<minValues.count)
        
        // Process each criteria pair
        for i in stride(from: 1, to: args.count, by: 2) {
            let criteriaRange = try evaluate(args[i])
            let criterion = try evaluate(args[i + 1])
            
            guard case .array(let criteriaRows) = criteriaRange else {
                return .error("VALUE")
            }
            let criteriaValues = criteriaRows.flatMap { $0 }
            
            guard criteriaValues.count == minValues.count else {
                return .error("VALUE")
            }
            
            // Filter matching indices
            matchingIndices = matchingIndices.filter { index in
                matchesCriteria(criteriaValues[index], criterion)
            }
        }
        
        // Find min of matching values
        let matchingNumbers = matchingIndices.compactMap { minValues[$0].asDouble }
        
        guard !matchingNumbers.isEmpty else {
            return .number(0) // Excel returns 0 if no matches
        }
        
        return .number(matchingNumbers.min() ?? 0)
    }
    
    /// AVERAGEIFS - Average with multiple criteria
    /// Syntax: AVERAGEIFS(average_range, criteria_range1, criterion1, [criteria_range2, criterion2], ...)
    private func evaluateAVERAGEIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count % 2 == 1 else {
            return .error("VALUE")
        }
        
        let avgRange = try evaluate(args[0])
        guard case .array(let avgRows) = avgRange else {
            return .error("VALUE")
        }
        let avgValues = avgRows.flatMap { $0 }
        
        // Build list of matching indices
        var matchingIndices = Set(0..<avgValues.count)
        
        // Process each criteria pair
        for i in stride(from: 1, to: args.count, by: 2) {
            let criteriaRange = try evaluate(args[i])
            let criterion = try evaluate(args[i + 1])
            
            guard case .array(let criteriaRows) = criteriaRange else {
                return .error("VALUE")
            }
            let criteriaValues = criteriaRows.flatMap { $0 }
            
            guard criteriaValues.count == avgValues.count else {
                return .error("VALUE")
            }
            
            // Filter matching indices
            matchingIndices = matchingIndices.filter { index in
                matchesCriteria(criteriaValues[index], criterion)
            }
        }
        
        // Calculate average of matching values
        let matchingNumbers = matchingIndices.compactMap { avgValues[$0].asDouble }
        
        guard !matchingNumbers.isEmpty else {
            return .error("DIV/0")
        }
        
        let sum = matchingNumbers.reduce(0, +)
        return .number(sum / Double(matchingNumbers.count))
    }
    
    // MARK: - Statistical Functions
    
    /// STDEV - Standard deviation
    /// Syntax: STDEV(number1, [number2], ...) or STDEV.S/STDEV.P
    private func evaluateSTDEV(_ args: [FormulaExpression], sample: Bool) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard numbers.count > (sample ? 1 : 0) else {
            return .error("DIV/0")
        }
        
        let mean = numbers.reduce(0, +) / Double(numbers.count)
        let sumSquaredDiffs = numbers.map { pow($0 - mean, 2) }.reduce(0, +)
        let divisor = sample ? Double(numbers.count - 1) : Double(numbers.count)
        
        return .number(sqrt(sumSquaredDiffs / divisor))
    }
    
    /// VAR - Variance
    /// Syntax: VAR(number1, [number2], ...) or VAR.S/VAR.P
    private func evaluateVAR(_ args: [FormulaExpression], sample: Bool) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard numbers.count > (sample ? 1 : 0) else {
            return .error("DIV/0")
        }
        
        let mean = numbers.reduce(0, +) / Double(numbers.count)
        let sumSquaredDiffs = numbers.map { pow($0 - mean, 2) }.reduce(0, +)
        let divisor = sample ? Double(numbers.count - 1) : Double(numbers.count)
        
        return .number(sumSquaredDiffs / divisor)
    }
    
    /// PERCENTILE - K-th percentile
    /// Syntax: PERCENTILE(array, k) where k is 0-1
    private func evaluatePERCENTILE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "PERCENTILE", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let kVal = try evaluate(args[1])
        
        guard let k = kVal.asDouble, k >= 0, k <= 1 else {
            return .error("NUM")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        numbers.sort()
        let index = k * Double(numbers.count - 1)
        let lower = Int(floor(index))
        let upper = Int(ceil(index))
        
        if lower == upper {
            return .number(numbers[lower])
        }
        
        let fraction = index - Double(lower)
        let result = numbers[lower] + fraction * (numbers[upper] - numbers[lower])
        return .number(result)
    }
    
    /// QUARTILE - Quartile value (0=min, 1=Q1, 2=median, 3=Q3, 4=max)
    /// Syntax: QUARTILE(array, quart)
    private func evaluateQUARTILE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "QUARTILE", expected: 2, got: args.count)
        }
        
        let quartVal = try evaluate(args[1])
        guard let quart = quartVal.asDouble, quart >= 0, quart <= 4 else {
            return .error("NUM")
        }
        
        let k = quart / 4.0
        return try evaluatePERCENTILE([args[0], .number(k)])
    }
    
    /// MODE - Most frequently occurring value
    /// Syntax: MODE(number1, [number2], ...)
    private func evaluateMODE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("N/A")
        }
        
        // Count frequencies
        var frequencies: [Double: Int] = [:]
        for num in numbers {
            frequencies[num, default: 0] += 1
        }
        
        // Find most frequent
        guard let (mode, count) = frequencies.max(by: { $0.value < $1.value }), count > 1 else {
            return .error("N/A") // No mode if all values appear only once
        }
        
        return .number(mode)
    }
    
    /// LARGE - K-th largest value
    /// Syntax: LARGE(array, k)
    private func evaluateLARGE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "LARGE", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let kVal = try evaluate(args[1])
        
        guard let k = kVal.asDouble, k >= 1 else {
            return .error("NUM")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty, Int(k) <= numbers.count else {
            return .error("NUM")
        }
        
        numbers.sort(by: >)
        return .number(numbers[Int(k) - 1])
    }
    
    /// SMALL - K-th smallest value
    /// Syntax: SMALL(array, k)
    private func evaluateSMALL(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "SMALL", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let kVal = try evaluate(args[1])
        
        guard let k = kVal.asDouble, k >= 1 else {
            return .error("NUM")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty, Int(k) <= numbers.count else {
            return .error("NUM")
        }
        
        numbers.sort()
        return .number(numbers[Int(k) - 1])
    }
    
    /// RANK - Rank of a number in a list
    /// Syntax: RANK(number, ref, [order]) where order: 0=descending, 1=ascending
    private func evaluateRANK(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "RANK", expected: 2, got: args.count)
        }
        
        let numberVal = try evaluate(args[0])
        guard let number = numberVal.asDouble else {
            return .error("VALUE")
        }
        
        let refVal = try evaluate(args[1])
        var numbers = flattenToNumbers(refVal)
        
        guard !numbers.isEmpty else {
            return .error("N/A")
        }
        
        let ascending = args.count == 3 ? (try evaluate(args[2]).asDouble ?? 0) != 0 : false
        
        numbers.sort(by: ascending ? (<) : (>))
        
        guard let rank = numbers.firstIndex(of: number) else {
            return .error("N/A")
        }
        
        return .number(Double(rank + 1))
    }
    
    /// CORREL - Correlation coefficient
    /// Syntax: CORREL(array1, array2)
    private func evaluateCORREL(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "CORREL", expected: 2, got: args.count)
        }
        
        let array1Val = try evaluate(args[0])
        let array2Val = try evaluate(args[1])
        
        let array1 = flattenToNumbers(array1Val)
        let array2 = flattenToNumbers(array2Val)
        
        guard array1.count == array2.count, !array1.isEmpty else {
            return .error("N/A")
        }
        
        let n = Double(array1.count)
        let mean1 = array1.reduce(0, +) / n
        let mean2 = array2.reduce(0, +) / n
        
        var sumProduct: Double = 0
        var sumSq1: Double = 0
        var sumSq2: Double = 0
        
        for i in 0..<array1.count {
            let diff1 = array1[i] - mean1
            let diff2 = array2[i] - mean2
            sumProduct += diff1 * diff2
            sumSq1 += diff1 * diff1
            sumSq2 += diff2 * diff2
        }
        
        guard sumSq1 > 0, sumSq2 > 0 else {
            return .error("DIV/0")
        }
        
        return .number(sumProduct / sqrt(sumSq1 * sumSq2))
    }
    
    // MARK: - Math and Trigonometric Functions
    
    /// SIN - Sine of an angle (in radians)
    private func evaluateSIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SIN", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        return .number(sin(num))
    }
    
    /// COS - Cosine of an angle (in radians)
    private func evaluateCOS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "COS", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        return .number(cos(num))
    }
    
    /// TAN - Tangent of an angle (in radians)
    private func evaluateTAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "TAN", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        return .number(tan(num))
    }
    
    /// ASIN - Arcsine (in radians)
    private func evaluateASIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ASIN", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        guard num >= -1 && num <= 1 else { return .error("NUM") }
        return .number(asin(num))
    }
    
    /// ACOS - Arccosine (in radians)
    private func evaluateACOS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ACOS", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        guard num >= -1 && num <= 1 else { return .error("NUM") }
        return .number(acos(num))
    }
    
    /// ATAN - Arctangent (in radians)
    private func evaluateATAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ATAN", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        return .number(atan(num))
    }
    
    /// ATAN2 - Arctangent of x and y coordinates
    private func evaluateATAN2(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "ATAN2", expected: 2, got: args.count)
        }
        let xVal = try evaluate(args[0])
        let yVal = try evaluate(args[1])
        guard let x = xVal.asDouble, let y = yVal.asDouble else { return .error("VALUE") }
        guard x != 0 || y != 0 else { return .error("DIV/0") }
        return .number(atan2(y, x))
    }
    
    /// PI - Returns π
    private func evaluatePI(_ args: [FormulaExpression]) -> FormulaValue {
        return .number(Double.pi)
    }
    
    /// RADIANS - Convert degrees to radians
    private func evaluateRADIANS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "RADIANS", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let degrees = val.asDouble else { return .error("VALUE") }
        return .number(degrees * Double.pi / 180)
    }
    
    /// DEGREES - Convert radians to degrees
    private func evaluateDEGREES(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "DEGREES", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let radians = val.asDouble else { return .error("VALUE") }
        return .number(radians * 180 / Double.pi)
    }
    
    /// LOG - Logarithm to specified base (default 10)
    private func evaluateLOG(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "LOG", expected: 1, got: args.count)
        }
        let numVal = try evaluate(args[0])
        guard let num = numVal.asDouble, num > 0 else { return .error("NUM") }
        
        let base = args.count == 2 ? (try evaluate(args[1]).asDouble ?? 10) : 10
        guard base > 0 && base != 1 else { return .error("NUM") }
        
        return .number(log(num) / log(base))
    }
    
    /// LOG10 - Base-10 logarithm
    private func evaluateLOG10(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "LOG10", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble, num > 0 else { return .error("NUM") }
        return .number(log10(num))
    }
    
    /// LN - Natural logarithm (base e)
    private func evaluateLN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "LN", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble, num > 0 else { return .error("NUM") }
        return .number(log(num))
    }
    
    /// EXP - e raised to a power
    private func evaluateEXP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "EXP", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        return .number(exp(num))
    }
    
    /// CEILING - Round up to nearest multiple
    private func evaluateCEILING(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "CEILING", expected: 1, got: args.count)
        }
        let numVal = try evaluate(args[0])
        guard let num = numVal.asDouble else { return .error("VALUE") }
        
        let significance = args.count == 2 ? (try evaluate(args[1]).asDouble ?? 1) : 1
        guard significance != 0 else { return .error("DIV/0") }
        
        return .number(ceil(num / significance) * significance)
    }
    
    /// FLOOR - Round down to nearest multiple
    private func evaluateFLOOR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "FLOOR", expected: 1, got: args.count)
        }
        let numVal = try evaluate(args[0])
        guard let num = numVal.asDouble else { return .error("VALUE") }
        
        let significance = args.count == 2 ? (try evaluate(args[1]).asDouble ?? 1) : 1
        guard significance != 0 else { return .error("DIV/0") }
        
        return .number(floor(num / significance) * significance)
    }
    
    /// TRUNC - Truncate to integer
    private func evaluateTRUNC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "TRUNC", expected: 1, got: args.count)
        }
        let numVal = try evaluate(args[0])
        guard let num = numVal.asDouble else { return .error("VALUE") }
        
        let digits = args.count == 2 ? Int(try evaluate(args[1]).asDouble ?? 0) : 0
        let multiplier = pow(10.0, Double(digits))
        
        return .number(trunc(num * multiplier) / multiplier)
    }
    
    /// SIGN - Sign of a number (-1, 0, or 1)
    private func evaluateSIGN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SIGN", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble else { return .error("VALUE") }
        
        if num > 0 { return .number(1) }
        if num < 0 { return .number(-1) }
        return .number(0)
    }
    
    /// FACT - Factorial
    private func evaluateFACT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "FACT", expected: 1, got: args.count)
        }
        let val = try evaluate(args[0])
        guard let num = val.asDouble, num >= 0 else { return .error("NUM") }
        
        let n = Int(num)
        guard n <= 170 else { return .error("NUM") } // Overflow protection
        
        if n <= 1 {
            return .number(1)  // 0! = 1, 1! = 1
        }
        
        var result: Double = 1
        for i in 2...n {
            result *= Double(i)
        }
        return .number(result)
    }
    
    /// SUMPRODUCT - Sum of products of corresponding ranges
    private func evaluateSUMPRODUCT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        // Evaluate all arrays
        var arrays: [[Double]] = []
        for arg in args {
            let val = try evaluate(arg)
            guard case .array(let rows) = val else {
                return .error("VALUE")
            }
            let numbers = rows.flatMap { $0 }.compactMap { $0.asDouble }
            arrays.append(numbers)
        }
        
        // Check all arrays have same length
        guard let firstCount = arrays.first?.count else {
            return .error("VALUE")
        }
        guard arrays.allSatisfy({ $0.count == firstCount }) else {
            return .error("VALUE")
        }
        
        // Calculate sum of products
        var sum: Double = 0
        for i in 0..<firstCount {
            var product: Double = 1
            for array in arrays {
                product *= array[i]
            }
            sum += product
        }
        
        return .number(sum)
    }
    
    /// GCD - Greatest common divisor
    private func evaluateGCD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Int] = []
        for arg in args {
            let val = try evaluate(arg)
            for num in flattenToNumbers(val) {
                guard num >= 0 else { return .error("NUM") }
                numbers.append(Int(num))
            }
        }
        
        guard !numbers.isEmpty else {
            return .error("VALUE")
        }
        
        func gcd(_ a: Int, _ b: Int) -> Int {
            var a = a, b = b
            while b != 0 {
                (a, b) = (b, a % b)
            }
            return abs(a)
        }
        
        var result = numbers[0]
        for num in numbers.dropFirst() {
            result = gcd(result, num)
        }
        
        return .number(Double(result))
    }
    
    /// LCM - Least common multiple
    private func evaluateLCM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Int] = []
        for arg in args {
            let val = try evaluate(arg)
            for num in flattenToNumbers(val) {
                guard num >= 0 else { return .error("NUM") }
                numbers.append(Int(num))
            }
        }
        
        guard !numbers.isEmpty else {
            return .error("VALUE")
        }
        
        func gcd(_ a: Int, _ b: Int) -> Int {
            var a = a, b = b
            while b != 0 {
                (a, b) = (b, a % b)
            }
            return abs(a)
        }
        
        func lcm(_ a: Int, _ b: Int) -> Int {
            return abs(a * b) / gcd(a, b)
        }
        
        var result = numbers[0]
        for num in numbers.dropFirst() {
            result = lcm(result, num)
        }
        
        return .number(Double(result))
    }
    
    // MARK: - Financial Functions
    
    /// PMT - Payment for a loan based on constant payments and interest rate
    private func evaluatePMT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "PMT", expected: 3, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        let nperVal = try evaluate(args[1])
        let pvVal = try evaluate(args[2])
        let fvVal = args.count > 3 ? try evaluate(args[3]) : .number(0)
        let typeVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        
        guard let rate = rateVal.asDouble,
              let nper = nperVal.asDouble,
              let pv = pvVal.asDouble,
              let fv = fvVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        if rate == 0 {
            return .number(-(pv + fv) / nper)
        }
        
        let pvif = pow(1 + rate, nper)
        let pmt = -rate * (pv * pvif + fv) / (pvif - 1)
        
        if type == 1 {
            return .number(pmt / (1 + rate))
        } else {
            return .number(pmt)
        }
    }
    
    /// PV - Present value of an investment
    private func evaluatePV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "PV", expected: 3, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        let nperVal = try evaluate(args[1])
        let pmtVal = try evaluate(args[2])
        let fvVal = args.count > 3 ? try evaluate(args[3]) : .number(0)
        let typeVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        
        guard let rate = rateVal.asDouble,
              let nper = nperVal.asDouble,
              let pmt = pmtVal.asDouble,
              let fv = fvVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        if rate == 0 {
            return .number(-pmt * nper - fv)
        }
        
        let pvif = pow(1 + rate, nper)
        let pv: Double
        
        if type == 1 {
            pv = (-pmt * (1 + rate) * (pvif - 1) / rate - fv) / pvif
        } else {
            pv = (-pmt * (pvif - 1) / rate - fv) / pvif
        }
        
        return .number(pv)
    }
    
    /// FV - Future value of an investment
    private func evaluateFV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "FV", expected: 3, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        let nperVal = try evaluate(args[1])
        let pmtVal = try evaluate(args[2])
        let pvVal = args.count > 3 ? try evaluate(args[3]) : .number(0)
        let typeVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        
        guard let rate = rateVal.asDouble,
              let nper = nperVal.asDouble,
              let pmt = pmtVal.asDouble,
              let pv = pvVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        if rate == 0 {
            return .number(-pv - pmt * nper)
        }
        
        let pvif = pow(1 + rate, nper)
        let fv: Double
        
        if type == 1 {
            fv = -pv * pvif - pmt * (1 + rate) * (pvif - 1) / rate
        } else {
            fv = -pv * pvif - pmt * (pvif - 1) / rate
        }
        
        return .number(fv)
    }
    
    /// RATE - Interest rate per period
    private func evaluateRATE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 6 else {
            throw FormulaError.invalidArgumentCount(function: "RATE", expected: 3, got: args.count)
        }
        
        let nperVal = try evaluate(args[0])
        let pmtVal = try evaluate(args[1])
        let pvVal = try evaluate(args[2])
        let fvVal = args.count > 3 ? try evaluate(args[3]) : .number(0)
        let typeVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        let guessVal = args.count > 5 ? try evaluate(args[5]) : .number(0.1)
        
        guard let nper = nperVal.asDouble,
              let pmt = pmtVal.asDouble,
              let pv = pvVal.asDouble,
              let fv = fvVal.asDouble,
              let type = typeVal.asDouble,
              var rate = guessVal.asDouble else {
            return .error("VALUE")
        }
        
        // Newton-Raphson method to solve for rate
        let maxIterations = 100
        let tolerance = 1e-7
        
        for _ in 0..<maxIterations {
            let pvif = pow(1 + rate, nper)
            let f: Double
            let fDerivative: Double
            
            if type == 1 {
                f = pv * pvif + pmt * (1 + rate) * (pvif - 1) / rate + fv
                fDerivative = nper * pv * pow(1 + rate, nper - 1) + 
                             pmt * ((1 + rate) * nper * pow(1 + rate, nper - 1) / rate + 
                             (1 + rate) * (pvif - 1) / rate - (1 + rate) * (pvif - 1) / (rate * rate) +
                             (pvif - 1) / rate)
            } else {
                f = pv * pvif + pmt * (pvif - 1) / rate + fv
                fDerivative = nper * pv * pow(1 + rate, nper - 1) +
                             pmt * (nper * pow(1 + rate, nper - 1) / rate - (pvif - 1) / (rate * rate))
            }
            
            let newRate = rate - f / fDerivative
            
            if abs(newRate - rate) < tolerance {
                return .number(newRate)
            }
            
            rate = newRate
        }
        
        return .error("NUM") // Did not converge
    }
    
    /// NPER - Number of periods for an investment
    private func evaluateNPER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "NPER", expected: 3, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        let pmtVal = try evaluate(args[1])
        let pvVal = try evaluate(args[2])
        let fvVal = args.count > 3 ? try evaluate(args[3]) : .number(0)
        let typeVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        
        guard let rate = rateVal.asDouble,
              let pmt = pmtVal.asDouble,
              let pv = pvVal.asDouble,
              let fv = fvVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        if rate == 0 {
            return .number(-(pv + fv) / pmt)
        }
        
        let adjustedPmt = type == 1 ? pmt * (1 + rate) : pmt
        let nper = log((adjustedPmt - fv * rate) / (adjustedPmt + pv * rate)) / log(1 + rate)
        
        return .number(nper)
    }
    
    /// IPMT - Interest payment for a given period
    private func evaluateIPMT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 4 && args.count <= 6 else {
            throw FormulaError.invalidArgumentCount(function: "IPMT", expected: 4, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        let perVal = try evaluate(args[1])
        let nperVal = try evaluate(args[2])
        let pvVal = try evaluate(args[3])
        let fvVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        let typeVal = args.count > 5 ? try evaluate(args[5]) : .number(0)
        
        guard let rate = rateVal.asDouble,
              let per = perVal.asDouble,
              let nper = nperVal.asDouble,
              let pv = pvVal.asDouble,
              let fv = fvVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard per >= 1 && per <= nper else {
            return .error("NUM")
        }
        
        // Calculate PMT first
        let pvif = pow(1 + rate, nper)
        let pmt = -rate * (pv * pvif + fv) / (pvif - 1)
        let adjustedPmt = type == 1 ? pmt / (1 + rate) : pmt
        
        // Calculate remaining balance at period per-1
        let remainingPV: Double
        if per == 1 && type == 1 {
            return .number(0)
        } else if per == 1 {
            remainingPV = pv
        } else {
            let perPaid = type == 1 ? per - 2 : per - 1
            remainingPV = pv * pow(1 + rate, perPaid) + adjustedPmt * (pow(1 + rate, perPaid) - 1) / rate
        }
        
        let ipmt = -remainingPV * rate
        
        return .number(ipmt)
    }
    
    /// PPMT - Principal payment for a given period
    private func evaluatePPMT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 4 && args.count <= 6 else {
            throw FormulaError.invalidArgumentCount(function: "PPMT", expected: 4, got: args.count)
        }
        
        // PPMT = PMT - IPMT
        let pmtResult = try evaluatePMT([args[0], args[2], args[3]] + 
                                       (args.count > 4 ? [args[4]] : []) + 
                                       (args.count > 5 ? [args[5]] : []))
        let ipmtResult = try evaluateIPMT(args)
        
        guard let pmt = pmtResult.asDouble,
              let ipmt = ipmtResult.asDouble else {
            return .error("VALUE")
        }
        
        return .number(pmt - ipmt)
    }
    
    /// NPV - Net present value of cash flows
    private func evaluateNPV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            throw FormulaError.invalidArgumentCount(function: "NPV", expected: 2, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        guard let rate = rateVal.asDouble else {
            return .error("VALUE")
        }
        
        var npv: Double = 0
        var period: Double = 1
        
        for i in 1..<args.count {
            let val = try evaluate(args[i])
            for cashFlow in flattenToNumbers(val) {
                npv += cashFlow / pow(1 + rate, period)
                period += 1
            }
        }
        
        return .number(npv)
    }
    
    /// IRR - Internal rate of return
    private func evaluateIRR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "IRR", expected: 1, got: args.count)
        }
        
        let valuesExpr = try evaluate(args[0])
        let guessVal = args.count > 1 ? try evaluate(args[1]) : .number(0.1)
        
        guard var rate = guessVal.asDouble else {
            return .error("VALUE")
        }
        
        let cashFlows = flattenToNumbers(valuesExpr)
        guard cashFlows.count >= 2 else {
            return .error("VALUE")
        }
        
        // Newton-Raphson method
        let maxIterations = 100
        let tolerance = 1e-7
        
        for _ in 0..<maxIterations {
            var npv: Double = 0
            var npvDerivative: Double = 0
            
            for (i, cashFlow) in cashFlows.enumerated() {
                let period = Double(i)
                npv += cashFlow / pow(1 + rate, period)
                npvDerivative += -period * cashFlow / pow(1 + rate, period + 1)
            }
            
            if abs(npv) < tolerance {
                return .number(rate)
            }
            
            rate = rate - npv / npvDerivative
        }
        
        return .error("NUM") // Did not converge
    }
    
    /// XNPV - Net present value for non-periodic cash flows
    private func evaluateXNPV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "XNPV", expected: 3, got: args.count)
        }
        
        let rateVal = try evaluate(args[0])
        let valuesExpr = try evaluate(args[1])
        let datesExpr = try evaluate(args[2])
        
        guard let rate = rateVal.asDouble else {
            return .error("VALUE")
        }
        
        let cashFlows = flattenToNumbers(valuesExpr)
        let dates = flattenToNumbers(datesExpr)
        
        guard cashFlows.count == dates.count && cashFlows.count >= 2 else {
            return .error("VALUE")
        }
        
        let baseDate = dates[0]
        var xnpv: Double = 0
        
        for (i, cashFlow) in cashFlows.enumerated() {
            let daysDiff = dates[i] - baseDate
            let years = daysDiff / 365.0
            xnpv += cashFlow / pow(1 + rate, years)
        }
        
        return .number(xnpv)
    }
    
    /// XIRR - Internal rate of return for non-periodic cash flows
    private func evaluateXIRR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "XIRR", expected: 2, got: args.count)
        }
        
        let valuesExpr = try evaluate(args[0])
        let datesExpr = try evaluate(args[1])
        let guessVal = args.count > 2 ? try evaluate(args[2]) : .number(0.1)
        
        guard var rate = guessVal.asDouble else {
            return .error("VALUE")
        }
        
        let cashFlows = flattenToNumbers(valuesExpr)
        let dates = flattenToNumbers(datesExpr)
        
        guard cashFlows.count == dates.count && cashFlows.count >= 2 else {
            return .error("VALUE")
        }
        
        let baseDate = dates[0]
        
        // Newton-Raphson method
        let maxIterations = 100
        let tolerance = 1e-7
        
        for _ in 0..<maxIterations {
            var xnpv: Double = 0
            var xnpvDerivative: Double = 0
            
            for (i, cashFlow) in cashFlows.enumerated() {
                let daysDiff = dates[i] - baseDate
                let years = daysDiff / 365.0
                let discount = pow(1 + rate, years)
                
                xnpv += cashFlow / discount
                xnpvDerivative += -years * cashFlow / (discount * (1 + rate))
            }
            
            if abs(xnpv) < tolerance {
                return .number(rate)
            }
            
            rate = rate - xnpv / xnpvDerivative
        }
        
        return .error("NUM") // Did not converge
    }
    
    // MARK: - Date/Time Functions (Extended)
    
    /// WEEKDAY - Returns the day of week (1-7, Sunday=1 by default)
    private func evaluateWEEKDAY(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "WEEKDAY", expected: 1, got: args.count)
        }
        
        let serialVal = try evaluate(args[0])
        let returnTypeVal = args.count > 1 ? try evaluate(args[1]) : .number(1)
        
        guard let serial = serialVal.asDouble,
              let returnType = returnTypeVal.asDouble else {
            return .error("VALUE")
        }
        
        // Excel serial date: Jan 1, 1900 is day 1 (which was a Monday in Excel's calendar)
        // But Excel incorrectly treats 1900 as a leap year, so we need to account for this
        let dayOfWeek = Int(serial.truncatingRemainder(dividingBy: 7))
        
        switch Int(returnType) {
        case 1: // 1 (Sunday) to 7 (Saturday)
            return .number(Double((dayOfWeek + 6) % 7 + 1))
        case 2: // 1 (Monday) to 7 (Sunday)
            return .number(Double((dayOfWeek + 5) % 7 + 1))
        case 3: // 0 (Monday) to 6 (Sunday)
            return .number(Double((dayOfWeek + 5) % 7))
        default:
            return .error("NUM")
        }
    }
    
    /// WEEKNUM - Returns the week number of the year
    private func evaluateWEEKNUM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "WEEKNUM", expected: 1, got: args.count)
        }
        
        let serialVal = try evaluate(args[0])
        let returnTypeVal = args.count > 1 ? try evaluate(args[1]) : .number(1)
        
        guard let serial = serialVal.asDouble,
              let _ = returnTypeVal.asDouble else {
            return .error("VALUE")
        }
        
        // Simplified: week number based on day of year
        let dayOfYear = Int(serial - 1) % 365 + 1
        let weekNum = (dayOfYear - 1) / 7 + 1
        
        return .number(Double(weekNum))
    }
    
    /// ISOWEEKNUM - Returns the ISO week number of the year
    private func evaluateISOWEEKNUM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISOWEEKNUM", expected: 1, got: args.count)
        }
        
        let serialVal = try evaluate(args[0])
        guard let serial = serialVal.asDouble else {
            return .error("VALUE")
        }
        
        // ISO week starts on Monday, week 1 contains Jan 4
        let dayOfYear = Int(serial - 1) % 365 + 1
        let weekNum = (dayOfYear + 3) / 7
        
        return .number(Double(max(1, weekNum)))
    }
    
    /// EOMONTH - Returns the last day of the month n months from start date
    private func evaluateEOMONTH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "EOMONTH", expected: 2, got: args.count)
        }
        
        let startDateVal = try evaluate(args[0])
        let monthsVal = try evaluate(args[1])
        
        guard let startDate = startDateVal.asDouble,
              let months = monthsVal.asDouble else {
            return .error("VALUE")
        }
        
        // Extract year, month, day from Excel serial
        let daysFrom1900 = Int(startDate)
        let year = 1900 + (daysFrom1900 - 1) / 365
        let dayOfYear = (daysFrom1900 - 1) % 365 + 1
        var month = (dayOfYear - 1) / 30 + 1
        
        // Add months
        month += Int(months)
        var adjustedYear = year
        while month > 12 {
            month -= 12
            adjustedYear += 1
        }
        while month < 1 {
            month += 12
            adjustedYear -= 1
        }
        
        // Get last day of target month
        let daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        var lastDay = daysInMonth[month - 1]
        
        // Adjust for leap year (including Excel's 1900 bug)
        if month == 2 && (adjustedYear % 4 == 0) {
            lastDay = 29
        }
        
        // Calculate new serial date (approximate)
        let newSerial = startDate + Double(Int(months) * 30) + Double(lastDay - 15)
        
        return .number(newSerial)
    }
    
    /// EDATE - Returns date n months from start date
    private func evaluateEDATE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "EDATE", expected: 2, got: args.count)
        }
        
        let startDateVal = try evaluate(args[0])
        let monthsVal = try evaluate(args[1])
        
        guard let startDate = startDateVal.asDouble,
              let months = monthsVal.asDouble else {
            return .error("VALUE")
        }
        
        // Approximate: 30 days per month
        return .number(startDate + months * 30)
    }
    
    /// NETWORKDAYS - Number of working days between two dates
    private func evaluateNETWORKDAYS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "NETWORKDAYS", expected: 2, got: args.count)
        }
        
        let startDateVal = try evaluate(args[0])
        let endDateVal = try evaluate(args[1])
        
        guard let startDate = startDateVal.asDouble,
              let endDate = endDateVal.asDouble else {
            return .error("VALUE")
        }
        
        let totalDays = Int(abs(endDate - startDate)) + 1
        let weeks = totalDays / 7
        let remainingDays = totalDays % 7
        
        // 5 working days per week
        let workDays = weeks * 5 + min(remainingDays, 5)
        
        return .number(Double(workDays))
    }
    
    /// WORKDAY - Returns date n working days from start
    private func evaluateWORKDAY(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "WORKDAY", expected: 2, got: args.count)
        }
        
        let startDateVal = try evaluate(args[0])
        let daysVal = try evaluate(args[1])
        
        guard let startDate = startDateVal.asDouble,
              let days = daysVal.asDouble else {
            return .error("VALUE")
        }
        
        // Approximate: 1.4 calendar days per work day (5/7 ratio)
        let calendarDays = days * 1.4
        
        return .number(startDate + calendarDays)
    }
    
    /// DATEDIF - Difference between two dates
    private func evaluateDATEDIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "DATEDIF", expected: 3, got: args.count)
        }
        
        let startDateVal = try evaluate(args[0])
        let endDateVal = try evaluate(args[1])
        let unitVal = try evaluate(args[2])
        
        guard let startDate = startDateVal.asDouble,
              let endDate = endDateVal.asDouble,
              case .string(let unit) = unitVal else {
            return .error("VALUE")
        }
        
        let days = endDate - startDate
        
        switch unit.uppercased() {
        case "D": // Days
            return .number(days)
        case "M": // Months (approximate)
            return .number(days / 30)
        case "Y": // Years (approximate)
            return .number(days / 365)
        case "MD": // Days ignoring months/years
            return .number(days.truncatingRemainder(dividingBy: 30))
        case "YM": // Months ignoring years
            return .number((days / 30).truncatingRemainder(dividingBy: 12))
        case "YD": // Days ignoring years
            return .number(days.truncatingRemainder(dividingBy: 365))
        default:
            return .error("NUM")
        }
    }
    
    /// YEARFRAC - Fraction of year between two dates
    private func evaluateYEARFRAC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "YEARFRAC", expected: 2, got: args.count)
        }
        
        let startDateVal = try evaluate(args[0])
        let endDateVal = try evaluate(args[1])
        let basisVal = args.count > 2 ? try evaluate(args[2]) : .number(0)
        
        guard let startDate = startDateVal.asDouble,
              let endDate = endDateVal.asDouble,
              let basis = basisVal.asDouble else {
            return .error("VALUE")
        }
        
        let days = abs(endDate - startDate)
        
        switch Int(basis) {
        case 0: // US (NASD) 30/360
            return .number(days / 360)
        case 1: // Actual/actual
            return .number(days / 365.25)
        case 2: // Actual/360
            return .number(days / 360)
        case 3: // Actual/365
            return .number(days / 365)
        case 4: // European 30/360
            return .number(days / 360)
        default:
            return .error("NUM")
        }
    }
    
    /// TIME - Returns serial number for a time
    private func evaluateTIME(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "TIME", expected: 3, got: args.count)
        }
        
        let hourVal = try evaluate(args[0])
        let minuteVal = try evaluate(args[1])
        let secondVal = try evaluate(args[2])
        
        guard let hour = hourVal.asDouble,
              let minute = minuteVal.asDouble,
              let second = secondVal.asDouble else {
            return .error("VALUE")
        }
        
        let totalSeconds = hour * 3600 + minute * 60 + second
        let fraction = totalSeconds / 86400.0 // Seconds in a day
        
        return .number(fraction)
    }
    
    /// TIMEVALUE - Converts time text to serial number
    private func evaluateTIMEVALUE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "TIMEVALUE", expected: 1, got: args.count)
        }
        
        let timeTextVal = try evaluate(args[0])
        guard case .string(let timeText) = timeTextVal else {
            return .error("VALUE")
        }
        
        // Parse simple time format "HH:MM:SS" or "HH:MM"
        let components = timeText.split(separator: ":").map { Double($0) ?? 0 }
        guard components.count >= 2 else {
            return .error("VALUE")
        }
        
        let hour = components[0]
        let minute = components[1]
        let second = components.count > 2 ? components[2] : 0
        
        let totalSeconds = hour * 3600 + minute * 60 + second
        let fraction = totalSeconds / 86400.0
        
        return .number(fraction)
    }
    
    /// HOUR - Returns the hour component (0-23)
    private func evaluateHOUR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "HOUR", expected: 1, got: args.count)
        }
        
        let serialVal = try evaluate(args[0])
        guard let serial = serialVal.asDouble else {
            return .error("VALUE")
        }
        
        let fraction = serial - floor(serial)
        let totalSeconds = fraction * 86400
        let hour = Int(totalSeconds) / 3600
        
        return .number(Double(hour))
    }
    
    /// MINUTE - Returns the minute component (0-59)
    private func evaluateMINUTE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "MINUTE", expected: 1, got: args.count)
        }
        
        let serialVal = try evaluate(args[0])
        guard let serial = serialVal.asDouble else {
            return .error("VALUE")
        }
        
        let fraction = serial - floor(serial)
        let totalSeconds = fraction * 86400
        let minute = (Int(totalSeconds) % 3600) / 60
        
        return .number(Double(minute))
    }
    
    /// SECOND - Returns the second component (0-59)
    private func evaluateSECOND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SECOND", expected: 1, got: args.count)
        }
        
        let serialVal = try evaluate(args[0])
        guard let serial = serialVal.asDouble else {
            return .error("VALUE")
        }
        
        let fraction = serial - floor(serial)
        let totalSeconds = fraction * 86400
        let second = Int(totalSeconds) % 60
        
        return .number(Double(second))
    }
    
    // MARK: - Text Functions (Extended)
    
    /// PROPER - Capitalizes first letter of each word
    private func evaluatePROPER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "PROPER", expected: 1, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        guard case .string(let text) = textVal else {
            return .error("VALUE")
        }
        
        let result = text.capitalized
        return .string(result)
    }
    
    /// CLEAN - Removes non-printable characters
    private func evaluateCLEAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "CLEAN", expected: 1, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        guard case .string(let text) = textVal else {
            return .error("VALUE")
        }
        
        // Remove characters with ASCII values 0-31
        let cleaned = text.filter { $0.unicodeScalars.first?.value ?? 0 >= 32 }
        return .string(String(cleaned))
    }
    
    /// CHAR - Returns character for ASCII code
    private func evaluateCHAR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "CHAR", expected: 1, got: args.count)
        }
        
        let numVal = try evaluate(args[0])
        guard let num = numVal.asDouble else {
            return .error("VALUE")
        }
        
        let code = Int(num)
        guard code > 0 && code <= 255 else {
            return .error("VALUE")
        }
        
        if let scalar = UnicodeScalar(code) {
            return .string(String(Character(scalar)))
        }
        
        return .error("VALUE")
    }
    
    /// CODE - Returns ASCII code for first character
    private func evaluateCODE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "CODE", expected: 1, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        guard case .string(let text) = textVal, !text.isEmpty else {
            return .error("VALUE")
        }
        
        if let firstChar = text.first, let scalar = firstChar.unicodeScalars.first {
            return .number(Double(scalar.value))
        }
        
        return .error("VALUE")
    }
    
    /// EXACT - Case-sensitive comparison
    private func evaluateEXACT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "EXACT", expected: 2, got: args.count)
        }
        
        let text1Val = try evaluate(args[0])
        let text2Val = try evaluate(args[1])
        
        let text1: String
        let text2: String
        
        switch (text1Val, text2Val) {
        case (.string(let t1), .string(let t2)):
            text1 = t1
            text2 = t2
        default:
            return .error("VALUE")
        }
        
        return .number(text1 == text2 ? 1 : 0)
    }
    
    /// REPLACE - Replaces part of text string
    private func evaluateREPLACE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 4 else {
            throw FormulaError.invalidArgumentCount(function: "REPLACE", expected: 4, got: args.count)
        }
        
        let oldTextVal = try evaluate(args[0])
        let startNumVal = try evaluate(args[1])
        let numCharsVal = try evaluate(args[2])
        let newTextVal = try evaluate(args[3])
        
        guard case .string(let oldText) = oldTextVal,
              let startNum = startNumVal.asDouble,
              let numChars = numCharsVal.asDouble,
              case .string(let newText) = newTextVal else {
            return .error("VALUE")
        }
        
        let start = Int(startNum) - 1  // Excel uses 1-based indexing
        let count = Int(numChars)
        
        guard start >= 0 && start <= oldText.count else {
            return .error("VALUE")
        }
        
        var result = oldText
        let startIndex = result.index(result.startIndex, offsetBy: start)
        let endIndex = result.index(startIndex, offsetBy: min(count, result.count - start))
        result.replaceSubrange(startIndex..<endIndex, with: newText)
        
        return .string(result)
    }
    
    /// REPT - Repeats text a given number of times
    private func evaluateREPT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "REPT", expected: 2, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        let numTimesVal = try evaluate(args[1])
        
        guard case .string(let text) = textVal,
              let numTimes = numTimesVal.asDouble else {
            return .error("VALUE")
        }
        
        let times = Int(numTimes)
        guard times >= 0 && times <= 32767 else {
            return .error("VALUE")
        }
        
        return .string(String(repeating: text, count: times))
    }
    
    /// VALUE - Converts text to number
    private func evaluateVALUE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "VALUE", expected: 1, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        
        switch textVal {
        case .number(let num):
            return .number(num)
        case .string(let text):
            // Remove common formatting characters
            let cleaned = text.replacingOccurrences(of: ",", with: "")
                             .replacingOccurrences(of: "$", with: "")
                             .replacingOccurrences(of: "%", with: "")
                             .trimmingCharacters(in: .whitespaces)
            
            if let num = Double(cleaned) {
                return .number(num)
            } else {
                return .error("VALUE")
            }
        default:
            return .error("VALUE")
        }
    }
    
    /// TEXTBEFORE - Returns text before delimiter (Excel 365)
    private func evaluateTEXTBEFORE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "TEXTBEFORE", expected: 2, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        let delimiterVal = try evaluate(args[1])
        
        guard case .string(let text) = textVal,
              case .string(let delimiter) = delimiterVal else {
            return .error("VALUE")
        }
        
        if let range = text.range(of: delimiter) {
            return .string(String(text[..<range.lowerBound]))
        } else {
            return .error("N/A")
        }
    }
    
    /// TEXTAFTER - Returns text after delimiter (Excel 365)
    private func evaluateTEXTAFTER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "TEXTAFTER", expected: 2, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        let delimiterVal = try evaluate(args[1])
        
        guard case .string(let text) = textVal,
              case .string(let delimiter) = delimiterVal else {
            return .error("VALUE")
        }
        
        if let range = text.range(of: delimiter) {
            return .string(String(text[range.upperBound...]))
        } else {
            return .error("N/A")
        }
    }
    
    /// TEXTSPLIT - Splits text into array (Excel 365)
    private func evaluateTEXTSPLIT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "TEXTSPLIT", expected: 2, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        let colDelimiterVal = try evaluate(args[1])
        
        guard case .string(let text) = textVal,
              case .string(let colDelimiter) = colDelimiterVal else {
            return .error("VALUE")
        }
        
        let parts = text.split(separator: colDelimiter).map { FormulaValue.string(String($0)) }
        
        // For now, return as a 1D array (single row)
        return .array([parts])
    }
    
    // MARK: - Lookup & Reference Functions (Extended)
    
    /// HLOOKUP - Horizontal lookup in a table
    private func evaluateHLOOKUP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "HLOOKUP", expected: 3, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let tableArray = try evaluate(args[1])
        let rowIndexNum = try evaluate(args[2])
        let rangeLookup = args.count > 3 ? try evaluate(args[3]) : .number(1)
        
        guard let rowIndex = rowIndexNum.asDouble,
              let exactMatch = rangeLookup.asDouble else {
            return .error("VALUE")
        }
        
        guard case .array(let rows) = tableArray, !rows.isEmpty else {
            return .error("N/A")
        }
        
        let rowIdx = Int(rowIndex) - 1
        guard rowIdx >= 0 && rowIdx < rows.count else {
            return .error("REF")
        }
        
        // Search in first row
        let searchRow = rows[0]
        
        for (colIdx, cell) in searchRow.enumerated() {
            if exactMatch == 0 {
                // Exact match
                if cell == lookupValue {
                    guard colIdx < rows[rowIdx].count else { return .error("REF") }
                    return rows[rowIdx][colIdx]
                }
            }
        }
        
        return .error("N/A")
    }
    
    /// XMATCH - Returns relative position of item in array (Excel 365)
    private func evaluateXMATCH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "XMATCH", expected: 2, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let lookupArray = try evaluate(args[1])
        let matchMode = args.count > 2 ? try evaluate(args[2]) : .number(0)
        
        guard let mode = matchMode.asDouble else {
            return .error("VALUE")
        }
        
        let values: [FormulaValue]
        if case .array(let rows) = lookupArray {
            // Check if it's a single column (each row has one element)
            if rows.first?.count == 1 {
                values = rows.map { $0[0] }
            } else if let firstRow = rows.first {
                // It's a single row
                values = firstRow
            } else {
                values = []
            }
        } else {
            values = [lookupArray]
        }
        
        for (index, value) in values.enumerated() {
            if Int(mode) == 0 { // Exact match
                if value == lookupValue {
                    return .number(Double(index + 1))
                }
            }
        }
        
        return .error("N/A")
    }
    
    /// OFFSET - Returns reference offset from starting point
    private func evaluateOFFSET(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 5 else {
            throw FormulaError.invalidArgumentCount(function: "OFFSET", expected: 3, got: args.count)
        }
        
        let referenceVal = try evaluate(args[0])
        let rowsVal = try evaluate(args[1])
        let colsVal = try evaluate(args[2])
        
        guard let _ = rowsVal.asDouble,
              let _ = colsVal.asDouble else {
            return .error("VALUE")
        }
        
        // Simplified: just return the reference value adjusted
        // In a full implementation, this would need access to sheet coordinates
        return referenceVal
    }
    
    /// INDIRECT - Returns reference specified by text string
    private func evaluateINDIRECT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "INDIRECT", expected: 1, got: args.count)
        }
        
        let refTextVal = try evaluate(args[0])
        
        guard case .string(let refText) = refTextVal else {
            return .error("VALUE")
        }
        
        // Simplified: try to evaluate as cell reference
        // In a full implementation, this would parse the reference string and look it up
        guard let cellRef = CellReference(refText) else {
            return .error("REF")
        }
        
        return try evaluate(.cellRef(cellRef))
    }
    
    /// CHOOSE - Returns value from list based on index
    private func evaluateCHOOSE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            throw FormulaError.invalidArgumentCount(function: "CHOOSE", expected: 2, got: args.count)
        }
        
        let indexVal = try evaluate(args[0])
        guard let index = indexVal.asDouble else {
            return .error("VALUE")
        }
        
        let idx = Int(index)
        guard idx >= 1 && idx < args.count else {
            return .error("VALUE")
        }
        
        return try evaluate(args[idx])
    }
    
    /// CHOOSECOLS - Returns specified columns from array (Excel 365)
    private func evaluateCHOOSECOLS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            throw FormulaError.invalidArgumentCount(function: "CHOOSECOLS", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        
        guard case .array(let rows) = arrayVal else {
            return .error("VALUE")
        }
        
        var colIndices: [Int] = []
        for i in 1..<args.count {
            let colVal = try evaluate(args[i])
            guard let colNum = colVal.asDouble else {
                return .error("VALUE")
            }
            colIndices.append(Int(colNum) - 1)
        }
        
        var result: [[FormulaValue]] = []
        for row in rows {
            var newRow: [FormulaValue] = []
            for colIdx in colIndices {
                if colIdx >= 0 && colIdx < row.count {
                    newRow.append(row[colIdx])
                } else {
                    return .error("REF")
                }
            }
            result.append(newRow)
        }
        
        return .array(result)
    }
    
    /// CHOOSEROWS - Returns specified rows from array (Excel 365)
    private func evaluateCHOOSEROWS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            throw FormulaError.invalidArgumentCount(function: "CHOOSEROWS", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        
        guard case .array(let rows) = arrayVal else {
            return .error("VALUE")
        }
        
        var result: [[FormulaValue]] = []
        for i in 1..<args.count {
            let rowVal = try evaluate(args[i])
            guard let rowNum = rowVal.asDouble else {
                return .error("VALUE")
            }
            
            let rowIdx = Int(rowNum) - 1
            if rowIdx >= 0 && rowIdx < rows.count {
                result.append(rows[rowIdx])
            } else {
                return .error("REF")
            }
        }
        
        return .array(result)
    }
    
    /// TRANSPOSE - Transposes array
    private func evaluateTRANSPOSE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "TRANSPOSE", expected: 1, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        
        guard case .array(let rows) = arrayVal, !rows.isEmpty else {
            return .error("VALUE")
        }
        
        let numRows = rows.count
        let numCols = rows[0].count
        
        var transposed: [[FormulaValue]] = Array(repeating: Array(repeating: .number(0), count: numRows), count: numCols)
        
        for i in 0..<numRows {
            for j in 0..<min(numCols, rows[i].count) {
                transposed[j][i] = rows[i][j]
            }
        }
        
        return .array(transposed)
    }
    
    /// ROWS - Returns number of rows in reference
    private func evaluateROWS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ROWS", expected: 1, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        
        if case .array(let rows) = arrayVal {
            return .number(Double(rows.count))
        }
        
        // Single value = 1 row
        return .number(1)
    }
    
    /// COLUMNS - Returns number of columns in reference
    private func evaluateCOLUMNS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "COLUMNS", expected: 1, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        
        if case .array(let rows) = arrayVal, let firstRow = rows.first {
            return .number(Double(firstRow.count))
        }
        
        // Single value = 1 column
        return .number(1)
    }
}

// MARK: - Helper Functions

/// Convert 0-based column index to letter (0 = A, 1 = B, ..., 25 = Z, 26 = AA)
private func columnLetterFrom(index: Int) -> String {
    var col = index + 1
    var result = ""
    
    while col > 0 {
        col -= 1
        let remainder = col % 26
        result = String(UnicodeScalar(UInt8(65 + remainder))) + result
        col /= 26
    }
    
    return result
}
