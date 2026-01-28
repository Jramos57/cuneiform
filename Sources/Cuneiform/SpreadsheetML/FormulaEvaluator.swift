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
        case "XOR":
            return try evaluateXOR(args)
        case "IFNA":
            return try evaluateIFNA(args)
        case "ISNA":
            return try evaluateISNA(args)
        case "ISREF":
            return try evaluateISREF(args)
        case "ISERR":
            return try evaluateISERR(args)
        case "TYPE":
            return try evaluateTYPE(args)
        case "N":
            return try evaluateN(args)
        case "NA":
            return evaluateNA(args)
        case "CELL":
            return try evaluateCELL(args)
        case "INFO":
            return try evaluateINFO(args)
        case "ISFORMULA":
            return try evaluateISFORMULA(args)
        case "ISEVEN":
            return try evaluateISEVEN(args)
        case "ISODD":
            return try evaluateISODD(args)
        case "SHEET":
            return try evaluateSHEET(args)
        case "SHEETS":
            return try evaluateSHEETS(args)
        case "ISTEXT":
            return try evaluateISTEXT(args)
        case "ISNUMBER":
            return try evaluateISNUMBER(args)
        case "ISLOGICAL":
            return try evaluateISLOGICAL(args)
        case "ISBLANK":
            return try evaluateISBLANK(args)
        case "ISNONTEXT":
            return try evaluateISNONTEXT(args)
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
        case "TEXTJOIN":
            return try evaluateTEXTJOIN(args)
        case "NUMBERVALUE":
            return try evaluateNUMBERVALUE(args)
        case "DOLLAR":
            return try evaluateDOLLAR(args)
        case "FIXED":
            return try evaluateFIXED(args)
        case "T":
            return try evaluateT(args)
        case "UNICODE":
            return try evaluateUNICODE(args)
        case "UNICHAR":
            return try evaluateUNICHAR(args)
        case "ARRAYTOTEXT":
            return try evaluateARRAYTOTEXT(args)
        case "VALUETOTEXT":
            return try evaluateVALUETOTEXT(args)
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
        case "ISERROR":
            return try evaluateISERROR(args)
        // Excel 365 high-priority functions
        case "XLOOKUP":
            return try evaluateXLOOKUP(args)
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
        // Dynamic array functions (Excel 365)
        case "FILTER":
            return try evaluateFILTER(args)
        case "SORT":
            return try evaluateSORT(args)
        case "SORTBY":
            return try evaluateSORTBY(args)
        case "UNIQUE":
            return try evaluateUNIQUE(args)
        case "SEQUENCE":
            return try evaluateSEQUENCE(args)
        case "RANDARRAY":
            return try evaluateRANDARRAY(args)
        case "TAKE":
            return try evaluateTAKE(args)
        case "DROP":
            return try evaluateDROP(args)
        case "EXPAND":
            return try evaluateEXPAND(args)
        case "VSTACK":
            return try evaluateVSTACK(args)
        case "HSTACK":
            return try evaluateHSTACK(args)
        case "TOCOL":
            return try evaluateTOCOL(args)
        case "TOROW":
            return try evaluateTOROW(args)
        // Database functions
        case "DSUM":
            return try evaluateDSUM(args)
        case "DAVERAGE":
            return try evaluateDAVERAGE(args)
        case "DCOUNT":
            return try evaluateDCOUNT(args)
        case "DCOUNTA":
            return try evaluateDCOUNTA(args)
        case "DMAX":
            return try evaluateDMAX(args)
        case "DMIN":
            return try evaluateDMIN(args)
        case "DGET":
            return try evaluateDGET(args)
        case "DPRODUCT":
            return try evaluateDPRODUCT(args)
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
        case "COVARIANCE.P":
            return try evaluateCOVARIANCE_P(args)
        case "COVARIANCE.S":
            return try evaluateCOVARIANCE_S(args)
        case "SKEW":
            return try evaluateSKEW(args)
        case "KURT":
            return try evaluateKURT(args)
        case "GEOMEAN":
            return try evaluateGEOMEAN(args)
        case "HARMEAN":
            return try evaluateHARMEAN(args)
        case "AVEDEV":
            return try evaluateAVEDEV(args)
        case "DEVSQ":
            return try evaluateDEVSQ(args)
        case "STANDARDIZE":
            return try evaluateSTANDARDIZE(args)
        case "CONFIDENCE.NORM":
            return try evaluateCONFIDENCE_NORM(args)
        case "FORECAST", "FORECAST.LINEAR":
            return try evaluateFORECAST(args)
        case "PERCENTILE.EXC":
            return try evaluatePERCENTILE_EXC(args)
        case "QUARTILE.EXC":
            return try evaluateQUARTILE_EXC(args)
        case "PERCENTRANK.INC":
            return try evaluatePERCENTRANK_INC(args)
        case "PERCENTRANK.EXC":
            return try evaluatePERCENTRANK_EXC(args)
        case "NORM.DIST":
            return try evaluateNORM_DIST(args)
        case "NORM.INV":
            return try evaluateNORM_INV(args)
        case "NORM.S.DIST":
            return try evaluateNORM_S_DIST(args)
        case "NORM.S.INV":
            return try evaluateNORM_S_INV(args)
        case "BINOM.DIST":
            return try evaluateBINOM_DIST(args)
        case "BINOM.INV":
            return try evaluateBINOM_INV(args)
        case "POISSON.DIST", "POISSON":
            return try evaluatePOISSON_DIST(args)
        case "EXPON.DIST":
            return try evaluateEXPON_DIST(args)
        case "CHISQ.DIST":
            return try evaluateCHISQ_DIST(args)
        case "CHISQ.INV":
            return try evaluateCHISQ_INV(args)
        case "T.DIST":
            return try evaluateT_DIST(args)
        case "T.INV":
            return try evaluateT_INV(args)
        case "F.DIST":
            return try evaluateF_DIST(args)
        case "F.INV":
            return try evaluateF_INV(args)
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
        case "POWER":
            return try evaluatePOWER(args)
        case "MROUND":
            return try evaluateMROUND(args)
        case "EVEN":
            return try evaluateEVEN(args)
        case "ODD":
            return try evaluateODD(args)
        case "QUOTIENT":
            return try evaluateQUOTIENT(args)
        case "RAND":
            return evaluateRAND(args)
        case "RANDBETWEEN":
            return try evaluateRANDBETWEEN(args)
        case "COMBIN":
            return try evaluateCOMBIN(args)
        case "PERMUT":
            return try evaluatePERMUT(args)
        case "MULTINOMIAL":
            return try evaluateMULTINOMIAL(args)
        // Engineering functions
        case "CONVERT":
            return try evaluateCONVERT(args)
        case "DELTA":
            return try evaluateDELTA(args)
        case "GESTEP":
            return try evaluateGESTEP(args)
        case "DEC2BIN":
            return try evaluateDEC2BIN(args)
        case "DEC2OCT":
            return try evaluateDEC2OCT(args)
        case "DEC2HEX":
            return try evaluateDEC2HEX(args)
        case "BIN2DEC":
            return try evaluateBIN2DEC(args)
        case "OCT2DEC":
            return try evaluateOCT2DEC(args)
        case "HEX2DEC":
            return try evaluateHEX2DEC(args)
        case "BIN2HEX":
            return try evaluateBIN2HEX(args)
        case "HEX2BIN":
            return try evaluateHEX2BIN(args)
        case "HEX2OCT":
            return try evaluateHEX2OCT(args)
        case "OCT2BIN":
            return try evaluateOCT2BIN(args)
        case "OCT2HEX":
            return try evaluateOCT2HEX(args)
        case "BIN2OCT":
            return try evaluateBIN2OCT(args)
        case "BITAND":
            return try evaluateBITAND(args)
        case "BITOR":
            return try evaluateBITOR(args)
        case "BITXOR":
            return try evaluateBITXOR(args)
        case "BITLSHIFT":
            return try evaluateBITLSHIFT(args)
        case "BITRSHIFT":
            return try evaluateBITRSHIFT(args)
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
        case "DB":
            return try evaluateDB(args)
        case "DDB":
            return try evaluateDDB(args)
        case "SLN":
            return try evaluateSLN(args)
        case "SYD":
            return try evaluateSYD(args)
        case "VDB":
            return try evaluateVDB(args)
        case "PRICE":
            return try evaluatePRICE(args)
        case "YIELD":
            return try evaluateYIELD(args)
        case "ACCRINT":
            return try evaluateACCRINT(args)
        case "CUMIPMT":
            return try evaluateCUMIPMT(args)
        case "CUMPRINC":
            return try evaluateCUMPRINC(args)
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
    
    /// COVARIANCE.P - Population covariance
    private func evaluateCOVARIANCE_P(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
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
        
        var covariance: Double = 0
        for i in 0..<array1.count {
            covariance += (array1[i] - mean1) * (array2[i] - mean2)
        }
        
        return .number(covariance / n)
    }
    
    /// COVARIANCE.S - Sample covariance
    private func evaluateCOVARIANCE_S(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let array1Val = try evaluate(args[0])
        let array2Val = try evaluate(args[1])
        
        let array1 = flattenToNumbers(array1Val)
        let array2 = flattenToNumbers(array2Val)
        
        guard array1.count == array2.count, array1.count > 1 else {
            return .error("N/A")
        }
        
        let n = Double(array1.count)
        let mean1 = array1.reduce(0, +) / n
        let mean2 = array2.reduce(0, +) / n
        
        var covariance: Double = 0
        for i in 0..<array1.count {
            covariance += (array1[i] - mean1) * (array2[i] - mean2)
        }
        
        return .number(covariance / (n - 1))
    }
    
    /// SKEW - Skewness of a distribution
    private func evaluateSKEW(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard numbers.count >= 3 else {
            return .error("DIV/0")
        }
        
        let n = Double(numbers.count)
        let mean = numbers.reduce(0, +) / n
        
        var m2: Double = 0
        var m3: Double = 0
        
        for num in numbers {
            let diff = num - mean
            m2 += diff * diff
            m3 += diff * diff * diff
        }
        
        let variance = m2 / n
        let stdDev = sqrt(variance)
        
        guard stdDev > 0 else {
            return .error("DIV/0")
        }
        
        let skewness = (n / ((n - 1) * (n - 2))) * (m3 / pow(stdDev, 3))
        return .number(skewness)
    }
    
    /// KURT - Kurtosis of a distribution
    private func evaluateKURT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard numbers.count >= 4 else {
            return .error("DIV/0")
        }
        
        let n = Double(numbers.count)
        let mean = numbers.reduce(0, +) / n
        
        var m2: Double = 0
        var m4: Double = 0
        
        for num in numbers {
            let diff = num - mean
            let diff2 = diff * diff
            m2 += diff2
            m4 += diff2 * diff2
        }
        
        let variance = m2 / n
        
        guard variance > 0 else {
            return .error("DIV/0")
        }
        
        let kurtosis = (n * (n + 1) * m4) / ((n - 1) * (n - 2) * (n - 3) * variance * variance) - 
                       (3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3))
        return .number(kurtosis)
    }
    
    /// GEOMEAN - Geometric mean
    private func evaluateGEOMEAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        // Check for non-positive values
        for num in numbers {
            if num <= 0 {
                return .error("NUM")
            }
        }
        
        let product = numbers.reduce(1.0) { $0 * $1 }
        let geomean = pow(product, 1.0 / Double(numbers.count))
        
        return .number(geomean)
    }
    
    /// HARMEAN - Harmonic mean
    private func evaluateHARMEAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        // Check for zero or negative values
        var sumReciprocals: Double = 0
        for num in numbers {
            if num <= 0 {
                return .error("NUM")
            }
            sumReciprocals += 1.0 / num
        }
        
        let harmean = Double(numbers.count) / sumReciprocals
        return .number(harmean)
    }
    
    /// AVEDEV - Average of absolute deviations from mean
    private func evaluateAVEDEV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        let mean = numbers.reduce(0, +) / Double(numbers.count)
        let sumAbsDev = numbers.reduce(0) { $0 + abs($1 - mean) }
        
        return .number(sumAbsDev / Double(numbers.count))
    }
    
    /// DEVSQ - Sum of squares of deviations from mean
    private func evaluateDEVSQ(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        let mean = numbers.reduce(0, +) / Double(numbers.count)
        let devsq = numbers.reduce(0) { result, num in
            let diff = num - mean
            return result + diff * diff
        }
        
        return .number(devsq)
    }
    
    /// STANDARDIZE - Returns a normalized value
    private func evaluateSTANDARDIZE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let meanVal = try evaluate(args[1])
        let stdDevVal = try evaluate(args[2])
        
        guard let x = xVal.asDouble,
              let mean = meanVal.asDouble,
              let stdDev = stdDevVal.asDouble else {
            return .error("VALUE")
        }
        
        guard stdDev > 0 else {
            return .error("NUM")
        }
        
        return .number((x - mean) / stdDev)
    }
    
    /// CONFIDENCE.NORM - Confidence interval for a population mean
    private func evaluateCONFIDENCE_NORM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let alphaVal = try evaluate(args[0])
        let stdDevVal = try evaluate(args[1])
        let sizeVal = try evaluate(args[2])
        
        guard let alpha = alphaVal.asDouble,
              let stdDev = stdDevVal.asDouble,
              let size = sizeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard alpha > 0, alpha < 1, stdDev > 0, size >= 1 else {
            return .error("NUM")
        }
        
        // Z-score for confidence level (approximate for common values)
        // For alpha=0.05, z≈1.96
        let z: Double
        if abs(alpha - 0.05) < 0.001 {
            z = 1.95996
        } else if abs(alpha - 0.01) < 0.001 {
            z = 2.57583
        } else {
            // Approximation using normal distribution
            z = sqrt(2) * erfInv(1 - alpha)
        }
        
        let margin = z * stdDev / sqrt(size)
        return .number(margin)
    }
    
    // Helper: Inverse error function (approximation)
    private func erfInv(_ x: Double) -> Double {
        let a = 0.147
        let b = 2.0 / (Double.pi * a) + log(1 - x * x) / 2.0
        let c = log(1 - x * x) / a
        
        let sign: Double = x < 0 ? -1 : 1
        return sign * sqrt(sqrt(b * b - c) - b)
    }
    
    /// FORECAST / FORECAST.LINEAR - Linear forecast based on existing values
    private func evaluateFORECAST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let knownYsVal = try evaluate(args[1])
        let knownXsVal = try evaluate(args[2])
        
        guard let x = xVal.asDouble else {
            return .error("VALUE")
        }
        
        let knownYs = flattenToNumbers(knownYsVal)
        let knownXs = flattenToNumbers(knownXsVal)
        
        guard !knownYs.isEmpty, knownYs.count == knownXs.count else {
            return .error("N/A")
        }
        
        // Calculate linear regression: y = mx + b
        let n = Double(knownXs.count)
        let sumX = knownXs.reduce(0, +)
        let sumY = knownYs.reduce(0, +)
        let sumXY = zip(knownXs, knownYs).map(*).reduce(0, +)
        let sumX2 = knownXs.map { $0 * $0 }.reduce(0, +)
        
        let denominator = n * sumX2 - sumX * sumX
        guard abs(denominator) > 1e-10 else {
            return .error("DIV/0")
        }
        
        // Slope (m) and intercept (b)
        let m = (n * sumXY - sumX * sumY) / denominator
        let b = (sumY - m * sumX) / n
        
        let forecast = m * x + b
        return .number(forecast)
    }
    
    /// PERCENTILE.EXC - Percentile (exclusive)
    private func evaluatePERCENTILE_EXC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let kVal = try evaluate(args[1])
        
        guard let k = kVal.asDouble else {
            return .error("VALUE")
        }
        
        guard k > 0, k < 1 else {
            return .error("NUM")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        numbers.sort()
        let n = Double(numbers.count)
        
        // Excel uses (n+1)*k - 1 as the index for exclusive
        let index = (n + 1) * k - 1
        
        guard index >= 0 && index < n else {
            return .error("NUM")
        }
        
        let lower = Int(floor(index))
        let upper = Int(ceil(index))
        
        if lower == upper {
            return .number(numbers[lower])
        } else {
            let fraction = index - Double(lower)
            let result = numbers[lower] + fraction * (numbers[upper] - numbers[lower])
            return .number(result)
        }
    }
    
    /// QUARTILE.EXC - Quartile (exclusive)
    private func evaluateQUARTILE_EXC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let quartVal = try evaluate(args[1])
        
        guard let quart = quartVal.asDouble else {
            return .error("VALUE")
        }
        
        let q = Int(quart)
        guard q >= 0 && q <= 4 else {
            return .error("NUM")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty else {
            return .error("NUM")
        }
        
        numbers.sort()
        
        if q == 0 {
            return .number(numbers.first!)
        } else if q == 4 {
            return .number(numbers.last!)
        }
        
        // Use PERCENTILE.EXC for q=1,2,3
        let k = Double(q) * 0.25
        return try evaluatePERCENTILE_EXC([args[0], FormulaExpression.number(k)])
    }
    
    /// PERCENTRANK.INC - Percent rank (inclusive)
    private func evaluatePERCENTRANK_INC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let xVal = try evaluate(args[1])
        let significanceVal = args.count > 2 ? try evaluate(args[2]) : .number(3)
        
        guard let x = xVal.asDouble,
              let significance = significanceVal.asDouble else {
            return .error("VALUE")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty else {
            return .error("N/A")
        }
        
        numbers.sort()
        
        // Find rank of x
        if let exactIndex = numbers.firstIndex(of: x) {
            let rank = Double(exactIndex) / Double(numbers.count - 1)
            let multiplier = pow(10.0, significance)
            return .number(floor(rank * multiplier) / multiplier)
        }
        
        // Interpolate if x is between values
        for i in 0..<numbers.count - 1 {
            if x > numbers[i] && x < numbers[i + 1] {
                let fraction = (x - numbers[i]) / (numbers[i + 1] - numbers[i])
                let rank = (Double(i) + fraction) / Double(numbers.count - 1)
                let multiplier = pow(10.0, significance)
                return .number(floor(rank * multiplier) / multiplier)
            }
        }
        
        return .error("N/A")
    }
    
    /// PERCENTRANK.EXC - Percent rank (exclusive)
    private func evaluatePERCENTRANK_EXC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let xVal = try evaluate(args[1])
        let significanceVal = args.count > 2 ? try evaluate(args[2]) : .number(3)
        
        guard let x = xVal.asDouble,
              let significance = significanceVal.asDouble else {
            return .error("VALUE")
        }
        
        var numbers = flattenToNumbers(arrayVal)
        guard !numbers.isEmpty else {
            return .error("N/A")
        }
        
        numbers.sort()
        let n = Double(numbers.count)
        
        // Find rank of x (exclusive uses n+1 divisor)
        if let exactIndex = numbers.firstIndex(of: x) {
            let rank = (Double(exactIndex) + 1) / (n + 1)
            let multiplier = pow(10.0, significance)
            return .number(floor(rank * multiplier) / multiplier)
        }
        
        // Interpolate if x is between values
        for i in 0..<numbers.count - 1 {
            if x > numbers[i] && x < numbers[i + 1] {
                let fraction = (x - numbers[i]) / (numbers[i + 1] - numbers[i])
                let rank = (Double(i) + 1 + fraction) / (n + 1)
                let multiplier = pow(10.0, significance)
                return .number(floor(rank * multiplier) / multiplier)
            }
        }
        
        return .error("N/A")
    }
    
    /// NORM.DIST - Normal distribution
    private func evaluateNORM_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 4 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let meanVal = try evaluate(args[1])
        let stdDevVal = try evaluate(args[2])
        let cumulativeVal = try evaluate(args[3])
        
        guard let x = xVal.asDouble,
              let mean = meanVal.asDouble,
              let stdDev = stdDevVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard stdDev > 0 else {
            return .error("NUM")
        }
        
        let z = (x - mean) / stdDev
        
        if cumulative != 0 {
            // Cumulative distribution function
            let result = 0.5 * (1 + erf(z / sqrt(2)))
            return .number(result)
        } else {
            // Probability density function
            let coefficient = 1.0 / (stdDev * sqrt(2 * Double.pi))
            let exponent = -0.5 * z * z
            let result = coefficient * exp(exponent)
            return .number(result)
        }
    }
    
    /// NORM.INV - Inverse normal distribution
    private func evaluateNORM_INV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let probabilityVal = try evaluate(args[0])
        let meanVal = try evaluate(args[1])
        let stdDevVal = try evaluate(args[2])
        
        guard let probability = probabilityVal.asDouble,
              let mean = meanVal.asDouble,
              let stdDev = stdDevVal.asDouble else {
            return .error("VALUE")
        }
        
        guard probability > 0, probability < 1, stdDev > 0 else {
            return .error("NUM")
        }
        
        // Use inverse error function
        let z = sqrt(2) * erfInv(2 * probability - 1)
        let result = mean + z * stdDev
        return .number(result)
    }
    
    /// NORM.S.DIST - Standard normal distribution
    private func evaluateNORM_S_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let zVal = try evaluate(args[0])
        let cumulativeVal = try evaluate(args[1])
        
        guard let z = zVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        if cumulative != 0 {
            // Cumulative distribution function
            let result = 0.5 * (1 + erf(z / sqrt(2)))
            return .number(result)
        } else {
            // Probability density function
            let result = exp(-0.5 * z * z) / sqrt(2 * Double.pi)
            return .number(result)
        }
    }
    
    /// NORM.S.INV - Inverse standard normal distribution
    private func evaluateNORM_S_INV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let probabilityVal = try evaluate(args[0])
        
        guard let probability = probabilityVal.asDouble else {
            return .error("VALUE")
        }
        
        guard probability > 0, probability < 1 else {
            return .error("NUM")
        }
        
        // Use inverse error function
        let z = sqrt(2) * erfInv(2 * probability - 1)
        return .number(z)
    }
    
    /// BINOM.DIST - Binomial distribution
    private func evaluateBINOM_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 4 else {
            return .error("VALUE")
        }
        
        let successesVal = try evaluate(args[0])
        let trialsVal = try evaluate(args[1])
        let probVal = try evaluate(args[2])
        let cumulativeVal = try evaluate(args[3])
        
        guard let successes = successesVal.asDouble,
              let trials = trialsVal.asDouble,
              let prob = probVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        let k = Int(successes)
        let n = Int(trials)
        
        guard k >= 0, n >= 0, k <= n, prob >= 0, prob <= 1 else {
            return .error("NUM")
        }
        
        if cumulative != 0 {
            // Cumulative probability: P(X <= k)
            var sum = 0.0
            for i in 0...k {
                sum += binomialPMF(i, n, prob)
            }
            return .number(sum)
        } else {
            // Probability mass function: P(X = k)
            return .number(binomialPMF(k, n, prob))
        }
    }
    
    /// BINOM.INV - Inverse binomial distribution
    private func evaluateBINOM_INV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let trialsVal = try evaluate(args[0])
        let probVal = try evaluate(args[1])
        let alphaVal = try evaluate(args[2])
        
        guard let trials = trialsVal.asDouble,
              let prob = probVal.asDouble,
              let alpha = alphaVal.asDouble else {
            return .error("VALUE")
        }
        
        let n = Int(trials)
        
        guard n >= 0, prob >= 0, prob <= 1, alpha >= 0, alpha <= 1 else {
            return .error("NUM")
        }
        
        // Find smallest k where cumulative probability >= alpha
        var cumulative = 0.0
        for k in 0...n {
            cumulative += binomialPMF(k, n, prob)
            if cumulative >= alpha {
                return .number(Double(k))
            }
        }
        
        return .number(Double(n))
    }
    
    /// POISSON.DIST - Poisson distribution
    private func evaluatePOISSON_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let meanVal = try evaluate(args[1])
        let cumulativeVal = try evaluate(args[2])
        
        guard let x = xVal.asDouble,
              let mean = meanVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        let k = Int(x)
        
        guard k >= 0, mean > 0 else {
            return .error("NUM")
        }
        
        if cumulative != 0 {
            // Cumulative probability: P(X <= k)
            var sum = 0.0
            for i in 0...k {
                sum += poissonPMF(i, mean)
            }
            return .number(sum)
        } else {
            // Probability mass function: P(X = k)
            return .number(poissonPMF(k, mean))
        }
    }
    
    /// EXPON.DIST - Exponential distribution
    private func evaluateEXPON_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let lambdaVal = try evaluate(args[1])
        let cumulativeVal = try evaluate(args[2])
        
        guard let x = xVal.asDouble,
              let lambda = lambdaVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard x >= 0, lambda > 0 else {
            return .error("NUM")
        }
        
        if cumulative != 0 {
            // Cumulative distribution: 1 - e^(-λx)
            return .number(1 - exp(-lambda * x))
        } else {
            // Probability density: λe^(-λx)
            return .number(lambda * exp(-lambda * x))
        }
    }
    
    /// CHISQ.DIST - Chi-squared distribution (simplified)
    private func evaluateCHISQ_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let dfVal = try evaluate(args[1])
        let cumulativeVal = try evaluate(args[2])
        
        guard let x = xVal.asDouble,
              let df = dfVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard x >= 0, df >= 1 else {
            return .error("NUM")
        }
        
        // Simplified approximation
        if cumulative != 0 {
            // Use normal approximation for large df
            if df > 30 {
                let z = (pow(x / df, 1.0/3.0) - (1 - 2.0/(9*df))) / sqrt(2.0/(9*df))
                return .number(0.5 * (1 + erf(z / sqrt(2))))
            }
            // Simple approximation
            return .number(min(1.0, x / (x + df)))
        } else {
            // PDF approximation
            let term1 = pow(x, df/2 - 1)
            let term2 = exp(-x/2)
            return .number(term1 * term2 / (pow(2, df/2) * tgamma(df/2)))
        }
    }
    
    /// CHISQ.INV - Inverse chi-squared (simplified)
    private func evaluateCHISQ_INV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let probVal = try evaluate(args[0])
        let dfVal = try evaluate(args[1])
        
        guard let prob = probVal.asDouble,
              let df = dfVal.asDouble else {
            return .error("VALUE")
        }
        
        guard prob > 0, prob < 1, df >= 1 else {
            return .error("NUM")
        }
        
        // Simple approximation: chi-squared ≈ df * (prob)^2
        return .number(df * prob / (1 - prob))
    }
    
    /// T.DIST - t-distribution (simplified)
    private func evaluateT_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let dfVal = try evaluate(args[1])
        let cumulativeVal = try evaluate(args[2])
        
        guard let x = xVal.asDouble,
              let df = dfVal.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard df >= 1 else {
            return .error("NUM")
        }
        
        // Simplified: for large df, t-distribution approaches normal
        if cumulative != 0 {
            if df > 30 {
                // Use normal approximation
                return .number(0.5 * (1 + erf(x / sqrt(2))))
            }
            // Simple approximation
            let result = 0.5 * (1 + x / sqrt(x * x + df))
            return .number(result)
        } else {
            // PDF: simplified
            let term = pow(1 + x * x / df, -(df + 1) / 2)
            return .number(term / sqrt(df * Double.pi))
        }
    }
    
    /// T.INV - Inverse t-distribution (simplified)
    private func evaluateT_INV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let probVal = try evaluate(args[0])
        let dfVal = try evaluate(args[1])
        
        guard let prob = probVal.asDouble,
              let df = dfVal.asDouble else {
            return .error("VALUE")
        }
        
        guard prob > 0, prob < 1, df >= 1 else {
            return .error("NUM")
        }
        
        // For large df, use normal inverse
        if df > 30 {
            let z = sqrt(2) * erfInv(2 * prob - 1)
            return .number(z)
        }
        
        // Simple approximation
        let z = sqrt(2) * erfInv(2 * prob - 1)
        return .number(z * sqrt(1 + z * z / (4 * df)))
    }
    
    /// F.DIST - F-distribution (simplified)
    private func evaluateF_DIST(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 4 else {
            return .error("VALUE")
        }
        
        let xVal = try evaluate(args[0])
        let df1Val = try evaluate(args[1])
        let df2Val = try evaluate(args[2])
        let cumulativeVal = try evaluate(args[3])
        
        guard let x = xVal.asDouble,
              let df1 = df1Val.asDouble,
              let df2 = df2Val.asDouble,
              let cumulative = cumulativeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard x >= 0, df1 >= 1, df2 >= 1 else {
            return .error("NUM")
        }
        
        // Simplified approximation
        if cumulative != 0 {
            // CDF approximation
            let t = df1 * x / (df1 * x + df2)
            return .number(t)
        } else {
            // PDF approximation
            let term1 = pow(df1 * x, df1) * pow(df2, df2)
            let term2 = pow(df1 * x + df2, df1 + df2)
            return .number(term1 / term2)
        }
    }
    
    /// F.INV - Inverse F-distribution (simplified)
    private func evaluateF_INV(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let probVal = try evaluate(args[0])
        let df1Val = try evaluate(args[1])
        let df2Val = try evaluate(args[2])
        
        guard let prob = probVal.asDouble,
              let df1 = df1Val.asDouble,
              let df2 = df2Val.asDouble else {
            return .error("VALUE")
        }
        
        guard prob > 0, prob < 1, df1 >= 1, df2 >= 1 else {
            return .error("NUM")
        }
        
        // Simple approximation
        return .number(df2 * prob / (df1 * (1 - prob)))
    }
    
    // Helper: Binomial PMF
    private func binomialPMF(_ k: Int, _ n: Int, _ p: Double) -> Double {
        let coeff = Double(binomialCoefficient(n, k))
        return coeff * pow(p, Double(k)) * pow(1 - p, Double(n - k))
    }
    
    // Helper: Binomial coefficient
    private func binomialCoefficient(_ n: Int, _ k: Int) -> Int {
        guard k >= 0, k <= n else { return 0 }
        var result = 1
        for i in 0..<min(k, n - k) {
            result = result * (n - i) / (i + 1)
        }
        return result
    }
    
    // Helper: Poisson PMF
    private func poissonPMF(_ k: Int, _ mean: Double) -> Double {
        return pow(mean, Double(k)) * exp(-mean) / Double(factorial(k))
    }
    
    // Helper: Factorial
    private func factorial(_ n: Int) -> Int {
        guard n > 1 else { return 1 }
        return n * factorial(n - 1)
    }
    
    // Helper: Error function (approximation)
    private func erf(_ x: Double) -> Double {
        // Abramowitz and Stegun approximation
        let a1 =  0.254829592
        let a2 = -0.284496736
        let a3 =  1.421413741
        let a4 = -1.453152027
        let a5 =  1.061405429
        let p  =  0.3275911
        
        let sign: Double = x < 0 ? -1 : 1
        let absX = abs(x)
        
        let t = 1.0 / (1.0 + p * absX)
        let y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * exp(-absX * absX)
        
        return sign * y
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
    
    /// POWER - Raises a number to a power
    private func evaluatePOWER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let numberVal = try evaluate(args[0])
        let powerVal = try evaluate(args[1])
        
        guard let number = numberVal.asDouble,
              let power = powerVal.asDouble else {
            return .error("VALUE")
        }
        
        return .number(pow(number, power))
    }
    
    /// MROUND - Rounds to the nearest multiple
    private func evaluateMROUND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let numberVal = try evaluate(args[0])
        let multipleVal = try evaluate(args[1])
        
        guard let number = numberVal.asDouble,
              let multiple = multipleVal.asDouble else {
            return .error("VALUE")
        }
        
        guard multiple != 0 else {
            return .error("DIV/0")
        }
        
        let result = round(number / multiple) * multiple
        return .number(result)
    }
    
    /// EVEN - Rounds up to nearest even integer
    private func evaluateEVEN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        let sign = num >= 0 ? 1.0 : -1.0
        let absNum = abs(num)
        let rounded = ceil(absNum)
        let even = rounded.truncatingRemainder(dividingBy: 2) == 0 ? rounded : rounded + 1
        
        return .number(sign * even)
    }
    
    /// ODD - Rounds up to nearest odd integer
    private func evaluateODD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        let sign = num >= 0 ? 1.0 : -1.0
        let absNum = abs(num)
        let rounded = ceil(absNum)
        let odd = rounded.truncatingRemainder(dividingBy: 2) == 1 ? rounded : rounded + 1
        
        return .number(sign * odd)
    }
    
    /// QUOTIENT - Integer portion of a division
    private func evaluateQUOTIENT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let numeratorVal = try evaluate(args[0])
        let denominatorVal = try evaluate(args[1])
        
        guard let numerator = numeratorVal.asDouble,
              let denominator = denominatorVal.asDouble else {
            return .error("VALUE")
        }
        
        guard denominator != 0 else {
            return .error("DIV/0")
        }
        
        return .number(Double(Int(numerator / denominator)))
    }
    
    /// RAND - Random number between 0 and 1
    private func evaluateRAND(_ args: [FormulaExpression]) -> FormulaValue {
        return .number(Double.random(in: 0..<1))
    }
    
    /// RANDBETWEEN - Random integer between two values
    private func evaluateRANDBETWEEN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let bottomVal = try evaluate(args[0])
        let topVal = try evaluate(args[1])
        
        guard let bottom = bottomVal.asDouble,
              let top = topVal.asDouble else {
            return .error("VALUE")
        }
        
        let min = Int(bottom)
        let max = Int(top)
        
        guard min <= max else {
            return .error("NUM")
        }
        
        return .number(Double(Int.random(in: min...max)))
    }
    
    /// COMBIN - Number of combinations
    private func evaluateCOMBIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let nVal = try evaluate(args[0])
        let kVal = try evaluate(args[1])
        
        guard let n = nVal.asDouble,
              let k = kVal.asDouble else {
            return .error("VALUE")
        }
        
        let nInt = Int(n)
        let kInt = Int(k)
        
        guard nInt >= 0, kInt >= 0, kInt <= nInt else {
            return .error("NUM")
        }
        
        func factorial(_ n: Int) -> Double {
            guard n > 1 else { return 1 }
            return (2...n).map(Double.init).reduce(1, *)
        }
        
        let result = factorial(nInt) / (factorial(kInt) * factorial(nInt - kInt))
        return .number(result)
    }
    
    /// PERMUT - Number of permutations
    private func evaluatePERMUT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let nVal = try evaluate(args[0])
        let kVal = try evaluate(args[1])
        
        guard let n = nVal.asDouble,
              let k = kVal.asDouble else {
            return .error("VALUE")
        }
        
        let nInt = Int(n)
        let kInt = Int(k)
        
        guard nInt >= 0, kInt >= 0, kInt <= nInt else {
            return .error("NUM")
        }
        
        func factorial(_ n: Int) -> Double {
            guard n > 1 else { return 1 }
            return (2...n).map(Double.init).reduce(1, *)
        }
        
        let result = factorial(nInt) / factorial(nInt - kInt)
        return .number(result)
    }
    
    /// MULTINOMIAL - Multinomial coefficient
    private func evaluateMULTINOMIAL(_ args: [FormulaExpression]) throws -> FormulaValue {
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
        
        func factorial(_ n: Int) -> Double {
            guard n > 1 else { return 1 }
            return (2...n).map(Double.init).reduce(1, *)
        }
        
        let sum = numbers.reduce(0, +)
        var result = factorial(sum)
        
        for num in numbers {
            result /= factorial(num)
        }
        
        return .number(result)
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
    
    // MARK: - Depreciation Functions
    
    /// DB - Declining balance depreciation
    private func evaluateDB(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 4 && args.count <= 5 else {
            return .error("VALUE")
        }
        
        let costVal = try evaluate(args[0])
        let salvageVal = try evaluate(args[1])
        let lifeVal = try evaluate(args[2])
        let periodVal = try evaluate(args[3])
        let monthVal = args.count > 4 ? try evaluate(args[4]) : .number(12)
        
        guard let cost = costVal.asDouble,
              let salvage = salvageVal.asDouble,
              let life = lifeVal.asDouble,
              let period = periodVal.asDouble,
              let month = monthVal.asDouble else {
            return .error("VALUE")
        }
        
        guard cost >= 0, salvage >= 0, life > 0, period > 0, month >= 1, month <= 12 else {
            return .error("NUM")
        }
        
        // Fixed declining balance rate
        let rate = 1 - pow(salvage / cost, 1 / life)
        
        var depreciation = 0.0
        let p = Int(period)
        
        if p == 1 {
            depreciation = cost * rate * month / 12
        } else {
            var totalDepreciation = cost * rate * month / 12
            for _ in 2...p {
                depreciation = (cost - totalDepreciation) * rate
                totalDepreciation += depreciation
            }
        }
        
        return .number(depreciation)
    }
    
    /// DDB - Double declining balance depreciation
    private func evaluateDDB(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 4 && args.count <= 5 else {
            return .error("VALUE")
        }
        
        let costVal = try evaluate(args[0])
        let salvageVal = try evaluate(args[1])
        let lifeVal = try evaluate(args[2])
        let periodVal = try evaluate(args[3])
        let factorVal = args.count > 4 ? try evaluate(args[4]) : .number(2)
        
        guard let cost = costVal.asDouble,
              let salvage = salvageVal.asDouble,
              let life = lifeVal.asDouble,
              let period = periodVal.asDouble,
              let factor = factorVal.asDouble else {
            return .error("VALUE")
        }
        
        guard cost >= 0, salvage >= 0, life > 0, period > 0, period <= life, factor > 0 else {
            return .error("NUM")
        }
        
        let rate = factor / life
        var bookValue = cost
        var depreciation = 0.0
        
        for _ in 1...Int(period) {
            depreciation = min(bookValue * rate, bookValue - salvage)
            bookValue -= depreciation
        }
        
        return .number(depreciation)
    }
    
    /// SLN - Straight-line depreciation
    private func evaluateSLN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let costVal = try evaluate(args[0])
        let salvageVal = try evaluate(args[1])
        let lifeVal = try evaluate(args[2])
        
        guard let cost = costVal.asDouble,
              let salvage = salvageVal.asDouble,
              let life = lifeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard life > 0 else {
            return .error("NUM")
        }
        
        let depreciation = (cost - salvage) / life
        return .number(depreciation)
    }
    
    /// SYD - Sum-of-years digits depreciation
    private func evaluateSYD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 4 else {
            return .error("VALUE")
        }
        
        let costVal = try evaluate(args[0])
        let salvageVal = try evaluate(args[1])
        let lifeVal = try evaluate(args[2])
        let periodVal = try evaluate(args[3])
        
        guard let cost = costVal.asDouble,
              let salvage = salvageVal.asDouble,
              let life = lifeVal.asDouble,
              let period = periodVal.asDouble else {
            return .error("VALUE")
        }
        
        guard life > 0, period > 0, period <= life else {
            return .error("NUM")
        }
        
        let depreciableBase = cost - salvage
        let sumOfYears = life * (life + 1) / 2
        let depreciation = depreciableBase * (life - period + 1) / sumOfYears
        
        return .number(depreciation)
    }
    
    /// VDB - Variable declining balance (simplified)
    private func evaluateVDB(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 5 && args.count <= 7 else {
            return .error("VALUE")
        }
        
        let costVal = try evaluate(args[0])
        let salvageVal = try evaluate(args[1])
        let lifeVal = try evaluate(args[2])
        let startPeriodVal = try evaluate(args[3])
        let endPeriodVal = try evaluate(args[4])
        let factorVal = args.count > 5 ? try evaluate(args[5]) : .number(2)
        
        guard let cost = costVal.asDouble,
              let salvage = salvageVal.asDouble,
              let life = lifeVal.asDouble,
              let startPeriod = startPeriodVal.asDouble,
              let endPeriod = endPeriodVal.asDouble,
              let factor = factorVal.asDouble else {
            return .error("VALUE")
        }
        
        guard life > 0, startPeriod >= 0, endPeriod <= life, startPeriod < endPeriod else {
            return .error("NUM")
        }
        
        // Simplified: use DDB for the period range
        let rate = factor / life
        var bookValue = cost
        var totalDepreciation = 0.0
        
        for p in 1...Int(endPeriod) {
            let depreciation = min(bookValue * rate, bookValue - salvage)
            bookValue -= depreciation
            
            if Double(p) > startPeriod {
                totalDepreciation += depreciation
            }
        }
        
        return .number(totalDepreciation)
    }
    
    // MARK: - Securities Functions
    
    /// PRICE - Security price (simplified)
    private func evaluatePRICE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 6 && args.count <= 7 else {
            return .error("VALUE")
        }
        
        let settlementVal = try evaluate(args[0])
        let maturityVal = try evaluate(args[1])
        let rateVal = try evaluate(args[2])
        let yldVal = try evaluate(args[3])
        let redemptionVal = try evaluate(args[4])
        let frequencyVal = try evaluate(args[5])
        
        guard let settlement = settlementVal.asDouble,
              let maturity = maturityVal.asDouble,
              let rate = rateVal.asDouble,
              let yld = yldVal.asDouble,
              let redemption = redemptionVal.asDouble,
              let frequency = frequencyVal.asDouble else {
            return .error("VALUE")
        }
        
        guard maturity > settlement, rate >= 0, yld >= 0, redemption > 0 else {
            return .error("NUM")
        }
        
        // Simplified bond pricing formula
        let periods = frequency * (maturity - settlement) / 365.0
        let coupon = redemption * rate / frequency
        
        var price = 0.0
        for n in 1...Int(periods) {
            price += coupon / pow(1 + yld / frequency, Double(n))
        }
        price += redemption / pow(1 + yld / frequency, periods)
        
        return .number(price)
    }
    
    /// YIELD - Security yield (simplified)
    private func evaluateYIELD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 6 && args.count <= 7 else {
            return .error("VALUE")
        }
        
        let settlementVal = try evaluate(args[0])
        let maturityVal = try evaluate(args[1])
        let rateVal = try evaluate(args[2])
        let priceVal = try evaluate(args[3])
        let redemptionVal = try evaluate(args[4])
        let frequencyVal = try evaluate(args[5])
        
        guard let settlement = settlementVal.asDouble,
              let maturity = maturityVal.asDouble,
              let rate = rateVal.asDouble,
              let price = priceVal.asDouble,
              let redemption = redemptionVal.asDouble,
              let frequency = frequencyVal.asDouble else {
            return .error("VALUE")
        }
        
        guard maturity > settlement, rate >= 0, price > 0, redemption > 0 else {
            return .error("NUM")
        }
        
        // Simplified yield approximation
        let periods = frequency * (maturity - settlement) / 365.0
        let coupon = redemption * rate / frequency
        let yield = (coupon * periods + (redemption - price)) / (price * periods)
        
        return .number(yield * frequency)
    }
    
    /// ACCRINT - Accrued interest (simplified)
    private func evaluateACCRINT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 6 && args.count <= 8 else {
            return .error("VALUE")
        }
        
        let issueVal = try evaluate(args[0])
        let _ = try evaluate(args[1])  // firstInterestVal - not used in simplified implementation
        let settlementVal = try evaluate(args[2])
        let rateVal = try evaluate(args[3])
        let parVal = try evaluate(args[4])
        let _ = try evaluate(args[5])  // frequencyVal - not used in simplified implementation
        
        guard let issue = issueVal.asDouble,
              let settlement = settlementVal.asDouble,
              let rate = rateVal.asDouble,
              let par = parVal.asDouble else {
            return .error("VALUE")
        }
        
        guard settlement > issue, rate >= 0, par > 0 else {
            return .error("NUM")
        }
        
        // Simplified: accrued interest = par * rate * (days / 365)
        let days = settlement - issue
        let accruedInterest = par * rate * days / 365.0
        
        return .number(accruedInterest)
    }
    
    /// CUMIPMT - Cumulative interest payment
    private func evaluateCUMIPMT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 6 else {
            return .error("VALUE")
        }
        
        let rateVal = try evaluate(args[0])
        let nperVal = try evaluate(args[1])
        let pvVal = try evaluate(args[2])
        let startPeriodVal = try evaluate(args[3])
        let endPeriodVal = try evaluate(args[4])
        let typeVal = try evaluate(args[5])
        
        guard let rate = rateVal.asDouble,
              let nper = nperVal.asDouble,
              let pv = pvVal.asDouble,
              let startPeriod = startPeriodVal.asDouble,
              let endPeriod = endPeriodVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard rate > 0, nper > 0, startPeriod >= 1, endPeriod <= nper, startPeriod <= endPeriod else {
            return .error("NUM")
        }
        
        // Calculate cumulative interest
        var cumInterest = 0.0
        var balance = pv
        
        let payment: Double
        if type == 0 {
            payment = pv * rate / (1 - pow(1 + rate, -nper))
        } else {
            payment = pv * rate / ((1 + rate) * (1 - pow(1 + rate, -nper)))
        }
        
        for period in 1...Int(endPeriod) {
            let interest = balance * rate
            if Double(period) >= startPeriod {
                cumInterest += interest
            }
            balance -= (payment - interest)
        }
        
        return .number(-cumInterest)
    }
    
    /// CUMPRINC - Cumulative principal payment
    private func evaluateCUMPRINC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 6 else {
            return .error("VALUE")
        }
        
        let rateVal = try evaluate(args[0])
        let nperVal = try evaluate(args[1])
        let pvVal = try evaluate(args[2])
        let startPeriodVal = try evaluate(args[3])
        let endPeriodVal = try evaluate(args[4])
        let typeVal = try evaluate(args[5])
        
        guard let rate = rateVal.asDouble,
              let nper = nperVal.asDouble,
              let pv = pvVal.asDouble,
              let startPeriod = startPeriodVal.asDouble,
              let endPeriod = endPeriodVal.asDouble,
              let type = typeVal.asDouble else {
            return .error("VALUE")
        }
        
        guard rate > 0, nper > 0, startPeriod >= 1, endPeriod <= nper, startPeriod <= endPeriod else {
            return .error("NUM")
        }
        
        // Calculate cumulative principal
        var cumPrincipal = 0.0
        var balance = pv
        
        let payment: Double
        if type == 0 {
            payment = pv * rate / (1 - pow(1 + rate, -nper))
        } else {
            payment = pv * rate / ((1 + rate) * (1 - pow(1 + rate, -nper)))
        }
        
        for period in 1...Int(endPeriod) {
            let interest = balance * rate
            let principal = payment - interest
            if Double(period) >= startPeriod {
                cumPrincipal += principal
            }
            balance -= principal
        }
        
        return .number(-cumPrincipal)
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
    
    /// NUMBERVALUE - Converts text to number in locale-independent way
    private func evaluateNUMBERVALUE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 3 else {
            return .error("VALUE")
        }
        
        let textVal = try evaluate(args[0])
        guard case .string(var text) = textVal else {
            return .error("VALUE")
        }
        
        // Remove whitespace
        text = text.trimmingCharacters(in: .whitespaces)
        
        // Handle optional decimal and group separators
        if args.count > 1 {
            let decimalSepVal = try evaluate(args[1])
            if case .string(let decimalSep) = decimalSepVal {
                text = text.replacingOccurrences(of: decimalSep, with: ".")
            }
        }
        
        if args.count > 2 {
            let groupSepVal = try evaluate(args[2])
            if case .string(let groupSep) = groupSepVal {
                text = text.replacingOccurrences(of: groupSep, with: "")
            }
        }
        
        // Remove currency symbols and percent
        let cleaned = text.replacingOccurrences(of: "$", with: "")
                          .replacingOccurrences(of: "€", with: "")
                          .replacingOccurrences(of: "£", with: "")
                          .replacingOccurrences(of: "%", with: "")
        
        guard let number = Double(cleaned) else {
            return .error("VALUE")
        }
        
        return .number(number)
    }
    
    /// DOLLAR - Formats a number as currency text
    private func evaluateDOLLAR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let numberVal = try evaluate(args[0])
        let decimalsVal = args.count > 1 ? try evaluate(args[1]) : .number(2)
        
        guard let number = numberVal.asDouble,
              let decimals = decimalsVal.asDouble else {
            return .error("VALUE")
        }
        
        let dec = Int(decimals)
        let baseFormatted = String(format: "%.\(max(0, dec))f", number)
        
        // Add thousands separators
        let parts = baseFormatted.split(separator: ".")
        let intPart = String(parts[0])
        let decPart = parts.count > 1 ? String(parts[1]) : ""
        
        var result = ""
        for (i, char) in intPart.reversed().enumerated() {
            if i > 0 && i % 3 == 0 {
                result.insert(",", at: result.startIndex)
            }
            result.insert(char, at: result.startIndex)
        }
        
        let formatted = dec > 0 ? "$\(result).\(decPart)" : "$\(result)"
        return .string(formatted)
    }
    
    /// FIXED - Formats a number as text with fixed decimals
    private func evaluateFIXED(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 3 else {
            return .error("VALUE")
        }
        
        let numberVal = try evaluate(args[0])
        let decimalsVal = args.count > 1 ? try evaluate(args[1]) : .number(2)
        let noCommasVal = args.count > 2 ? try evaluate(args[2]) : .number(0)
        
        guard let number = numberVal.asDouble,
              let decimals = decimalsVal.asDouble,
              let noCommas = noCommasVal.asDouble else {
            return .error("VALUE")
        }
        
        let dec = Int(decimals)
        var formatted = String(format: "%.\(max(0, dec))f", number)
        
        // Add thousands separators unless noCommas is true
        if noCommas == 0 {
            let parts = formatted.split(separator: ".")
            let intPart = String(parts[0])
            let decPart = parts.count > 1 ? String(parts[1]) : ""
            
            var result = ""
            for (i, char) in intPart.reversed().enumerated() {
                if i > 0 && i % 3 == 0 {
                    result.insert(",", at: result.startIndex)
                }
                result.insert(char, at: result.startIndex)
            }
            
            formatted = decPart.isEmpty ? result : "\(result).\(decPart)"
        }
        
        return .string(formatted)
    }
    
    /// T - Returns text or empty string
    private func evaluateT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        if case .string(let text) = val {
            return .string(text)
        } else {
            return .string("")
        }
    }
    
    /// UNICODE - Returns Unicode code point of first character
    private func evaluateUNICODE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        guard case .string(let text) = val, !text.isEmpty else {
            return .error("VALUE")
        }
        
        let scalar = text.unicodeScalars.first!
        return .number(Double(scalar.value))
    }
    
    /// UNICHAR - Returns character for Unicode code point
    private func evaluateUNICHAR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        guard let number = val.asDouble else {
            return .error("VALUE")
        }
        
        let codePoint = UInt32(number)
        guard let scalar = Unicode.Scalar(codePoint) else {
            return .error("VALUE")
        }
        
        return .string(String(scalar))
    }
    
    /// ARRAYTOTEXT - Converts array to text
    private func evaluateARRAYTOTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        
        // Convert to text representation
        func valueToString(_ val: FormulaValue) -> String {
            switch val {
            case .number(let n):
                return String(n)
            case .string(let s):
                return "\"\(s)\""
            case .boolean(let b):
                return b ? "TRUE" : "FALSE"
            case .error(let e):
                return "#\(e)!"
            case .array(let rows):
                let rowStrings = rows.map { row in
                    row.map { valueToString($0) }.joined(separator: ", ")
                }
                return "{\(rowStrings.joined(separator: "; "))}"
            }
        }
        
        return .string(valueToString(arrayVal))
    }
    
    /// VALUETOTEXT - Converts value to text
    private func evaluateVALUETOTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .number(let n):
            return .string(String(n))
        case .string(let s):
            return .string(s)
        case .boolean(let b):
            return .string(b ? "TRUE" : "FALSE")
        case .error(let e):
            return .string("#\(e)!")
        case .array:
            // For arrays, convert to simple representation
            return try evaluateARRAYTOTEXT(args)
        }
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
    
    // MARK: - Logical & Information Functions (Extended)
    
    /// XOR - Exclusive OR logical operation
    private func evaluateXOR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            throw FormulaError.invalidArgumentCount(function: "XOR", expected: 1, got: 0)
        }
        
        var trueCount = 0
        
        for arg in args {
            let val = try evaluate(arg)
            
            switch val {
            case .number(let num):
                if num != 0 {
                    trueCount += 1
                }
            case .boolean(let b):
                if b {
                    trueCount += 1
                }
            default:
                break
            }
        }
        
        // XOR is true if odd number of arguments are true
        return .number(trueCount % 2 == 1 ? 1 : 0)
    }
    
    /// IFNA - Returns value if not #N/A error
    private func evaluateIFNA(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "IFNA", expected: 2, got: args.count)
        }
        
        let value = try evaluate(args[0])
        
        if case .error(let err) = value, err == "N/A" {
            return try evaluate(args[1])
        }
        
        return value
    }
    
    /// ISNA - Returns TRUE if value is #N/A error
    private func evaluateISNA(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISNA", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        if case .error(let err) = val, err == "N/A" {
            return .number(1)
        }
        
        return .number(0)
    }
    
    /// ISREF - Returns TRUE if value is a reference
    private func evaluateISREF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISREF", expected: 1, got: args.count)
        }
        
        // Check if the argument itself (not evaluated) is a reference
        if case .cellRef(_) = args[0] {
            return .number(1)
        } else if case .range(_, _) = args[0] {
            return .number(1)
        }
        
        return .number(0)
    }
    
    /// ISERR - Returns TRUE if value is any error except #N/A
    private func evaluateISERR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISERR", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        if case .error(let err) = val, err != "N/A" {
            return .number(1)
        }
        
        return .number(0)
    }
    
    /// TYPE - Returns type of value (1=number, 2=text, 4=boolean, 16=error, 64=array)
    private func evaluateTYPE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "TYPE", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .number:
            return .number(1)
        case .string:
            return .number(2)
        case .boolean:
            return .number(4)
        case .error:
            return .number(16)
        case .array:
            return .number(64)
        }
    }
    
    /// N - Converts value to number
    private func evaluateN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "N", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .number(let num):
            return .number(num)
        case .boolean(let b):
            return .number(b ? 1 : 0)
        case .error:
            return val
        default:
            return .number(0)
        }
    }
    
    /// NA - Returns #N/A error
    private func evaluateNA(_ args: [FormulaExpression]) -> FormulaValue {
        return .error("N/A")
    }
    
    /// CELL - Returns information about cell
    private func evaluateCELL(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "CELL", expected: 1, got: args.count)
        }
        
        let infoTypeVal = try evaluate(args[0])
        
        guard case .string(let infoType) = infoTypeVal else {
            return .error("VALUE")
        }
        
        // Simplified implementation - return basic info
        switch infoType.lowercased() {
        case "address":
            return .string("$A$1")
        case "col":
            return .number(1)
        case "row":
            return .number(1)
        case "type":
            return .string("v")  // v = value
        case "width":
            return .number(10)
        default:
            return .error("VALUE")
        }
    }
    
    /// INFO - Returns information about operating environment
    private func evaluateINFO(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "INFO", expected: 1, got: args.count)
        }
        
        let typeVal = try evaluate(args[0])
        
        guard case .string(let type) = typeVal else {
            return .error("VALUE")
        }
        
        // Simplified implementation
        switch type.lowercased() {
        case "system":
            return .string("mac")
        case "osversion":
            return .string("14.0")
        case "release":
            return .string("1.0")
        case "recalc":
            return .string("Automatic")
        case "numfile":
            return .number(1)
        default:
            return .error("VALUE")
        }
    }
    
    /// ISFORMULA - Check if cell contains a formula
    private func evaluateISFORMULA(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        // In this simplified implementation, we check if the expression is a cell reference
        // For now, return false as we don't track formula status in the evaluator
        // Full implementation would need access to the cell's formula property
        return .boolean(false)
    }
    
    /// ISEVEN - Check if number is even
    private func evaluateISEVEN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        return .boolean(intNum % 2 == 0)
    }
    
    /// ISODD - Check if number is odd
    private func evaluateISODD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        return .boolean(intNum % 2 != 0)
    }
    
    /// SHEET - Get sheet number (simplified - returns 1)
    private func evaluateSHEET(_ args: [FormulaExpression]) throws -> FormulaValue {
        // Simplified: always return 1 for current sheet
        // Full implementation would need workbook context
        return .number(1)
    }
    
    /// SHEETS - Get total number of sheets (simplified - returns 1)
    private func evaluateSHEETS(_ args: [FormulaExpression]) throws -> FormulaValue {
        // Simplified: always return 1
        // Full implementation would need workbook context
        return .number(1)
    }
    
    /// ISLOGICAL - Check if value is logical
    private func evaluateISLOGICAL(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        
        if case .boolean = val {
            return .boolean(true)
        }
        return .boolean(false)
    }
    
    /// ISNONTEXT - Check if value is not text
    private func evaluateISNONTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let val = try evaluate(args[0])
        
        if case .string = val {
            return .boolean(false)
        }
        return .boolean(true)
    }
    
    // MARK: - Dynamic Array Functions (Excel 365)
    
    /// FILTER - Filters array based on criteria
    private func evaluateFILTER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "FILTER", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let includeVal = try evaluate(args[1])
        
        guard case .array(let rows) = arrayVal,
              case .array(let includeRows) = includeVal else {
            return .error("VALUE")
        }
        
        var filtered: [[FormulaValue]] = []
        
        for (i, row) in rows.enumerated() {
            if i < includeRows.count {
                let includeRow = includeRows[i]
                if let firstInclude = includeRow.first,
                   let num = firstInclude.asDouble,
                   num != 0 {
                    filtered.append(row)
                }
            }
        }
        
        if filtered.isEmpty {
            if args.count > 2 {
                return try evaluate(args[2])
            }
            return .error("CALC")
        }
        
        return .array(filtered)
    }
    
    /// SORT - Sorts array
    private func evaluateSORT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "SORT", expected: 1, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let sortIndexVal = args.count > 1 ? try evaluate(args[1]) : .number(1)
        let sortOrderVal = args.count > 2 ? try evaluate(args[2]) : .number(1)
        
        guard case .array(var rows) = arrayVal else {
            return .error("VALUE")
        }
        
        let sortIndex = (sortIndexVal.asDouble.map { Int($0) } ?? 1) - 1
        let ascending = (sortOrderVal.asDouble ?? 1) == 1
        
        rows.sort { row1, row2 in
            guard sortIndex >= 0 && sortIndex < row1.count && sortIndex < row2.count else {
                return false
            }
            
            let val1 = row1[sortIndex]
            let val2 = row2[sortIndex]
            
            // Compare values
            switch (val1, val2) {
            case (.number(let n1), .number(let n2)):
                return ascending ? n1 < n2 : n1 > n2
            case (.string(let s1), .string(let s2)):
                return ascending ? s1 < s2 : s1 > s2
            default:
                return false
            }
        }
        
        return .array(rows)
    }
    
    /// SORTBY - Sorts array by another array
    private func evaluateSORTBY(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            throw FormulaError.invalidArgumentCount(function: "SORTBY", expected: 2, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let byArrayVal = try evaluate(args[1])
        let sortOrderVal = args.count > 2 ? try evaluate(args[2]) : .number(1)
        
        guard case .array(let rows) = arrayVal,
              case .array(let byRows) = byArrayVal else {
            return .error("VALUE")
        }
        
        let ascending = (sortOrderVal.asDouble ?? 1) == 1
        
        // Create indexed array with sort keys
        var indexed: [(row: [FormulaValue], sortKey: FormulaValue)] = []
        for (i, row) in rows.enumerated() {
            if i < byRows.count, let sortKey = byRows[i].first {
                indexed.append((row, sortKey))
            }
        }
        
        // Sort by sort keys
        indexed.sort { item1, item2 in
            switch (item1.sortKey, item2.sortKey) {
            case (.number(let n1), .number(let n2)):
                return ascending ? n1 < n2 : n1 > n2
            case (.string(let s1), .string(let s2)):
                return ascending ? s1 < s2 : s1 > s2
            default:
                return false
            }
        }
        
        return .array(indexed.map { $0.row })
    }
    
    /// UNIQUE - Returns unique values from array
    private func evaluateUNIQUE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "UNIQUE", expected: 1, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        
        guard case .array(let rows) = arrayVal else {
            return .error("VALUE")
        }
        
        var unique: [[FormulaValue]] = []
        var seen: Set<String> = []
        
        for row in rows {
            // Create a key from the row
            let key = row.map { $0.asString }.joined(separator: "|")
            
            if !seen.contains(key) {
                seen.insert(key)
                unique.append(row)
            }
        }
        
        return .array(unique)
    }
    
    /// SEQUENCE - Generates sequence of numbers
    private func evaluateSEQUENCE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "SEQUENCE", expected: 1, got: args.count)
        }
        
        let rowsVal = try evaluate(args[0])
        let colsVal = args.count > 1 ? try evaluate(args[1]) : .number(1)
        let startVal = args.count > 2 ? try evaluate(args[2]) : .number(1)
        let stepVal = args.count > 3 ? try evaluate(args[3]) : .number(1)
        
        guard let numRows = rowsVal.asDouble,
              let numCols = colsVal.asDouble,
              let start = startVal.asDouble,
              let step = stepVal.asDouble else {
            return .error("VALUE")
        }
        
        let rows = Int(numRows)
        let cols = Int(numCols)
        
        var result: [[FormulaValue]] = []
        var current = start
        
        for _ in 0..<rows {
            var row: [FormulaValue] = []
            for _ in 0..<cols {
                row.append(.number(current))
                current += step
            }
            result.append(row)
        }
        
        return .array(result)
    }
    
    /// RANDARRAY - Generates array of random numbers
    private func evaluateRANDARRAY(_ args: [FormulaExpression]) throws -> FormulaValue {
        let rowsVal = args.count > 0 ? try evaluate(args[0]) : .number(1)
        let colsVal = args.count > 1 ? try evaluate(args[1]) : .number(1)
        let minVal = args.count > 2 ? try evaluate(args[2]) : .number(0)
        let maxVal = args.count > 3 ? try evaluate(args[3]) : .number(1)
        let wholeNumberVal = args.count > 4 ? try evaluate(args[4]) : .number(0)
        
        guard let numRows = rowsVal.asDouble,
              let numCols = colsVal.asDouble,
              let min = minVal.asDouble,
              let max = maxVal.asDouble,
              let wholeNumber = wholeNumberVal.asDouble else {
            return .error("VALUE")
        }
        
        let rows = Int(numRows)
        let cols = Int(numCols)
        let isInteger = wholeNumber != 0
        
        var result: [[FormulaValue]] = []
        
        for _ in 0..<rows {
            var row: [FormulaValue] = []
            for _ in 0..<cols {
                let random = Double.random(in: min...max)
                let value = isInteger ? Double(Int(random)) : random
                row.append(.number(value))
            }
            result.append(row)
        }
        
        return .array(result)
    }
    
    // MARK: - More Dynamic Array Functions (Excel 365)
    
    /// TAKE - Returns a specified number of contiguous rows or columns from the start or end of an array
    /// Syntax: TAKE(array, rows, [columns])
    private func evaluateTAKE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let rowsVal = try evaluate(args[1])
        let colsVal = args.count > 2 ? try evaluate(args[2]) : nil
        
        guard let rowCount = rowsVal.asDouble else {
            return .error("VALUE")
        }
        
        var array: [[FormulaValue]]
        if case .array(let arr) = arrayVal {
            array = arr
        } else {
            // Single value treated as 1x1 array
            array = [[arrayVal]]
        }
        
        guard !array.isEmpty else {
            return .array([])
        }
        
        let rowsToTake = Int(rowCount)
        let actualRows = array.count
        
        // Negative means take from end
        let startRow: Int
        let endRow: Int
        if rowsToTake < 0 {
            startRow = max(0, actualRows + rowsToTake)
            endRow = actualRows
        } else {
            startRow = 0
            endRow = min(rowsToTake, actualRows)
        }
        
        var result = Array(array[startRow..<endRow])
        
        // Handle column selection if provided
        if let colsVal = colsVal, let colCount = colsVal.asDouble {
            let colsToTake = Int(colCount)
            result = result.map { row in
                let actualCols = row.count
                if colsToTake < 0 {
                    let startCol = max(0, actualCols + colsToTake)
                    return Array(row[startCol..<actualCols])
                } else {
                    let endCol = min(colsToTake, actualCols)
                    return Array(row[0..<endCol])
                }
            }
        }
        
        return .array(result)
    }
    
    /// DROP - Excludes a specified number of rows or columns from the start or end of an array
    /// Syntax: DROP(array, rows, [columns])
    private func evaluateDROP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let rowsVal = try evaluate(args[1])
        let colsVal = args.count > 2 ? try evaluate(args[2]) : nil
        
        guard let rowCount = rowsVal.asDouble else {
            return .error("VALUE")
        }
        
        var array: [[FormulaValue]]
        if case .array(let arr) = arrayVal {
            array = arr
        } else {
            array = [[arrayVal]]
        }
        
        guard !array.isEmpty else {
            return .array([])
        }
        
        let rowsToDrop = Int(rowCount)
        let actualRows = array.count
        
        // Negative means drop from end
        let startRow: Int
        let endRow: Int
        if rowsToDrop < 0 {
            startRow = 0
            endRow = max(0, actualRows + rowsToDrop)
        } else {
            startRow = min(rowsToDrop, actualRows)
            endRow = actualRows
        }
        
        var result = Array(array[startRow..<endRow])
        
        // Handle column drop if provided
        if let colsVal = colsVal, let colCount = colsVal.asDouble {
            let colsToDrop = Int(colCount)
            result = result.map { row in
                let actualCols = row.count
                if colsToDrop < 0 {
                    let endCol = max(0, actualCols + colsToDrop)
                    return Array(row[0..<endCol])
                } else {
                    let startCol = min(colsToDrop, actualCols)
                    return Array(row[startCol..<actualCols])
                }
            }
        }
        
        return .array(result)
    }
    
    /// EXPAND - Expands or pads an array to specified row and column dimensions
    /// Syntax: EXPAND(array, rows, [columns], [pad_with])
    private func evaluateEXPAND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let rowsVal = try evaluate(args[1])
        let colsVal = args.count > 2 ? try evaluate(args[2]) : nil
        let padVal = args.count > 3 ? try evaluate(args[3]) : .error("N/A")
        
        guard let targetRows = rowsVal.asDouble else {
            return .error("VALUE")
        }
        
        var array: [[FormulaValue]]
        if case .array(let arr) = arrayVal {
            array = arr
        } else {
            array = [[arrayVal]]
        }
        
        let rows = Int(targetRows)
        let cols: Int
        if let colsVal = colsVal, let c = colsVal.asDouble {
            cols = Int(c)
        } else {
            cols = array.first?.count ?? 1
        }
        
        var result: [[FormulaValue]] = []
        
        for r in 0..<rows {
            var row: [FormulaValue] = []
            for c in 0..<cols {
                if r < array.count && c < array[r].count {
                    row.append(array[r][c])
                } else {
                    row.append(padVal)
                }
            }
            result.append(row)
        }
        
        return .array(result)
    }
    
    /// VSTACK - Appends arrays vertically and in sequence to return a larger array
    /// Syntax: VSTACK(array1, [array2], ...)
    private func evaluateVSTACK(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var result: [[FormulaValue]] = []
        var maxCols = 0
        
        // First pass: determine max columns
        for arg in args {
            let val = try evaluate(arg)
            if case .array(let arr) = val {
                for row in arr {
                    maxCols = max(maxCols, row.count)
                }
            } else {
                maxCols = max(maxCols, 1)
            }
        }
        
        // Second pass: build result
        for arg in args {
            let val = try evaluate(arg)
            if case .array(let arr) = val {
                for row in arr {
                    var newRow = row
                    // Pad with #N/A if needed
                    while newRow.count < maxCols {
                        newRow.append(.error("N/A"))
                    }
                    result.append(newRow)
                }
            } else {
                // Single value becomes a row
                var newRow = [val]
                while newRow.count < maxCols {
                    newRow.append(.error("N/A"))
                }
                result.append(newRow)
            }
        }
        
        return .array(result)
    }
    
    /// HSTACK - Appends arrays horizontally and in sequence to return a larger array
    /// Syntax: HSTACK(array1, [array2], ...)
    private func evaluateHSTACK(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var arrays: [[[FormulaValue]]] = []
        var maxRows = 0
        
        // Convert all arguments to arrays and find max rows
        for arg in args {
            let val = try evaluate(arg)
            if case .array(let arr) = val {
                arrays.append(arr)
                maxRows = max(maxRows, arr.count)
            } else {
                // Single value becomes 1x1 array
                arrays.append([[val]])
                maxRows = max(maxRows, 1)
            }
        }
        
        var result: [[FormulaValue]] = []
        
        for r in 0..<maxRows {
            var row: [FormulaValue] = []
            for arr in arrays {
                if r < arr.count {
                    row.append(contentsOf: arr[r])
                } else {
                    // Pad with #N/A for missing rows
                    let cols = arr.first?.count ?? 1
                    for _ in 0..<cols {
                        row.append(.error("N/A"))
                    }
                }
            }
            result.append(row)
        }
        
        return .array(result)
    }
    
    /// TOCOL - Returns the array as one column
    /// Syntax: TOCOL(array, [ignore], [scan_by_column])
    private func evaluateTOCOL(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let ignoreVal = args.count > 1 ? try evaluate(args[1]) : .number(0)
        // scan_by_column parameter (args[2]) controls order - default is by row
        
        guard let ignore = ignoreVal.asDouble else {
            return .error("VALUE")
        }
        
        let ignoreMode = Int(ignore)
        
        var array: [[FormulaValue]]
        if case .array(let arr) = arrayVal {
            array = arr
        } else {
            array = [[arrayVal]]
        }
        
        var result: [[FormulaValue]] = []
        
        for row in array {
            for cell in row {
                // ignoreMode: 0=keep all, 1=ignore blanks, 2=ignore errors, 3=ignore blanks+errors
                var shouldInclude = true
                
                if ignoreMode == 1 || ignoreMode == 3 {
                    // Ignore blanks (treat as missing cells or empty strings)
                    if case .string(let s) = cell, s.isEmpty {
                        shouldInclude = false
                    }
                }
                
                if ignoreMode == 2 || ignoreMode == 3 {
                    // Ignore errors
                    if case .error = cell {
                        shouldInclude = false
                    }
                }
                
                if shouldInclude {
                    result.append([cell])
                }
            }
        }
        
        return .array(result)
    }
    
    /// TOROW - Returns the array as one row
    /// Syntax: TOROW(array, [ignore], [scan_by_column])
    private func evaluateTOROW(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        let arrayVal = try evaluate(args[0])
        let ignoreVal = args.count > 1 ? try evaluate(args[1]) : .number(0)
        
        guard let ignore = ignoreVal.asDouble else {
            return .error("VALUE")
        }
        
        let ignoreMode = Int(ignore)
        
        var array: [[FormulaValue]]
        if case .array(let arr) = arrayVal {
            array = arr
        } else {
            array = [[arrayVal]]
        }
        
        var row: [FormulaValue] = []
        
        for r in array {
            for cell in r {
                var shouldInclude = true
                
                if ignoreMode == 1 || ignoreMode == 3 {
                    if case .string(let s) = cell, s.isEmpty {
                        shouldInclude = false
                    }
                }
                
                if ignoreMode == 2 || ignoreMode == 3 {
                    if case .error = cell {
                        shouldInclude = false
                    }
                }
                
                if shouldInclude {
                    row.append(cell)
                }
            }
        }
        
        return .array([row])
    }
    
    // MARK: - Database Functions
    
    /// DSUM - Sums values in a database that match criteria
    /// Syntax: DSUM(database, field, criteria)
    private func evaluateDSUM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecords(database: args[0], field: args[1], criteria: args[2])
        let sum = records.reduce(0.0) { $0 + $1 }
        return .number(sum)
    }
    
    /// DAVERAGE - Averages values in a database that match criteria
    /// Syntax: DAVERAGE(database, field, criteria)
    private func evaluateDAVERAGE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecords(database: args[0], field: args[1], criteria: args[2])
        guard !records.isEmpty else {
            return .error("DIV/0")
        }
        let avg = records.reduce(0.0, +) / Double(records.count)
        return .number(avg)
    }
    
    /// DCOUNT - Counts cells containing numbers in a database that match criteria
    /// Syntax: DCOUNT(database, field, criteria)
    private func evaluateDCOUNT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecords(database: args[0], field: args[1], criteria: args[2])
        return .number(Double(records.count))
    }
    
    /// DCOUNTA - Counts non-empty cells in a database that match criteria
    /// Syntax: DCOUNTA(database, field, criteria)
    private func evaluateDCOUNTA(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        // For DCOUNTA, we need all matching records including non-numeric
        let records = try getDatabaseRecordsAll(database: args[0], field: args[1], criteria: args[2])
        return .number(Double(records.count))
    }
    
    /// DMAX - Returns the maximum value from selected database entries
    /// Syntax: DMAX(database, field, criteria)
    private func evaluateDMAX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecords(database: args[0], field: args[1], criteria: args[2])
        guard let max = records.max() else {
            return .number(0)
        }
        return .number(max)
    }
    
    /// DMIN - Returns the minimum value from selected database entries
    /// Syntax: DMIN(database, field, criteria)
    private func evaluateDMIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecords(database: args[0], field: args[1], criteria: args[2])
        guard let min = records.min() else {
            return .number(0)
        }
        return .number(min)
    }
    
    /// DGET - Extracts a single value from a database that matches criteria
    /// Syntax: DGET(database, field, criteria)
    private func evaluateDGET(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecordsAll(database: args[0], field: args[1], criteria: args[2])
        
        if records.isEmpty {
            return .error("VALUE")
        } else if records.count > 1 {
            return .error("NUM")  // Multiple records found
        } else {
            return records[0]
        }
    }
    
    /// DPRODUCT - Multiplies values in a database that match criteria
    /// Syntax: DPRODUCT(database, field, criteria)
    private func evaluateDPRODUCT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 else {
            return .error("VALUE")
        }
        
        let records = try getDatabaseRecords(database: args[0], field: args[1], criteria: args[2])
        guard !records.isEmpty else {
            return .number(0)
        }
        let product = records.reduce(1.0) { $0 * $1 }
        return .number(product)
    }
    
    // Helper function for database functions - returns numeric values only
    private func getDatabaseRecords(database: FormulaExpression, field: FormulaExpression, criteria: FormulaExpression) throws -> [Double] {
        let allRecords = try getDatabaseRecordsAll(database: database, field: field, criteria: criteria)
        return allRecords.compactMap { $0.asDouble }
    }
    
    // Helper function for database functions - returns all matching values
    private func getDatabaseRecordsAll(database: FormulaExpression, field: FormulaExpression, criteria: FormulaExpression) throws -> [FormulaValue] {
        // Evaluate database range
        let dbVal = try evaluate(database)
        guard case .array(let dbArray) = dbVal, !dbArray.isEmpty else {
            return []
        }
        
        // First row is headers
        let headers = dbArray[0]
        let dataRows = Array(dbArray.dropFirst())
        
        // Determine field column index
        let fieldVal = try evaluate(field)
        var fieldIndex: Int?
        
        if let fieldNum = fieldVal.asDouble {
            // Field specified as column number (1-based)
            fieldIndex = Int(fieldNum) - 1
        } else if case .string(let fieldName) = fieldVal {
            // Field specified as column name
            fieldIndex = headers.firstIndex { header in
                if case .string(let headerName) = header {
                    return headerName.lowercased() == fieldName.lowercased()
                }
                return false
            }
        }
        
        guard let colIndex = fieldIndex, colIndex >= 0, colIndex < headers.count else {
            return []
        }
        
        // Evaluate criteria range
        let critVal = try evaluate(criteria)
        guard case .array(let critArray) = critVal, critArray.count >= 2 else {
            return []
        }
        
        // First row of criteria is field names, subsequent rows are conditions
        let critHeaders = critArray[0]
        let critRows = Array(critArray.dropFirst())
        
        // Build criteria map: column name -> [values to match]
        var criteriaMap: [String: [FormulaValue]] = [:]
        for (idx, header) in critHeaders.enumerated() {
            if case .string(let headerName) = header {
                var values: [FormulaValue] = []
                for critRow in critRows {
                    if idx < critRow.count {
                        values.append(critRow[idx])
                    }
                }
                criteriaMap[headerName.lowercased()] = values
            }
        }
        
        // Filter data rows based on criteria
        var matchingValues: [FormulaValue] = []
        
        for dataRow in dataRows {
            var matches = true
            
            // Check each criterion
            for (critFieldName, critValues) in criteriaMap {
                // Find column index for this criterion field
                guard let critColIndex = headers.firstIndex(where: { header in
                    if case .string(let name) = header {
                        return name.lowercased() == critFieldName
                    }
                    return false
                }) else {
                    matches = false
                    break
                }
                
                // Check if data row value matches any of the criteria values
                guard critColIndex < dataRow.count else {
                    matches = false
                    break
                }
                
                let cellValue = dataRow[critColIndex]
                var foundMatch = false
                
                for critValue in critValues {
                    if matchesCriteria(cellValue, critValue) {
                        foundMatch = true
                        break
                    }
                }
                
                if !foundMatch {
                    matches = false
                    break
                }
            }
            
            if matches && colIndex < dataRow.count {
                matchingValues.append(dataRow[colIndex])
            }
        }
        
        return matchingValues
    }
    
    // MARK: - Engineering Functions
    
    /// CONVERT - Unit conversion
    private func evaluateCONVERT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            return .error("VALUE")
        }
        
        let numberVal = try evaluate(args[0])
        let fromUnitVal = try evaluate(args[1])
        let toUnitVal = try evaluate(args[2])
        
        guard let number = numberVal.asDouble,
              case .string(let fromUnit) = fromUnitVal,
              case .string(let toUnit) = toUnitVal else {
            return .error("VALUE")
        }
        
        // Comprehensive unit conversion table
        // Format: [unit: (base_unit_multiplier, category)]
        let units: [String: (Double, String)] = [
            // Weight and mass
            "g": (1.0, "weight"), "sg": (1.0, "weight"), "kg": (1000.0, "weight"),
            "lbm": (453.59237, "weight"), "u": (1.66053906660e-27, "weight"),
            "ozm": (28.349523125, "weight"), "stone": (6350.29318, "weight"),
            "ton": (907184.74, "weight"), "grain": (0.06479891, "weight"),
            "cwt": (45359.237, "weight"), "shweight": (45359.237, "weight"),
            "uk_cwt": (50802.34544, "weight"), "lcwt": (50802.34544, "weight"),
            "uk_ton": (1016046.9088, "weight"), "brton": (1016046.9088, "weight"),
            
            // Distance
            "m": (1.0, "distance"), "mi": (1609.344, "distance"), "Nmi": (1852.0, "distance"),
            "in": (0.0254, "distance"), "ft": (0.3048, "distance"), "yd": (0.9144, "distance"),
            "ang": (1e-10, "distance"), "ell": (1.143, "distance"), "ly": (9.46073e15, "distance"),
            "parsec": (3.08568e16, "distance"), "pc": (3.08568e16, "distance"),
            "Pica": (0.00423333333, "distance"), "pica": (0.00423333333, "distance"),
            "survey_mi": (1609.347219, "distance"),
            
            // Time
            "yr": (31557600.0, "time"), "day": (86400.0, "time"), "d": (86400.0, "time"),
            "hr": (3600.0, "time"), "mn": (60.0, "time"), "sec": (1.0, "time"),
            "s": (1.0, "time"),
            
            // Pressure
            "Pa": (1.0, "pressure"), "p": (1.0, "pressure"), "atm": (101325.0, "pressure"),
            "at": (98066.5, "pressure"), "mmHg": (133.322387415, "pressure"),
            "psi": (6894.757293168, "pressure"), "Torr": (133.322368421, "pressure"),
            
            // Force
            "N": (1.0, "force"), "dyn": (1e-5, "force"), "dy": (1e-5, "force"),
            "lbf": (4.4482216152605, "force"), "pond": (0.00980665, "force"),
            
            // Energy
            "J": (1.0, "energy"), "e": (1.0, "energy"), "c": (4.1868, "energy"),
            "cal": (4.1868, "energy"), "eV": (1.602176634e-19, "energy"),
            "ev": (1.602176634e-19, "energy"), "HPh": (2684519.537696172792, "energy"),
            "hh": (2684519.537696172792, "energy"), "Wh": (3600.0, "energy"),
            "wh": (3600.0, "energy"), "flb": (1.3558179483314004, "energy"),
            "BTU": (1055.05585262, "energy"), "btu": (1055.05585262, "energy"),
            
            // Power
            "W": (1.0, "power"), "w": (1.0, "power"), "HP": (745.69987158227022, "power"),
            "h": (745.69987158227022, "power"), "PS": (735.49875, "power"),
            
            // Magnetism
            "T": (1.0, "magnetism"), "ga": (0.0001, "magnetism"),
            
            // Temperature (special handling needed)
            "C": (1.0, "temp"), "F": (1.0, "temp"), "K": (1.0, "temp"),
            
            // Liquid measure
            "tsp": (0.00492892159375, "liquid"), "tspm": (0.000005, "liquid"),
            "tbs": (0.01478676478125, "liquid"), "oz": (0.0295735295625, "liquid"),
            "cup": (0.0002365882365, "liquid"), "pt": (0.000473176473, "liquid"),
            "us_pt": (0.000473176473, "liquid"), "uk_pt": (0.00056826125, "liquid"),
            "qt": (0.000946352946, "liquid"), "uk_qt": (0.0011365225, "liquid"),
            "gal": (0.003785411784, "liquid"), "uk_gal": (0.00454609, "liquid"),
            "l": (0.001, "liquid"), "L": (0.001, "liquid"), "lt": (0.001, "liquid"),
            "ang3": (1e-30, "liquid"), "ang^3": (1e-30, "liquid"),
            "barrel": (0.158987294928, "liquid"), "bushel": (0.03523907016688, "liquid"),
            "GRT": (2.8316846592, "liquid"), "regton": (2.8316846592, "liquid"),
            "MTON": (1.13267386368, "liquid"),
            
            // Area
            "m2": (1.0, "area"), "m^2": (1.0, "area"), "mi2": (2589988.110336, "area"),
            "mi^2": (2589988.110336, "area"), "Nmi2": (3429904.0, "area"),
            "Nmi^2": (3429904.0, "area"), "in2": (0.00064516, "area"),
            "in^2": (0.00064516, "area"), "ft2": (0.09290304, "area"),
            "ft^2": (0.09290304, "area"), "yd2": (0.83612736, "area"),
            "yd^2": (0.83612736, "area"), "ang2": (1e-20, "area"),
            "ang^2": (1e-20, "area"), "Picapt2": (1.792111e-5, "area"),
            "Picapt^2": (1.792111e-5, "area"), "Pica2": (1.792111e-5, "area"),
            "Pica^2": (1.792111e-5, "area"), "Morgen": (2500.0, "area"),
            "ar": (100.0, "area"), "acre": (4046.8564224, "area"),
            "uk_acre": (4046.8564224, "area"), "us_acre": (4046.87261, "area"),
            "ly2": (8.95054e31, "area"), "ly^2": (8.95054e31, "area"),
            "ha": (10000.0, "area"),
            
            // Information
            "bit": (1.0, "info"), "byte": (8.0, "info"),
            
            // Speed
            "m/s": (1.0, "speed"), "m/sec": (1.0, "speed"), "m/h": (0.00027777777777778, "speed"),
            "m/hr": (0.00027777777777778, "speed"), "mph": (0.44704, "speed"),
            "kn": (0.514444444444444, "speed"), "admkn": (0.514444444444444, "speed")
        ]
        
        // Check if units exist and are in same category
        guard let fromConv = units[fromUnit], let toConv = units[toUnit] else {
            return .error("N/A")
        }
        
        let (fromMult, fromCat) = fromConv
        let (toMult, toCat) = toConv
        
        // Temperature requires special handling
        if fromCat == "temp" || toCat == "temp" {
            if fromCat != "temp" || toCat != "temp" {
                return .error("N/A")
            }
            
            var celsius: Double
            switch fromUnit {
            case "C": celsius = number
            case "F": celsius = (number - 32) * 5/9
            case "K": celsius = number - 273.15
            default: return .error("N/A")
            }
            
            let result: Double
            switch toUnit {
            case "C": result = celsius
            case "F": result = celsius * 9/5 + 32
            case "K": result = celsius + 273.15
            default: return .error("N/A")
            }
            
            return .number(result)
        }
        
        // Check category match
        guard fromCat == toCat else {
            return .error("N/A")
        }
        
        // Convert: from -> base -> to
        let result = number * fromMult / toMult
        return .number(result)
    }
    
    /// DELTA - Tests if two values are equal (returns 1 if equal, 0 otherwise)
    private func evaluateDELTA(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let num1Val = try evaluate(args[0])
        let num2Val = args.count > 1 ? try evaluate(args[1]) : .number(0)
        
        guard let num1 = num1Val.asDouble,
              let num2 = num2Val.asDouble else {
            return .error("VALUE")
        }
        
        return .number(num1 == num2 ? 1 : 0)
    }
    
    /// GESTEP - Tests if number >= step (returns 1 if true, 0 otherwise)
    private func evaluateGESTEP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let numVal = try evaluate(args[0])
        let stepVal = args.count > 1 ? try evaluate(args[1]) : .number(0)
        
        guard let num = numVal.asDouble,
              let step = stepVal.asDouble else {
            return .error("VALUE")
        }
        
        return .number(num >= step ? 1 : 0)
    }
    
    /// DEC2BIN - Decimal to binary
    private func evaluateDEC2BIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let numVal = try evaluate(args[0])
        let placesVal = args.count > 1 ? try evaluate(args[1]) : nil
        
        guard let num = numVal.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        
        // Binary range: -512 to 511
        guard intNum >= -512 && intNum <= 511 else {
            return .error("NUM")
        }
        
        let binary: String
        if intNum >= 0 {
            binary = String(intNum, radix: 2)
        } else {
            // Two's complement for negative numbers (10-bit)
            let positive = (1 << 10) + intNum
            binary = String(positive, radix: 2)
        }
        
        if let placesVal = placesVal, let places = placesVal.asDouble {
            let p = Int(places)
            guard p >= 0 else {
                return .error("NUM")
            }
            if binary.count > p {
                return .error("NUM")
            }
            return .string(String(repeating: "0", count: max(0, p - binary.count)) + binary)
        }
        
        return .string(binary)
    }
    
    /// DEC2OCT - Decimal to octal
    private func evaluateDEC2OCT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let numVal = try evaluate(args[0])
        let placesVal = args.count > 1 ? try evaluate(args[1]) : nil
        
        guard let num = numVal.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        
        // Octal range: -536870912 to 536870911 (30-bit)
        guard intNum >= -536870912 && intNum <= 536870911 else {
            return .error("NUM")
        }
        
        let octal: String
        if intNum >= 0 {
            octal = String(intNum, radix: 8)
        } else {
            // Two's complement for negative numbers (30-bit)
            let positive = (1 << 30) + intNum
            octal = String(positive, radix: 8)
        }
        
        if let placesVal = placesVal, let places = placesVal.asDouble {
            let p = Int(places)
            guard p >= 0 else {
                return .error("NUM")
            }
            if octal.count > p {
                return .error("NUM")
            }
            return .string(String(repeating: "0", count: max(0, p - octal.count)) + octal)
        }
        
        return .string(octal)
    }
    
    /// DEC2HEX - Decimal to hexadecimal
    private func evaluateDEC2HEX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        let numVal = try evaluate(args[0])
        let placesVal = args.count > 1 ? try evaluate(args[1]) : nil
        
        guard let num = numVal.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        
        // Hex range: -549755813888 to 549755813887 (40-bit)
        guard intNum >= -549755813888 && intNum <= 549755813887 else {
            return .error("NUM")
        }
        
        let hex: String
        if intNum >= 0 {
            hex = String(intNum, radix: 16).uppercased()
        } else {
            // Two's complement for negative numbers (40-bit)
            let positive = (1 << 40) + intNum
            hex = String(positive, radix: 16).uppercased()
        }
        
        if let placesVal = placesVal, let places = placesVal.asDouble {
            let p = Int(places)
            guard p >= 0 else {
                return .error("NUM")
            }
            if hex.count > p {
                return .error("NUM")
            }
            return .string(String(repeating: "0", count: max(0, p - hex.count)) + hex)
        }
        
        return .string(hex)
    }
    
    /// BIN2DEC - Binary to decimal
    private func evaluateBIN2DEC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let binVal = try evaluate(args[0])
        
        guard case .string(let binary) = binVal else {
            return .error("VALUE")
        }
        
        // Must be 10 digits or fewer
        guard binary.count <= 10, binary.allSatisfy({ $0 == "0" || $0 == "1" }) else {
            return .error("NUM")
        }
        
        // Check if negative (starts with 1 in 10-bit two's complement)
        if binary.count == 10 && binary.first == "1" {
            // Negative number - two's complement
            guard let positive = Int(binary, radix: 2) else {
                return .error("NUM")
            }
            let result = positive - (1 << 10)
            return .number(Double(result))
        } else {
            guard let result = Int(binary, radix: 2) else {
                return .error("NUM")
            }
            return .number(Double(result))
        }
    }
    
    /// OCT2DEC - Octal to decimal
    private func evaluateOCT2DEC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let octVal = try evaluate(args[0])
        
        guard case .string(let octal) = octVal else {
            return .error("VALUE")
        }
        
        // Must be 10 digits or fewer
        guard octal.count <= 10, octal.allSatisfy({ "01234567".contains($0) }) else {
            return .error("NUM")
        }
        
        // Check if negative (starts with 4-7 in 10-digit octal)
        if octal.count == 10, let firstDigit = octal.first, "4567".contains(firstDigit) {
            // Negative number - two's complement
            guard let positive = Int(octal, radix: 8) else {
                return .error("NUM")
            }
            let result = positive - (1 << 30)
            return .number(Double(result))
        } else {
            guard let result = Int(octal, radix: 8) else {
                return .error("NUM")
            }
            return .number(Double(result))
        }
    }
    
    /// HEX2DEC - Hexadecimal to decimal
    private func evaluateHEX2DEC(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            return .error("VALUE")
        }
        
        let hexVal = try evaluate(args[0])
        
        guard case .string(let hex) = hexVal else {
            return .error("VALUE")
        }
        
        // Must be 10 digits or fewer
        let hexUpper = hex.uppercased()
        guard hexUpper.count <= 10, hexUpper.allSatisfy({ "0123456789ABCDEF".contains($0) }) else {
            return .error("NUM")
        }
        
        // Check if negative (starts with 8-F in 10-digit hex)
        if hexUpper.count == 10, let firstDigit = hexUpper.first, "89ABCDEF".contains(firstDigit) {
            // Negative number - two's complement
            guard let positive = Int(hexUpper, radix: 16) else {
                return .error("NUM")
            }
            let result = positive - (1 << 40)
            return .number(Double(result))
        } else {
            guard let result = Int(hexUpper, radix: 16) else {
                return .error("NUM")
            }
            return .number(Double(result))
        }
    }
    
    /// BIN2HEX - Binary to hexadecimal
    private func evaluateBIN2HEX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        // Convert binary to decimal first
        let decVal = try evaluateBIN2DEC([args[0]])
        
        guard let dec = decVal.asDouble else {
            return decVal // Return error from BIN2DEC
        }
        
        // Convert decimal to hex
        let hexArgs = args.count > 1 ? [FormulaExpression.number(dec), args[1]] : [FormulaExpression.number(dec)]
        return try evaluateDEC2HEX(hexArgs)
    }
    
    /// HEX2BIN - Hexadecimal to binary
    private func evaluateHEX2BIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        // Convert hex to decimal first
        let decVal = try evaluateHEX2DEC([args[0]])
        
        guard let dec = decVal.asDouble else {
            return decVal // Return error from HEX2DEC
        }
        
        // Convert decimal to binary
        let binArgs = args.count > 1 ? [FormulaExpression.number(dec), args[1]] : [FormulaExpression.number(dec)]
        return try evaluateDEC2BIN(binArgs)
    }
    
    /// HEX2OCT - Hexadecimal to octal
    private func evaluateHEX2OCT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        // Convert hex to decimal first
        let decVal = try evaluateHEX2DEC([args[0]])
        
        guard let dec = decVal.asDouble else {
            return decVal // Return error from HEX2DEC
        }
        
        // Convert decimal to octal
        let octArgs = args.count > 1 ? [FormulaExpression.number(dec), args[1]] : [FormulaExpression.number(dec)]
        return try evaluateDEC2OCT(octArgs)
    }
    
    /// OCT2BIN - Octal to binary
    private func evaluateOCT2BIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        // Convert octal to decimal first
        let decVal = try evaluateOCT2DEC([args[0]])
        
        guard let dec = decVal.asDouble else {
            return decVal // Return error from OCT2DEC
        }
        
        // Convert decimal to binary
        let binArgs = args.count > 1 ? [FormulaExpression.number(dec), args[1]] : [FormulaExpression.number(dec)]
        return try evaluateDEC2BIN(binArgs)
    }
    
    /// OCT2HEX - Octal to hexadecimal
    private func evaluateOCT2HEX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        // Convert octal to decimal first
        let decVal = try evaluateOCT2DEC([args[0]])
        
        guard let dec = decVal.asDouble else {
            return decVal // Return error from OCT2DEC
        }
        
        // Convert decimal to hex
        let hexArgs = args.count > 1 ? [FormulaExpression.number(dec), args[1]] : [FormulaExpression.number(dec)]
        return try evaluateDEC2HEX(hexArgs)
    }
    
    /// BIN2OCT - Binary to octal
    private func evaluateBIN2OCT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error("VALUE")
        }
        
        // Convert binary to decimal first
        let decVal = try evaluateBIN2DEC([args[0]])
        
        guard let dec = decVal.asDouble else {
            return decVal // Return error from BIN2DEC
        }
        
        // Convert decimal to octal
        let octArgs = args.count > 1 ? [FormulaExpression.number(dec), args[1]] : [FormulaExpression.number(dec)]
        return try evaluateDEC2OCT(octArgs)
    }
    
    /// BITAND - Bitwise AND operation
    private func evaluateBITAND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let num1Val = try evaluate(args[0])
        let num2Val = try evaluate(args[1])
        
        guard let num1 = num1Val.asDouble,
              let num2 = num2Val.asDouble else {
            return .error("VALUE")
        }
        
        let int1 = Int(num1)
        let int2 = Int(num2)
        
        // Must be non-negative and less than 2^48
        guard int1 >= 0 && int1 < (1 << 48) && int2 >= 0 && int2 < (1 << 48) else {
            return .error("NUM")
        }
        
        let result = int1 & int2
        return .number(Double(result))
    }
    
    /// BITOR - Bitwise OR operation
    private func evaluateBITOR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let num1Val = try evaluate(args[0])
        let num2Val = try evaluate(args[1])
        
        guard let num1 = num1Val.asDouble,
              let num2 = num2Val.asDouble else {
            return .error("VALUE")
        }
        
        let int1 = Int(num1)
        let int2 = Int(num2)
        
        // Must be non-negative and less than 2^48
        guard int1 >= 0 && int1 < (1 << 48) && int2 >= 0 && int2 < (1 << 48) else {
            return .error("NUM")
        }
        
        let result = int1 | int2
        return .number(Double(result))
    }
    
    /// BITXOR - Bitwise XOR operation
    private func evaluateBITXOR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let num1Val = try evaluate(args[0])
        let num2Val = try evaluate(args[1])
        
        guard let num1 = num1Val.asDouble,
              let num2 = num2Val.asDouble else {
            return .error("VALUE")
        }
        
        let int1 = Int(num1)
        let int2 = Int(num2)
        
        // Must be non-negative and less than 2^48
        guard int1 >= 0 && int1 < (1 << 48) && int2 >= 0 && int2 < (1 << 48) else {
            return .error("NUM")
        }
        
        let result = int1 ^ int2
        return .number(Double(result))
    }
    
    /// BITLSHIFT - Bitwise left shift
    private func evaluateBITLSHIFT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let numVal = try evaluate(args[0])
        let shiftVal = try evaluate(args[1])
        
        guard let num = numVal.asDouble,
              let shift = shiftVal.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        let intShift = Int(shift)
        
        // Number must be non-negative and less than 2^48
        guard intNum >= 0 && intNum < (1 << 48) else {
            return .error("NUM")
        }
        
        // Shift amount must be valid
        guard abs(intShift) < 54 else {
            return .error("NUM")
        }
        
        let result: Int
        if intShift >= 0 {
            result = intNum << intShift
        } else {
            result = intNum >> (-intShift)
        }
        
        // Result must be less than 2^48
        guard result < (1 << 48) else {
            return .error("NUM")
        }
        
        return .number(Double(result))
    }
    
    /// BITRSHIFT - Bitwise right shift
    private func evaluateBITRSHIFT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            return .error("VALUE")
        }
        
        let numVal = try evaluate(args[0])
        let shiftVal = try evaluate(args[1])
        
        guard let num = numVal.asDouble,
              let shift = shiftVal.asDouble else {
            return .error("VALUE")
        }
        
        let intNum = Int(num)
        let intShift = Int(shift)
        
        // Number must be non-negative and less than 2^48
        guard intNum >= 0 && intNum < (1 << 48) else {
            return .error("NUM")
        }
        
        // Shift amount must be valid
        guard abs(intShift) < 54 else {
            return .error("NUM")
        }
        
        let result: Int
        if intShift >= 0 {
            result = intNum >> intShift
        } else {
            result = intNum << (-intShift)
        }
        
        // Result must be less than 2^48
        guard result < (1 << 48) else {
            return .error("NUM")
        }
        
        return .number(Double(result))
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
