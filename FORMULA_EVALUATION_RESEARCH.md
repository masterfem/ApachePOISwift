# Excel Formula Evaluation in Swift - Comprehensive Research Summary

## Executive Summary

This document provides a comprehensive analysis of how to implement Excel formula evaluation in Swift, drawing from research on:
- Apache POI (Java) architecture and formula handling
- ECMA-376 Office Open XML specification
- Swift mathematical expression parsers and evaluators
- Best practices for formula parsing and evaluation

The key insight is that Excel formula evaluation requires three distinct phases:
1. **Parsing**: Convert infix formula string (e.g., "=A1+B1*2") into structured tokens
2. **Building**: Create a dependency graph and resolve cell references
3. **Evaluation**: Execute in correct order with proper error handling and circular reference detection

## Current State in ApachePOISwift

### What We Already Have

1. **CellReference.swift** (COMPLETE)
   - Parses A1 notation: "A1", "$A$1", "AA100", etc.
   - Handles absolute/relative references
   - Converts between column letters and indices
   - Status: Fully functional, well-tested

2. **Formula Storage (ExcelCell.swift)**
   - `setFormula(String)` method to store formulas
   - `formula` property to read formula
   - `hasFormula` property to check if cell contains formula
   - Formulas preserved during save/load
   - Status: Basic write-only support (formulas saved but not evaluated)

3. **Formula Persistence**
   - FormulaTests.swift: 18 tests covering basic formula storage
   - Can set, save, and reload formulas without modification
   - Handles complex formulas with nested functions

### What's Missing

1. **Formula Parsing**: No tokenizer or parser
2. **Formula Evaluation**: No calculation engine
3. **Cell Dependencies**: No dependency graph
4. **Error Handling**: No circular reference detection
5. **Function Library**: No built-in Excel function implementations

## Architecture Recommendation

### Phase 1: Formula Parser (Foundation)

**Approach**: Tokenizer-based recursive descent parser

Convert formula string into tokens, then build expression tree:

```
Input: "=SUM(A1:A10)+B5*2"
  ↓
Tokenizer
  ↓
[FunctionToken(SUM), RangeToken(A1:A10), OperatorToken(+), 
 CellToken(B5), OperatorToken(*), NumberToken(2)]
  ↓
Parser (builds AST)
  ↓
[BinaryOp(
  left: FunctionCall(SUM, RangeRef(A1:A10)),
  op: +,
  right: BinaryOp(
    left: CellRef(B5),
    op: *,
    right: Number(2)
  )
)]
```

**Key Components**:

1. **FormulaTokenizer.swift**
   - Breaks formula string into tokens
   - Handles: operators, functions, cell/range references, literals, parentheses
   - Operator precedence: `^` > `*,/` > `+,-` > `&` > `=,<>,<,<=,>,>=`

2. **FormulaParser.swift**
   - Recursive descent parser converting tokens to AST
   - Handles operator precedence and associativity
   - Supports function calls with variable arguments

3. **FormulaAST.swift**
   - Abstract Syntax Tree node types
   - Expression, BinaryOp, UnaryOp, FunctionCall, CellRef, RangeRef, Literal

**Implementation Strategy**:
- Use Swift enums for AST nodes (type-safe)
- Implement as extension to CellReference parsing
- Minimal dependencies (no external libraries required)
- Pattern: Match Apache POI's approach of tokenizing first, then parsing

### Phase 2: Cell Dependency Resolution

**Problem**: Formulas reference other cells that may have formulas
**Solution**: Build directed graph, detect cycles, determine evaluation order

```
Cell A1: =5
Cell B1: =A1+1
Cell C1: =B1*2
Cell D1: =C1+A1

Dependency Graph:
A1 → B1 → C1 → D1
 ↘________↗

Topological Order: [A1, B1, C1, D1]
Circular Detection: If D1=A1+B1 AND A1=D1, circular reference!
```

**Key Components**:

1. **DependencyGraph.swift**
   - Extract cell references from AST
   - Build directed graph of cell → referenced cells
   - Implement DFS for cycle detection
   - Compute topological sort for evaluation order

2. **CellDependencyResolver.swift**
   - Given a sheet, compute all dependencies
   - Validate no circular references
   - Return evaluation order

### Phase 3: Formula Evaluator

**Problem**: Execute AST with actual cell values
**Solution**: Stack-based evaluator with function library

```
BinaryOp(
  left: CellRef(B5),
  op: *,
  right: Number(2)
)

Evaluation:
1. Evaluate left: CellRef(B5) → get value 10 from sheet
2. Evaluate right: Number(2) → 2
3. Apply operator: * → 10 * 2 = 20
4. Return 20
```

**Key Components**:

1. **FormulaEvaluator.swift**
   - Traverse AST and evaluate nodes
   - Handle errors: invalid cell ref, type mismatch, #DIV/0!, etc.
   - Support lazy evaluation for IF/AND/OR

2. **ExcelFunctionLibrary.swift**
   - Implement Excel functions
   - Return ExcelValue (can be number, string, boolean, error, or array)
   - Support variable arguments (SUM, etc.)

3. **ExcelValue.swift**
   - Result type: number, string, boolean, error, array
   - Type coercion rules (Excel's loose typing)
   - Error values: #DIV/0!, #N/A, #VALUE!, #REF!, #NAME?, #NUM!

### Phase 4: Integration with Existing Code

**Extend ExcelCell**:
```swift
public func evaluateFormula(in sheet: ExcelSheet) throws -> CellValue {
    guard hasFormula else { return value }
    let evaluator = FormulaEvaluator(sheet: sheet)
    return try evaluator.evaluate(formula!)
}
```

**Extend ExcelWorkbook**:
```swift
public func recalculateAllFormulas() throws {
    for sheet in sheets {
        try sheet.recalculateFormulas()
    }
}

public func recalculateDependentFormulas(afterModifying cell: ExcelCell) throws {
    // Only recalculate cells that depend on this cell
    let deps = dependencyGraph.dependentsOf(cell.reference)
    for dependent in deps {
        try recalculateCell(dependent)
    }
}
```

## Function Implementation Priority

### Tier 1: Essential (MVP - Week 1-2)
These functions are critical for basic Marbar template compatibility:

1. **SUM(range, [range2], ...)** - Add values
2. **AVERAGE(range, [range2], ...)** - Mean value
3. **COUNT(range, [range2], ...)** - Count numeric cells
4. **COUNTA(range, [range2], ...)** - Count non-empty cells
5. **IF(condition, true_value, false_value)** - Conditional logic
6. **Arithmetic Operators**: +, -, *, /, ^

### Tier 2: Important (Week 3-4)
Common in business templates:

7. **MIN(range, [range2], ...)** - Minimum value
8. **MAX(range, [range2], ...)** - Maximum value
9. **CONCATENATE(text1, text2, ...)** or **& operator** - Join strings
10. **AND(value1, value2, ...)** - Logical AND
11. **OR(value1, value2, ...)** - Logical OR
12. **NOT(value)** - Logical NOT
13. **VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])** - Vertical lookup
14. **INDEX(array, row, [column])** - Return value at position

### Tier 3: Common (Week 5-6)
Used in many reports:

15. **SUMIF(range, criteria, [sum_range])** - Conditional sum
16. **COUNTIF(range, criteria)** - Conditional count
17. **AVERAGEIF(range, criteria, [average_range])** - Conditional average
18. **MATCH(lookup_value, lookup_array, [match_type])** - Find position
19. **IFERROR(value, value_if_error)** - Error handling
20. **UPPER(text)** - Convert to uppercase
21. **LOWER(text)** - Convert to lowercase
22. **LEN(text)** - String length
23. **TRIM(text)** - Remove leading/trailing spaces

### Tier 4: Advanced (Week 7-8)
Specialized functions:

24. **SUMIFS(sum_range, criteria_range1, criteria1, ...)** - Multi-criteria sum
25. **COUNTIFS(criteria_range1, criteria1, ...)** - Multi-criteria count
26. **ROUND(number, num_digits)** - Round number
27. **ROUNDUP(number, num_digits)** - Round up
28. **ROUNDDOWN(number, num_digits)** - Round down
29. **ABS(number)** - Absolute value
30. **SQRT(number)** - Square root
31. **MEDIAN(number1, [number2], ...)** - Median value
32. **STDEV(number1, [number2], ...)** - Standard deviation
33. **LEFT(text, num_chars)** - Leftmost characters
34. **RIGHT(text, num_chars)** - Rightmost characters
35. **MID(text, start_num, num_chars)** - Middle characters
36. **FIND(find_text, within_text, [start_num])** - Find text position
37. **REPLACE(old_text, start_num, num_chars, new_text)** - Replace text
38. **TODAY()** - Current date
39. **NOW()** - Current date and time
40. **YEAR(serial_number)** - Extract year
41. **MONTH(serial_number)** - Extract month
42. **DAY(serial_number)** - Extract day

### Functions to Defer (Not Priority)
- HLOOKUP (use VLOOKUP pattern)
- INDIRECT (complex reference handling)
- OFFSET (complex reference handling)
- Complex matrix operations (MMULT, TRANSPOSE)
- Solver/Optimization functions
- Advanced statistical functions (FORECAST, TREND)

## Code Structure Proposal

```
ApachePOISwift/
├── Sources/ApachePOISwift/
│   └── Formulas/                          [NEW]
│       ├── FormulaTokenizer.swift        - Lexical analysis
│       ├── FormulaParser.swift           - Syntax analysis
│       ├── FormulaAST.swift              - Expression tree nodes
│       ├── FormulaEvaluator.swift        - Runtime evaluation
│       ├── ExcelFunctionLibrary.swift    - Built-in functions
│       ├── ExcelFunctionRegistry.swift   - Function management
│       ├── ExcelValue.swift              - Result wrapper
│       ├── DependencyGraph.swift         - Cell dependencies
│       └── CellDependencyResolver.swift  - Evaluation order
│
├── Tests/ApachePOISwiftTests/
│   └── FormulaTests/                     [EXPAND]
│       ├── FormulaTokenizerTests.swift
│       ├── FormulaParserTests.swift
│       ├── FormulaEvaluatorTests.swift
│       ├── ExcelFunctionTests.swift
│       └── DependencyGraphTests.swift
│
└── Examples/
    └── FormulaEvaluationExamples.swift   [EXPAND]
```

## Technical Deep Dives

### 1. Tokenizer Implementation

```swift
enum FormulaToken: Equatable {
    case number(Double)
    case string(String)
    case cellRef(String)        // "A1", "$B$2"
    case rangeRef(String)       // "A1:C10"
    case function(String)       // "SUM", "VLOOKUP"
    case `operator`(String)     // "+", "-", "*", "/", "^"
    case comparison(String)     // "=", "<>", "<", ">", "<=", ">="
    case boolean(Bool)          // TRUE, FALSE
    case leftParen
    case rightParen
    case leftBracket
    case rightBracket
    case comma
    case semicolon              // For array rows (varies by locale)
    case colon                  // For ranges
    case ampersand              // String concatenation
    case whitespace
    case error(String)
}

class FormulaTokenizer {
    func tokenize(_ formula: String) throws -> [FormulaToken] {
        // Handle formula string character by character
        // Track state: in string literal? in function name? etc.
        // Return array of tokens
    }
}
```

**Key Challenges**:
- Excel has different separators in different locales (`;` vs `,`)
- Quoted strings can contain special characters
- Cell references vs function names (both use letters)
- Operators can be unary or binary (e.g., `-5` vs `A1-B1`)

### 2. Operator Precedence (Per ECMA-376)

```
Precedence (highest to lowest):
1. :        Range (A1:C10)
2. (space)  Intersection
3. ,        Union (multiple ranges)
4. ^        Exponentiation
5. - + ~    Unary minus, plus, NOT (right-associative)
6. * / %    Multiplication, division, modulo
7. + -      Addition, subtraction
8. &        String concatenation
9. = <> < > <= >= Like  Comparison
10. NOT     Logical NOT
11. AND     Logical AND
12. OR      Logical OR
```

### 3. Cell Reference Parsing in Formulas

Formulas contain several reference types:

```swift
// Simple cell reference
=A1              // CellRef("A1")
=$A$1            // Absolute reference
=A$1             // Row absolute, col relative
=$A1             // Col absolute, row relative

// Range reference
=A1:A10          // RangeRef("A1", "A10")
=$A$1:$B$10      // Absolute range
=A1:B5           // Rectangle

// Sheet reference
=Sheet1!A1       // Cross-sheet
=Sheet1!A1:A10   // Cross-sheet range
='Sheet 1'!A1    // Quoted sheet name with spaces

// 3D reference (less common)
=Sheet1:Sheet3!A1  // Multiple sheets

// Named range (advanced)
=MyNamedRange    // User-defined named range
```

### 4. Circular Reference Detection

Use Depth-First Search (DFS) with color marking:

```swift
enum VertexColor {
    case white   // Unvisited
    case gray    // Currently visiting (in stack)
    case black   // Visited
}

func hasCycle(in dependencies: [String: Set<String>]) -> Bool {
    var colors = [String: VertexColor]()
    
    for node in dependencies.keys {
        if colors[node] == nil {
            if dfs(node, colors: &colors, dependencies: dependencies) {
                return true
            }
        }
    }
    return false
}

func dfs(_ node: String, 
         colors: inout [String: VertexColor], 
         dependencies: [String: Set<String>]) -> Bool {
    colors[node] = .gray
    
    for neighbor in dependencies[node] ?? [] {
        switch colors[neighbor] {
        case .gray:     return true      // Back edge = cycle!
        case .black:    break            // Already processed
        case .white, nil:
            if dfs(neighbor, colors: &colors, dependencies: dependencies) {
                return true
            }
        }
    }
    
    colors[node] = .black
    return false
}
```

### 5. Type Coercion Rules (Excel's Loose Typing)

Excel attempts type conversions automatically:

```swift
extension ExcelValue {
    // Convert to number for arithmetic operations
    var asNumber: Double? {
        switch self {
        case .number(let n):        return n
        case .string(let s):        return Double(s)  // "123" → 123, "abc" → nil
        case .boolean(let b):       return b ? 1 : 0
        case .date(let d):          return excelDateNumber(d)
        case .error:                return nil
        case .array:                return nil        // Arrays don't auto-convert
        }
    }
    
    // Convert to string for concatenation
    var asString: String {
        switch self {
        case .number(let n):        return String(n)
        case .string(let s):        return s
        case .boolean(let b):       return b ? "TRUE" : "FALSE"
        case .date(let d):          return dateFormatter.string(from: d)
        case .error(let e):         return e.description
        case .array:                return "#VALUE!"
        }
    }
    
    // Convert to boolean for logical operations
    var asBoolean: Bool {
        switch self {
        case .number(let n):        return n != 0
        case .string(let s):        return !s.isEmpty
        case .boolean(let b):       return b
        case .date:                 return true      // Dates are truthy
        case .error:                return false
        case .array:                return false
        }
    }
}
```

### 6. Array Formula Handling (Advanced)

Excel supports array formulas with Ctrl+Shift+Enter:

```swift
// Array formula: =SUM(IF(A1:A10>5, A1:A10, 0))
// Without array handling, this returns single value
// With array handling, processes each element

// For MVP, can store formula but mark as non-array
// For Phase 5, implement array operations:
struct ArrayValue {
    let rows: Int
    let cols: Int
    let data: [[ExcelValue]]
    
    subscript(row: Int, col: Int) -> ExcelValue {
        return data[row][col]
    }
}
```

## Integration with Marbar Template

The Marbar template (26 sheets, 61KB macros) will be the real-world test:

```swift
// Example: Generate Marbar report with formula evaluation
let workbook = try ExcelWorkbook(fileURL: marbarURL)
let sheet = try workbook.sheet(named: "GENERALES")

// Set input values
sheet.cell("B5").setValue(.string("PAD-123"))
sheet.cell("B10").setValue(.number(1500))

// Existing formulas in template automatically recalculate
// Example: B25 = =SUM(B10:B20) will show updated total

// Evaluate specific cell's formula
if let formulaResult = try sheet.cell("B25").evaluateFormula(in: sheet) {
    print("Result: \(formulaResult)")
}

// Or recalculate entire sheet (respects dependencies)
try sheet.recalculateAllFormulas()

// Save with both formulas AND cached results
try workbook.save(to: outputURL)
```

## Performance Considerations

### Caching Strategy

1. **Parse Cache**: Formulas parsed once, AST reused
2. **Evaluation Cache**: Results cached until inputs change
3. **Dependency Cache**: Graph computed once at load time

```swift
class FormulaEvaluationCache {
    private var evaluatedValues: [String: ExcelValue] = [:]
    private var parsedFormulas: [String: FormulaAST] = [:]
    
    func cachedEvaluate(_ cellRef: String, formula: String) -> ExcelValue? {
        if let cached = evaluatedValues[cellRef] {
            return cached
        }
        return nil
    }
    
    func invalidate(cellRef: String) {
        evaluatedValues[cellRef] = nil
        // Also invalidate dependents
    }
}
```

### Memory Management

- Large sheets (10k+ rows) → lazy evaluation
- Don't load all cell values into memory
- Stream evaluation through dependent cells
- Implement LRU cache for large datasets

## Testing Strategy

### Unit Tests (FormulaTokenizerTests.swift)
```swift
func testTokenizeSimpleFormula() {
    let tokens = try tokenizer.tokenize("=A1+B1")
    XCTAssertEqual(tokens, [
        .cellRef("A1"),
        .operator("+"),
        .cellRef("B1")
    ])
}

func testTokenizeFunctionCall() {
    let tokens = try tokenizer.tokenize("=SUM(A1:A10)")
    XCTAssertEqual(tokens, [
        .function("SUM"),
        .leftParen,
        .rangeRef("A1:A10"),
        .rightParen
    ])
}

func testOperatorPrecedence() {
    let tokens = try tokenizer.tokenize("=2+3*4")
    // Should parse as 2+(3*4), not (2+3)*4
}

func testComplexFormula() {
    let tokens = try tokenizer.tokenize(
        "=IF(SUM(A1:A10)>100,\"High\",\"Low\")"
    )
    // Verify all tokens including nested function
}
```

### Integration Tests (FormulaEvaluatorTests.swift)
```swift
func testEvaluateSimpleFormula() {
    let sheet = try createTestSheet(values: ["A1": 10, "B1": 20])
    let result = try evaluate("=A1+B1", in: sheet)
    XCTAssertEqual(result.asNumber, 30)
}

func testEvaluateSUMWithRange() {
    let sheet = try createTestSheet(values: [
        "A1": 10, "A2": 20, "A3": 30
    ])
    let result = try evaluate("=SUM(A1:A3)", in: sheet)
    XCTAssertEqual(result.asNumber, 60)
}

func testCircularReferenceDetection() {
    // A1 = =B1, B1 = =A1
    XCTAssertThrowsError(try evaluateSheet()) { error in
        XCTAssert(error is CircularReferenceError)
    }
}

func testDependencyEvaluationOrder() {
    let sheet = try createTestSheet(formulas: [
        "A1": "=5",
        "B1": "=A1+1",
        "C1": "=B1*2"
    ])
    try sheet.recalculateAllFormulas()
    
    XCTAssertEqual(sheet.cell("A1").value.asNumber, 5)
    XCTAssertEqual(sheet.cell("B1").value.asNumber, 6)
    XCTAssertEqual(sheet.cell("C1").value.asNumber, 12)
}
```

## Comparison: Apache POI vs Our Implementation

| Feature | Apache POI (Java) | ApachePOISwift | Approach |
|---------|-------------------|-----------------|----------|
| Parser | FormulaParser (recursive descent) | FormulaParser (same) | Match POI |
| Token Format | PTG tokens (Excel binary format) | FormulaToken (Swift enum) | Simplified for Swift |
| Evaluation | FormulaEvaluator + 202 functions | FormulaEvaluator + 40-50 functions | Phased implementation |
| Caching | WorkbookEvaluator maintains cache | FormulaEvaluationCache | Similar pattern |
| Circular Refs | Detected during evaluation | DependencyGraph pre-check | Proactive detection |
| Dependencies | Computed on-demand | Pre-computed topological sort | Optimized for Swift |

## Implementation Timeline

### Week 1: Foundation
- FormulaTokenizer (+ unit tests)
- FormulaAST types
- Start FormulaParser (basic operators)

### Week 2: Parser
- Complete FormulaParser (all operators + precedence)
- Cell/range reference parsing
- Function call parsing
- Unit tests covering complex formulas

### Week 3: Dependency Resolution
- DependencyGraph (extract references from AST)
- Circular reference detection
- Topological sort for evaluation order

### Week 4: Basic Evaluation
- FormulaEvaluator (traverse AST)
- Implement Tier 1 functions (SUM, COUNT, AVERAGE, IF, arithmetic)
- Error handling (#DIV/0!, #VALUE!, etc.)

### Week 5: More Functions
- Implement Tier 2 functions (MIN, MAX, AND, OR, VLOOKUP, INDEX)
- Array operations basics
- Update ExcelCell.evaluateFormula()

### Week 6: Integration & Marbar Testing
- Test with actual Marbar template
- Performance optimization
- Cache implementation

### Week 7: Advanced Functions
- Implement Tier 3 & 4 functions as needed
- Fine-tune based on Marbar requirements

## Reference Implementation Example: SUM Function

```swift
class SumFunction: ExcelFunction {
    let name = "SUM"
    
    func evaluate(args: [[ExcelValue]]) throws -> ExcelValue {
        var total: Double = 0
        
        for argArray in args {
            for value in argArray {
                switch value {
                case .number(let n):
                    total += n
                case .boolean(let b):
                    total += b ? 1 : 0
                case .string(let s):
                    // Strings in numeric context = 0 (or error)
                    if let n = Double(s) {
                        total += n
                    }
                case .empty:
                    // Empty cells ignored
                    break
                case .error(let e):
                    return .error(e)
                case .date(let d):
                    total += excelDateNumber(d)
                case .array:
                    throw FormulaError.arrayInvalidInContext
                }
            }
        }
        
        return .number(total)
    }
}
```

## Recommendations for MVP

1. **Start with Tokenizer**: Solid foundation, easiest to test
2. **Use Swift enums aggressively**: Type safety for AST nodes
3. **Don't implement array formulas initially**: Rare in most sheets
4. **Test against Marbar template early**: Real-world validation
5. **Cache formula parse results**: Biggest performance win
6. **Implement Tier 1 functions first**: 80/20 rule applies

## Known Limitations & Future Work

1. **Not Implementing**:
   - Array formulas (Ctrl+Shift+Enter)
   - User-defined functions (UDFs)
   - VBA macro execution (binary format too complex)
   - Goal Seek / Solver
   - Some Analysis ToolPak functions

2. **Simplified For Swift**:
   - Don't replicate Excel's full error handling (too complex)
   - Limited implicit type conversions
   - Simpler number format handling

3. **Performance Trade-offs**:
   - Recursive descent parser (not as efficient as hand-coded tokenizer)
   - AST interpretation (not compiled bytecode like Excel)
   - Acceptable for most use cases, not for massive spreadsheets (100k+ cells)

4. **Future Enhancements**:
   - Named ranges
   - INDIRECT() function
   - Array formulas
   - Custom function registration
   - Formula auditing tools
   - What-if analysis

## Conclusion

The proposed architecture follows Apache POI's proven approach while adapting to Swift's type system and performance characteristics. By implementing a traditional three-phase pipeline (tokenize → parse → evaluate) with explicit dependency resolution, we can achieve a solid Excel formula evaluator that:

- Handles the Marbar template's real-world needs
- Maintains compatibility with Excel's formula syntax
- Provides clear extension points for new functions
- Avoids circular reference pitfalls
- Performs well on typical spreadsheet sizes

The key success factors are:
1. Comprehensive tokenization (handle all Excel syntax)
2. Correct operator precedence (critical for correctness)
3. Robust dependency resolution (avoid silent bugs)
4. Extensible function library (easy to add new functions)
5. Early real-world testing (with Marbar template)

