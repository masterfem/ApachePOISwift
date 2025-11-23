# Excel Formula Evaluation Research - Document Index

## Overview

This directory contains comprehensive research on implementing Excel formula evaluation in Swift for the ApachePOISwift project.

**Research Date**: November 23, 2025  
**Status**: Complete and ready for implementation planning

---

## Document Guide

### 1. FORMULA_RESEARCH_SUMMARY.txt (START HERE)
**Purpose**: Executive summary and quick reference  
**Length**: ~400 lines  
**Best For**: Getting up to speed quickly, understanding the big picture

**Contains**:
- Current situation and what we have
- High-level recommended approach
- Required new files (9 files total)
- Critical challenges and solutions
- Function implementation priorities
- Integration points with existing code
- Performance targets
- Implementation timeline (7 weeks)
- Success criteria
- Quick start guide

**Start Here**: Read this first to understand the plan.

---

### 2. FORMULA_ARCHITECTURE_SUMMARY.txt
**Purpose**: Quick reference guide for architecture and design patterns  
**Length**: ~313 lines  
**Best For**: During implementation, quick lookup of architecture decisions

**Contains**:
- Current state (what we have vs. what's missing)
- Three-phase pipeline overview
- Function priority tiers with specific functions
- Key challenges and solutions
- Swift-specific design patterns
- Integration points with existing code
- Testing approach
- Performance targets
- Apache POI reference mapping
- Implementation timeline breakdown
- Success criteria
- Next steps

**Use During Implementation**: Keep this handy for quick reference.

---

### 3. FORMULA_EVALUATION_RESEARCH.md (DEEP DIVE)
**Purpose**: Comprehensive technical guide with implementation details  
**Length**: ~761 lines  
**Best For**: Understanding technical details, reference implementation, testing

**Contains**:
- Executive summary
- Current state in ApachePOISwift
- Architecture recommendations (4 phases)
- Function implementation priority (41 functions across 4 tiers)
- Proposed code structure
- Technical deep dives (6 topics)
- Integration with Marbar template
- Performance considerations
- Testing strategy (with code examples)
- Apache POI vs our implementation comparison
- Implementation timeline
- Reference implementation example (SUM function)
- Recommendations for MVP
- Known limitations and future work

**Read For Depth**: When you need technical details or reference code.

---

## Research Topics

### What was researched:

1. **Apache POI (Java)**
   - Formula parsing architecture
   - Token representation (PTG tokens)
   - FormulaEvaluator design
   - Function library structure (202+ functions)
   - Caching strategy
   - Circular reference handling

2. **ECMA-376 Specification**
   - Office Open XML (OOXML) format
   - Formula grammar (Section L.2.16)
   - Operator precedence rules
   - Cell reference formats
   - Type coercion rules

3. **Swift Expression Parsing**
   - nicklockwood/Expression library
   - Parser combinator approach
   - Tokenization techniques
   - Type handling in Swift

4. **Existing ApachePOISwift Code**
   - CellReference.swift (complete cell reference parsing)
   - ExcelCell.swift (formula storage)
   - Formula persistence and testing
   - Integration points

### What was found:

1. **No Single Perfect Solution**
   - Apache POI: Complex, designed for Java
   - Swift libraries: Don't handle Excel-specific syntax
   - Custom implementation: Best approach for our needs

2. **Three-Phase Pipeline is Proven**
   - Tokenize → Parse → Evaluate
   - Used by Apache POI
   - Used by commercial Excel implementations
   - Clear separation of concerns

3. **Key Challenges Identified**
   - Operator precedence (complex in Excel)
   - Circular reference detection (critical for correctness)
   - Cell reference resolution (multiple formats)
   - Type coercion (Excel's loose typing)
   - Performance on large sheets

4. **Swift is Well-Suited**
   - Enums for AST (type-safe)
   - Protocol-based functions (extensible)
   - Swift's error handling (proper exception model)
   - No external dependencies needed

---

## Quick Facts

### Architecture
- **Phase 1**: Tokenizer → Parser (produces AST)
- **Phase 2**: Dependency Graph (topological sort, cycle detection)
- **Phase 3**: Evaluator (traverse AST, call functions)

### Code to Create
- 9 new Swift files
- ~3,600 lines of code
- ~2,000 lines of tests

### Functions to Implement
- **MVP (Week 1-4)**: 6 essential functions
- **Phase 2 (Week 3-4)**: 8 important functions
- **Phase 3 (Week 5-6)**: 9 common functions
- **Phase 4 (Week 7-8)**: 18 advanced functions
- **Total**: 41 functions across 4 tiers

### Timeline
- **Week 1-2**: Foundation (tokenizer, AST, parser)
- **Week 3**: Dependency resolution
- **Week 4**: Basic evaluation + Tier 1 functions
- **Week 5-6**: Tier 2 functions + Marbar testing
- **Week 7**: Tier 3 & 4 functions as needed

### Success Metrics
- Parse simple formulas: ✓
- Evaluate basic functions: ✓
- Handle cell references: ✓
- Detect circular references: ✓
- Work with Marbar template: ✓
- >90% test coverage: ✓
- <500ms recalculation: ✓

---

## How to Use These Documents

### For Project Planning
→ Read: FORMULA_RESEARCH_SUMMARY.txt (entire file)

### For Architecture Understanding
→ Read: FORMULA_ARCHITECTURE_SUMMARY.txt (entire file)

### For Implementation Details
→ Read: FORMULA_EVALUATION_RESEARCH.md (full guide)

### During Implementation
→ Reference: FORMULA_ARCHITECTURE_SUMMARY.txt (quick lookup)

### For Testing
→ Reference: FORMULA_EVALUATION_RESEARCH.md (section: Testing Strategy)

### For Code Examples
→ Reference: FORMULA_EVALUATION_RESEARCH.md (section: Technical Deep Dives)

### For Function List
→ Reference: FORMULA_ARCHITECTURE_SUMMARY.txt (section: Function Priority)

---

## Next Steps

1. **Read all three documents** (in order)
   - FORMULA_RESEARCH_SUMMARY.txt
   - FORMULA_ARCHITECTURE_SUMMARY.txt
   - FORMULA_EVALUATION_RESEARCH.md

2. **Create directory structure**
   ```
   Sources/ApachePOISwift/Formulas/
   ├── FormulaTokenizer.swift
   ├── FormulaParser.swift
   ├── FormulaAST.swift
   ├── FormulaEvaluator.swift
   ├── ExcelFunctionLibrary.swift
   ├── ExcelFunctionRegistry.swift
   ├── ExcelValue.swift
   ├── DependencyGraph.swift
   └── CellDependencyResolver.swift
   ```

3. **Start implementation**
   - First: FormulaTokenizer.swift (foundation)
   - Then: FormulaAST.swift (data structures)
   - Then: FormulaParser.swift (syntax analysis)
   - Priority: Get tokenization and parsing working first

4. **Write tests early**
   - Comprehensive tokenizer tests
   - Parser tests with operator precedence
   - Real-world formula examples

5. **Test with Marbar template**
   - Open actual 26-sheet template
   - Verify formula compatibility
   - Ensure VBA macros still work

---

## Key Insights

### What Makes This Hard
1. Excel formula syntax is complex and inconsistent
2. Operator precedence must be exactly correct
3. Circular references must be detected
4. Type coercion rules are unintuitive

### What Makes This Possible
1. Clear three-phase architecture (proven by Apache POI)
2. Swift's type system (enum AST is elegant)
3. Well-defined ECMA-376 specification
4. Existing cell reference parser (reuse this)

### Critical Success Factors
1. **Start with tokenizer** - foundation for everything
2. **Correct operator precedence** - small bug = wrong results
3. **Robust dependency detection** - avoid circular reference bugs
4. **Phased function implementation** - 80/20 rule applies
5. **Test with real templates** - catch edge cases early

---

## Document Statistics

| Document | Lines | Purpose |
|----------|-------|---------|
| FORMULA_RESEARCH_SUMMARY.txt | ~400 | Executive summary |
| FORMULA_ARCHITECTURE_SUMMARY.txt | ~313 | Quick reference |
| FORMULA_EVALUATION_RESEARCH.md | ~761 | Technical details |
| **Total** | **~1,474** | Complete research |

---

## Research Sources

- Apache POI Documentation (official)
- ECMA-376 Office Open XML Specification
- Swift Expression Parser Libraries (nicklockwood/Expression)
- Existing ApachePOISwift Codebase
- Best Practices in Formula Evaluation

---

## Conclusion

These documents provide everything needed to implement Excel formula evaluation in Swift for ApachePOISwift.

The recommended three-phase architecture (Tokenize → Parse → Evaluate) is proven, well-documented, and achievable.

**Start with**: FORMULA_RESEARCH_SUMMARY.txt  
**Reference during implementation**: FORMULA_ARCHITECTURE_SUMMARY.txt  
**Go deeper when needed**: FORMULA_EVALUATION_RESEARCH.md

---

**Last Updated**: November 23, 2025  
**Status**: Ready for implementation
