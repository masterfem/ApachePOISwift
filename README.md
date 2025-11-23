# Apache POI for Swift

A pure Swift library for reading and writing Excel .xlsx/.xlsm files with VBA macro preservation, inspired by Apache POI (Java).

## 🎯 Project Status

**Phase 1**: ✅ **COMPLETE** - Foundation (Reading Excel files)
**Phase 2**: ✅ **COMPLETE** - Write Support (Modifying and saving)
**Phase 3**: ✅ **COMPLETE** - Macro Preservation
**Phase 4**: ✅ **COMPLETE** - Styles & Formatting
**Next Phase**: Phase 5 - Advanced Formulas
**License**: Apache 2.0

### What Works Now (Phases 1-4)
- ✅ Open .xlsx and .xlsm files
- ✅ Read cell values (strings, numbers, booleans, formulas, dates)
- ✅ **Modify cell values** (strings, numbers, booleans, formulas)
- ✅ **Save workbooks** to .xlsx/.xlsm files
- ✅ **Preserve VBA macros** during save operations
- ✅ **Read cell styles** (fonts, fills, borders, number formats)
- ✅ **Preserve cell styles** during save operations
- ✅ **Copy styles** between cells
- ✅ Access sheets by index or name
- ✅ Parse shared strings and inline strings
- ✅ Handle large files (tested with 3.2MB, 26-sheet workbooks with complex styling)
- ✅ **Full round-trip compatibility** (save → reload → verify)

## 📚 Documentation

- **[Examples/BasicReadExample.swift](./Examples/BasicReadExample.swift)** - 9 practical code examples
- **[Examples/StyleExample.swift](./Examples/StyleExample.swift)** - 6 style manipulation examples
- **[Integration Tests](./Tests/ApachePOISwiftTests/IntegrationTests.swift)** - Real-world usage examples
- **[Style Tests](./Tests/ApachePOISwiftTests/StyleTests.swift)** - Style reading and manipulation tests

### API Overview

```swift
// Open a workbook
let workbook = try ExcelWorkbook(fileURL: url)

// Access sheets
let sheet = try workbook.sheet(at: 0)           // By index
let sheet = try workbook.sheet(named: "Sales")  // By name

// Read cells
let cell = try sheet.cell("A1")                 // A1 notation
let cell = try sheet.cell(column: 0, row: 0)    // By indices

// Get cell values
switch cell.value {
case .string(let text): print(text)
case .number(let num): print(num)
case .boolean(let bool): print(bool)
case .formula(let formula): print(formula)
case .empty: print("Empty cell")
default: break
}

// Check for macros
if workbook.hasVBAMacros {
    print("This file contains VBA macros")
}

// Modify cells (Phase 2)
cell.setValue(.string("Updated Value"))
cell.setValue(.number(42.5))
cell.setValue(.boolean(true))
cell.setValue(.formula("=SUM(A1:A10)"))

// Read styles (Phase 4)
if let font = cell.font {
    print("Font: \(font.name ?? "unknown") \(font.size ?? 0)pt")
    print("Bold: \(font.bold), Italic: \(font.italic)")
}

if let fill = cell.fill {
    print("Background: \(fill.patternType.rawValue)")
    if let color = fill.foregroundColor {
        print("Color: \(color)")
    }
}

// Copy style to another cell
let targetCell = try sheet.cell("B2")
targetCell.setStyleIndex(cell.styleIndex)

// Save workbook (styles are automatically preserved)
try workbook.save(to: outputURL)
```

## 🚀 Quick Start

### Installation

Add to your `Package.swift`:

```swift
dependencies: [
    .package(url: "https://github.com/masterfem/ApachePOISwift.git", branch: "main")
]
```

### Basic Usage

```swift
import ApachePOISwift

// Open an Excel file
let workbook = try ExcelWorkbook(fileURL: fileURL)

// Read data
let sheet = try workbook.sheet(at: 0)
let cell = try sheet.cell("A1")
print(cell.value)  // Prints the cell value

// Iterate through sheets
for sheet in workbook.allSheets {
    print("Sheet: \(sheet.name)")
    for cell in sheet.nonEmptyCells() {
        print("  \(cell.reference): \(cell.value)")
    }
}
```

### Write Support (Phase 2 - Complete!)

```swift
// Modify data
let workbook = try ExcelWorkbook(fileURL: templateURL)
let sheet = try workbook.sheet(named: "Sales")

// Set cell values
try sheet.cell("A1").setValue(.string("Updated Text"))
try sheet.cell("B1").setValue(.number(123.45))
try sheet.cell("C1").setValue(.formula("=SUM(B1:B10)"))

// Save with macros preserved
try workbook.save(to: outputURL)

// Macros are intact! ✅
let reloaded = try ExcelWorkbook(fileURL: outputURL)
print("Has macros: \(reloaded.hasVBAMacros)")  // true
```

## 🏗️ Why This Library?

**No existing Swift library** supports:
- Full .xlsx/.xlsm read/write
- VBA macro preservation
- Complex Excel features (formulas, charts, styles)

**This library fills that gap** by:
- Treating .xlsx/.xlsm as ZIP + XML (standard format)
- Preserving vbaProject.bin untouched (macros intact)
- Following Apache POI's proven architecture
- Pure Swift (iOS/macOS/watchOS compatible)

## 🧪 Testing

Run the test suite:

```bash
swift test
```

Current test coverage:
- ✅ **35 tests passing** (all green!)
- ✅ Unit tests for cell reference parsing
- ✅ Integration tests with real 3.2MB Excel file (26 sheets, VBA macros, complex styling)
- ✅ Write tests (modify, save, reload, verify)
- ✅ Macro preservation tests
- ✅ Style reading and manipulation tests
- ✅ Inline string round-trip tests
- ✅ Error handling tests

## 🗺️ Roadmap

- [x] **Phase 1: Foundation** - Read .xlsx/.xlsm files ✅ **COMPLETE**
- [x] **Phase 2: Write Support** - Modify cells and save files ✅ **COMPLETE**
- [x] **Phase 3: Macro Preservation** - Save .xlsm with VBA intact ✅ **COMPLETE**
- [x] **Phase 4: Styles & Formatting** - Read/preserve fonts, colors, borders ✅ **COMPLETE**
- [ ] **Phase 5: Advanced Formulas** - Write formulas, optional evaluation (Next)
- [ ] **Phase 6: Advanced Features** - Charts, conditional formatting, data validation

## 🤝 Contributing

Contributions welcome! This project follows standard Swift Package Manager conventions.

### Development

```bash
git clone https://github.com/masterfem/ApachePOISwift.git
cd ApachePOISwift
swift build
swift test
```

## 📄 License

Apache License 2.0 (same as Apache POI for compatibility)

## 🙏 Acknowledgments

Inspired by [Apache POI](https://poi.apache.org/) - the industry-standard Java library for Microsoft Office formats.

---

**Created**: November 22, 2024
**Status**: Phases 1-4 Complete (Read/Write/Macros/Styles) - Phase 5 Next (Advanced Formulas)
