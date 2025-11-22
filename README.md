# Apache POI for Swift

A pure Swift library for reading and writing Excel .xlsx/.xlsm files with VBA macro preservation, inspired by Apache POI (Java).

## 🎯 Project Status

**Phase 1**: ✅ **COMPLETE** - Foundation (Reading Excel files)
**Next Phase**: Phase 2 - Write Support
**License**: Apache 2.0

### What Works Now (Phase 1)
- ✅ Open .xlsx and .xlsm files
- ✅ Read cell values (strings, numbers, booleans, formulas)
- ✅ Access sheets by index or name
- ✅ Detect VBA macros
- ✅ Parse shared strings
- ✅ Handle large files (tested with 3.2MB, 26-sheet workbooks)

## 📚 Documentation

- **[Examples/BasicReadExample.swift](./Examples/BasicReadExample.swift)** - 9 practical code examples
- **[Integration Tests](./Tests/ApachePOISwiftTests/IntegrationTests.swift)** - Real-world usage examples

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

### Write Support (Coming Soon - Phase 2)

```swift
// Future: Modify data
let sheet = try workbook.sheet(named: "Sales")
sheet.cell("A1").value = "Updated"  // Phase 2
sheet.cell("B1").value = 123.45     // Phase 2

// Save with macros preserved
try workbook.save(to: outputURL)     // Phase 2
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
- ✅ 17 tests passing
- ✅ Unit tests for cell reference parsing
- ✅ Integration tests with real 3.2MB Excel file (26 sheets, VBA macros)
- ✅ Error handling tests

## 🗺️ Roadmap

- [x] **Phase 1: Foundation** - Read .xlsx/.xlsm files ✅ **COMPLETE**
- [ ] **Phase 2: Write Support** - Modify cells and save files (Next)
- [ ] **Phase 3: Macro Preservation** - Save .xlsm with VBA intact
- [ ] **Phase 4: Styles & Formatting** - Fonts, colors, borders
- [ ] **Phase 5: Formulas** - Write formulas, optional evaluation
- [ ] **Phase 6: Advanced Features** - Charts, conditional formatting

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
**Status**: Phase 1 Complete (Reading) - Phase 2 In Progress (Writing)
