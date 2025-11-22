# Apache POI for Swift

A pure Swift library for reading and writing Excel .xlsx/.xlsm files with VBA macro preservation, inspired by Apache POI (Java).

## 🎯 Project Status

**Current Phase**: Architecture & Planning Complete
**Implementation**: Not yet started (ready for development)
**License**: Apache 2.0

## 📚 Documentation

See **[CLAUDE.md](./CLAUDE.md)** for comprehensive implementation guide including:
- Excel file format deep dive (ECMA-376, ZIP + XML)
- Apache POI architecture mapping (Java → Swift)
- 6-phase implementation roadmap (12 weeks)
- Technical specifications and code examples
- Integration with Solids Control app

## 🚀 Quick Start (Future)

```swift
import ApachePOISwift

// Open existing file
let workbook = try ExcelWorkbook(fileURL: templateURL)

// Modify data
let sheet = try workbook.sheet(named: "Sales")
sheet.cell("A1").value = "Updated"
sheet.cell("B1").value = 123.45

// Save with macros preserved
try workbook.save(to: outputURL)
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

## 📦 Installation (Future)

```swift
// Package.swift
dependencies: [
    .package(url: "https://github.com/masterfem/ApachePOISwift.git", from: "1.0.0")
]
```

## 🤝 Contributing

This project is in active development. See CLAUDE.md for implementation roadmap.

## 📄 License

Apache License 2.0 (same as Apache POI for compatibility)

---

**Project Base**: `/Users/masterfem/ApachePOISwift`
**Parent Project**: Solids Control App (`/Users/masterfem/SolidsControlNqn`)
**Created**: November 22, 2024
