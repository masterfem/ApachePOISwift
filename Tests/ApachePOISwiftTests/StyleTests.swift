//
// StyleTests.swift
// ApachePOISwiftTests
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class StyleTests: XCTestCase {
    var marbarFileURL: URL!

    override func setUp() {
        super.setUp()

        // Get test file from test bundle
        let bundle = Bundle.module
        guard let url = bundle.url(forResource: "marbar_template", withExtension: "xlsm", subdirectory: "TestResources") else {
            XCTFail("Could not find marbar_template.xlsm in test resources")
            return
        }
        marbarFileURL = url
    }

    // MARK: - Style Parsing Tests

    func testLoadStyles() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Verify styles were loaded
        XCTAssertNotNil(workbook.stylesData, "Styles data should be loaded")

        guard let styles = workbook.stylesData else {
            XCTFail("Styles data is nil")
            return
        }

        // Verify fonts were parsed
        XCTAssertGreaterThan(styles.fonts.count, 0, "Should have parsed fonts")
        print("Loaded \(styles.fonts.count) fonts")

        // Verify fills were parsed
        XCTAssertGreaterThan(styles.fills.count, 0, "Should have parsed fills")
        print("Loaded \(styles.fills.count) fills")

        // Verify borders were parsed
        XCTAssertGreaterThan(styles.borders.count, 0, "Should have parsed borders")
        print("Loaded \(styles.borders.count) borders")

        // Verify cell styles were parsed
        XCTAssertGreaterThan(styles.cellStyles.count, 0, "Should have parsed cell styles")
        print("Loaded \(styles.cellStyles.count) cell styles")
    }

    func testFontParsing() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        guard let styles = workbook.stylesData else {
            XCTFail("Styles data is nil")
            return
        }

        // Test first font (usually default font)
        XCTAssertGreaterThan(styles.fonts.count, 0, "Should have at least one font")

        let firstFont = styles.fonts[0]

        // Fonts should have basic properties
        print("First font: name=\(firstFont.name ?? "nil"), size=\(firstFont.size ?? 0), bold=\(firstFont.bold), italic=\(firstFont.italic)")

        // Common fonts in Excel templates
        let fontNames = styles.fonts.compactMap { $0.name }
        XCTAssertGreaterThan(fontNames.count, 0, "Should have fonts with names")
    }

    func testFillParsing() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        guard let styles = workbook.stylesData else {
            XCTFail("Styles data is nil")
            return
        }

        XCTAssertGreaterThan(styles.fills.count, 0, "Should have fills")

        let firstFill = styles.fills[0]
        print("First fill: pattern=\(firstFill.patternType.rawValue), fg=\(firstFill.foregroundColor ?? "nil"), bg=\(firstFill.backgroundColor ?? "nil")")

        // Count solid fills
        let solidFills = styles.fills.filter { $0.patternType == .solid }
        print("Found \(solidFills.count) solid fills")
    }

    func testBorderParsing() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        guard let styles = workbook.stylesData else {
            XCTFail("Styles data is nil")
            return
        }

        XCTAssertGreaterThan(styles.borders.count, 0, "Should have borders")

        // Find a border with actual border edges
        let bordersWithEdges = styles.borders.filter { border in
            border.left != nil || border.right != nil || border.top != nil || border.bottom != nil
        }

        print("Found \(bordersWithEdges.count) borders with edges")

        if let firstBorderWithEdge = bordersWithEdges.first {
            if let left = firstBorderWithEdge.left {
                print("Left border: style=\(left.style.rawValue), color=\(left.color ?? "nil")")
            }
        }
    }

    // MARK: - Cell Style Access Tests

    func testCellStyleAccess() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)  // First sheet

        // Find a cell with a style
        var foundStyledCell = false
        for row in 0..<20 {
            for col in 0..<10 {
                if let cell = try? sheet.cell(column: col, row: row),
                   let styleIndex = cell.styleIndex {
                    print("Cell \(cell.reference) has style index \(styleIndex)")

                    // Access style components
                    if let style = cell.style {
                        print("  Style: fontId=\(style.fontId ?? -1), fillId=\(style.fillId ?? -1), borderId=\(style.borderId ?? -1)")
                    }

                    if let font = cell.font {
                        print("  Font: \(font.name ?? "unknown") \(font.size ?? 0)pt, bold=\(font.bold), italic=\(font.italic)")
                    }

                    if let fill = cell.fill {
                        print("  Fill: \(fill.patternType.rawValue), color=\(fill.foregroundColor ?? "none")")
                    }

                    if let border = cell.border {
                        let edges = [
                            ("left", border.left),
                            ("right", border.right),
                            ("top", border.top),
                            ("bottom", border.bottom)
                        ].compactMap { name, edge in
                            edge != nil ? "\(name)=\(edge!.style.rawValue)" : nil
                        }
                        if !edges.isEmpty {
                            print("  Border: \(edges.joined(separator: ", "))")
                        }
                    }

                    foundStyledCell = true
                    break
                }
            }
            if foundStyledCell { break }
        }

        XCTAssertTrue(foundStyledCell, "Should find at least one cell with a style")
    }

    func testStylePreservation() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        print("Test output directory: \(outputDirectory.path)")

        // Load workbook
        let workbook1 = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet1 = try workbook1.sheet(at: 0)

        // Find a styled cell
        var testCell: ExcelCell?
        var originalStyleIndex: Int?

        for row in 0..<20 {
            for col in 0..<10 {
                if let cell = try? sheet1.cell(column: col, row: row),
                   cell.styleIndex != nil {
                    testCell = cell
                    originalStyleIndex = cell.styleIndex
                    break
                }
            }
            if testCell != nil { break }
        }

        guard let cell = testCell, let originalStyle = originalStyleIndex else {
            XCTFail("Could not find a styled cell for testing")
            return
        }

        print("Testing with cell \(cell.reference), original style index: \(originalStyle)")

        // Modify cell value (style should be preserved)
        try cell.setValue(.string("Test Value"))

        // Verify style is still there
        XCTAssertEqual(cell.styleIndex, originalStyle, "Style should be preserved after value change")

        // Save
        let outputURL = outputDirectory.appendingPathComponent("styled_output.xlsm")
        try workbook1.save(to: outputURL)

        // Reload
        let workbook2 = try ExcelWorkbook(fileURL: outputURL)
        let sheet2 = try workbook2.sheet(at: 0)
        let reloadedCell = try sheet2.cell(cell.reference)

        // Verify style was preserved through save/reload
        XCTAssertEqual(reloadedCell.styleIndex, originalStyle, "Style index should be preserved after save/reload")

        // Verify value was updated
        if case .string(let value) = reloadedCell.value {
            XCTAssertEqual(value, "Test Value", "Cell value should be updated")
        } else {
            XCTFail("Cell value should be a string")
        }
    }

    func testSetStyleIndex() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook1 = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet1 = try workbook1.sheet(at: 0)

        // Get a cell
        let cell = try sheet1.cell("A1")
        let originalStyleIndex = cell.styleIndex

        print("Original style index: \(originalStyleIndex ?? -1)")

        // Set new style index
        cell.setStyleIndex(5)

        XCTAssertEqual(cell.styleIndex, 5, "Style index should be updated")

        // Save
        let outputURL = outputDirectory.appendingPathComponent("modified_style.xlsm")
        try workbook1.save(to: outputURL)

        // Reload and verify
        let workbook2 = try ExcelWorkbook(fileURL: outputURL)
        let sheet2 = try workbook2.sheet(at: 0)
        let reloadedCell = try sheet2.cell("A1")

        XCTAssertEqual(reloadedCell.styleIndex, 5, "Modified style index should be preserved")
    }

    func testNumberFormat() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        guard let styles = workbook.stylesData else {
            XCTFail("Styles data is nil")
            return
        }

        // Test built-in number formats
        let generalFormat = NumberFormat(formatId: 0)
        XCTAssertEqual(generalFormat.getFormatCode(), "General")

        let percentFormat = NumberFormat(formatId: 9)
        XCTAssertEqual(percentFormat.getFormatCode(), "0%")

        let dateFormat = NumberFormat(formatId: 14)
        XCTAssertEqual(dateFormat.getFormatCode(), "mm-dd-yy")

        // Check if template has custom number formats
        if !styles.numberFormats.isEmpty {
            print("Template has \(styles.numberFormats.count) custom number formats")
            for format in styles.numberFormats.prefix(5) {
                print("  Format \(format.formatId): \(format.formatCode ?? "nil")")
            }
        }
    }
}
