//
// StyleCreationTests.swift
// ApachePOISwiftTests
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class StyleCreationTests: XCTestCase {
    var marbarFileURL: URL!

    override func setUp() {
        super.setUp()

        let bundle = Bundle.module
        guard let url = bundle.url(forResource: "marbar_template", withExtension: "xlsm", subdirectory: "TestResources") else {
            XCTFail("Could not find marbar_template.xlsm in test resources")
            return
        }
        marbarFileURL = url
    }

    // MARK: - Font Creation Tests

    func testApplyFont() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z99")  // Use empty cell to avoid template styles

        // Apply custom font
        let customFont = Font(name: "Arial", size: 14, bold: true, italic: true, color: "FFFF0000")
        cell.applyFont(customFont)

        // Save
        let outputURL = outputDirectory.appendingPathComponent("font_test.xlsm")
        try workbook.save(to: outputURL)

        // Reload and verify
        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z99")

        XCTAssertNotNil(reloadedCell.font, "Cell should have a font")
        if let font = reloadedCell.font {
            XCTAssertEqual(font.name, "Arial")
            XCTAssertEqual(font.size, 14)
            XCTAssertTrue(font.bold)
            XCTAssertTrue(font.italic)
        }
    }

    func testMakeBold() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z91")

        cell.setValue(.string("Bold Text"))
        cell.makeBold()

        let outputURL = outputDirectory.appendingPathComponent("bold_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z91")

        XCTAssertNotNil(reloadedCell.font)
        XCTAssertTrue(reloadedCell.font?.bold ?? false, "Cell should be bold")
    }

    func testMakeItalic() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z92")

        cell.setValue(.string("Italic Text"))
        cell.makeItalic()

        let outputURL = outputDirectory.appendingPathComponent("italic_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z92")

        XCTAssertNotNil(reloadedCell.font)
        XCTAssertTrue(reloadedCell.font?.italic ?? false, "Cell should be italic")
    }

    // MARK: - Fill Creation Tests

    func testSetBackgroundColor() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z93")

        cell.setValue(.string("Red Background"))
        cell.setBackgroundColor("FFFF0000") // Red

        let outputURL = outputDirectory.appendingPathComponent("background_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z93")

        XCTAssertNotNil(reloadedCell.fill)
        XCTAssertEqual(reloadedCell.fill?.patternType, .solid)
        XCTAssertEqual(reloadedCell.fill?.foregroundColor, "FFFF0000")
    }

    func testApplyFill() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z94")

        cell.applyFill(Fill(patternType: .solid, foregroundColor: "FF00FF00")) // Green

        let outputURL = outputDirectory.appendingPathComponent("fill_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z94")

        XCTAssertNotNil(reloadedCell.fill)
        XCTAssertEqual(reloadedCell.fill?.foregroundColor, "FF00FF00")
    }

    // MARK: - Border Creation Tests

    func testSetBorder() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z95")

        cell.setValue(.string("Bordered Cell"))
        cell.setBorder(style: .medium, color: "FF000000")

        let outputURL = outputDirectory.appendingPathComponent("border_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z95")

        XCTAssertNotNil(reloadedCell.border)
        XCTAssertEqual(reloadedCell.border?.left?.style, .medium)
        XCTAssertEqual(reloadedCell.border?.right?.style, .medium)
        XCTAssertEqual(reloadedCell.border?.top?.style, .medium)
        XCTAssertEqual(reloadedCell.border?.bottom?.style, .medium)
    }

    func testApplyBorder() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z96")

        let customBorder = Border(
            left: BorderEdge(style: .thin, color: "FF000000"),
            right: BorderEdge(style: .thick, color: "FFFF0000"),
            top: BorderEdge(style: .double, color: "FF0000FF"),
            bottom: BorderEdge(style: .thin, color: "FF00FF00")
        )
        cell.applyBorder(customBorder)

        let outputURL = outputDirectory.appendingPathComponent("custom_border_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z96")

        XCTAssertNotNil(reloadedCell.border)
        XCTAssertEqual(reloadedCell.border?.left?.style, .thin)
        XCTAssertEqual(reloadedCell.border?.right?.style, .thick)
        XCTAssertEqual(reloadedCell.border?.top?.style, .double)
        XCTAssertEqual(reloadedCell.border?.bottom?.style, .thin)
    }

    // MARK: - Combined Style Tests

    func testApplyCompleteStyle() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z97")

        cell.setValue(.string("Fully Styled"))
        cell.applyStyle(
            font: Font(name: "Times New Roman", size: 16, bold: true, color: "FFFFFFFF"),
            fill: Fill(patternType: .solid, foregroundColor: "FF0000FF"),
            border: Border(
                left: BorderEdge(style: .medium),
                right: BorderEdge(style: .medium),
                top: BorderEdge(style: .medium),
                bottom: BorderEdge(style: .medium)
            ),
            horizontalAlignment: .center,
            verticalAlignment: .center,
            wrapText: true
        )

        let outputURL = outputDirectory.appendingPathComponent("complete_style_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z97")

        // Verify font
        XCTAssertNotNil(reloadedCell.font)
        XCTAssertEqual(reloadedCell.font?.name, "Times New Roman")
        XCTAssertEqual(reloadedCell.font?.size, 16)
        XCTAssertTrue(reloadedCell.font?.bold ?? false)

        // Verify fill
        XCTAssertNotNil(reloadedCell.fill)
        XCTAssertEqual(reloadedCell.fill?.patternType, .solid)
        XCTAssertEqual(reloadedCell.fill?.foregroundColor, "FF0000FF")

        // Verify border
        XCTAssertNotNil(reloadedCell.border)
        XCTAssertEqual(reloadedCell.border?.left?.style, .medium)

        // Verify alignment
        XCTAssertNotNil(reloadedCell.style)
        XCTAssertEqual(reloadedCell.style?.horizontalAlignment, .center)
        XCTAssertEqual(reloadedCell.style?.verticalAlignment, .center)
        XCTAssertTrue(reloadedCell.style?.wrapText ?? false)
    }

    func testMultipleCellsStyling() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Style a range of cells
        for col in 0..<5 {
            let cell = try sheet.cell(column: col, row: 0)
            cell.setValue(.string("Header \(col + 1)"))
            cell.makeBold()
            cell.setBackgroundColor("FF4472C4") // Blue
        }

        let outputURL = outputDirectory.appendingPathComponent("multiple_cells_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)

        // Verify all cells are styled
        for col in 0..<5 {
            let cell = try reloadedSheet.cell(column: col, row: 0)
            XCTAssertTrue(cell.font?.bold ?? false, "Cell at column \(col) should be bold")
            XCTAssertEqual(cell.fill?.foregroundColor, "FF4472C4", "Cell at column \(col) should have blue background")
        }
    }

    func testNumberFormat() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z98")

        cell.setValue(.number(1234.5678))
        cell.applyNumberFormat(NumberFormat(formatId: 2)) // 0.00 format

        let outputURL = outputDirectory.appendingPathComponent("number_format_test.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z98")

        XCTAssertNotNil(reloadedCell.numberFormat)
        XCTAssertEqual(reloadedCell.numberFormat?.formatId, 2)
    }
}
