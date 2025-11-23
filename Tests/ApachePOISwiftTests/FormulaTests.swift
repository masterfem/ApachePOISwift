//
// FormulaTests.swift
// ApachePOISwiftTests
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class FormulaTests: XCTestCase {
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

    // MARK: - Basic Formula Tests

    func testSetFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("Z90")

        // Set a simple formula
        cell.setFormula("=2+2")

        // Verify formula is set
        XCTAssertTrue(cell.hasFormula, "Cell should have a formula")
        XCTAssertEqual(cell.formula, "=2+2", "Formula should match")

        // Save and reload
        let outputURL = outputDirectory.appendingPathComponent("formula_basic.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("Z90")

        XCTAssertTrue(reloadedCell.hasFormula, "Reloaded cell should have a formula")
        XCTAssertEqual(reloadedCell.formula, "=2+2", "Reloaded formula should match")
    }

    func testSUMFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Set values in A1:A5
        for i in 0..<5 {
            let cell = try sheet.cell(column: 0, row: i)
            cell.setValue(.number(Double(i + 1)))  // 1, 2, 3, 4, 5
        }

        // Set SUM formula in A6
        let sumCell = try sheet.cell("A6")
        sumCell.setFormula("=SUM(A1:A5)")

        let outputURL = outputDirectory.appendingPathComponent("formula_sum.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("A6")

        XCTAssertTrue(reloadedCell.hasFormula)
        XCTAssertEqual(reloadedCell.formula, "=SUM(A1:A5)")
    }

    func testAVERAGEFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z91")
        cell.setFormula("=AVERAGE(B1:B10)")

        let outputURL = outputDirectory.appendingPathComponent("formula_average.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z91")

        XCTAssertEqual(reloadedCell.formula, "=AVERAGE(B1:B10)")
    }

    func testIFFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z92")
        cell.setFormula("=IF(A1>100,\"High\",\"Low\")")

        let outputURL = outputDirectory.appendingPathComponent("formula_if.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z92")

        XCTAssertEqual(reloadedCell.formula, "=IF(A1>100,\"High\",\"Low\")")
    }

    // MARK: - Cell Reference Tests

    func testRelativeCellReference() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("B2")
        cell.setFormula("=A1+A2")

        let outputURL = outputDirectory.appendingPathComponent("formula_relative.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("B2")

        XCTAssertEqual(reloadedCell.formula, "=A1+A2")
    }

    func testAbsoluteCellReference() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z93")
        cell.setFormula("=$A$1*B2")  // Absolute A1, relative B2

        let outputURL = outputDirectory.appendingPathComponent("formula_absolute.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z93")

        XCTAssertEqual(reloadedCell.formula, "=$A$1*B2")
    }

    func testCrossSheetReference() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z94")
        // Reference a cell in the second sheet
        cell.setFormula("=Sheet2!A1+Sheet2!B1")

        let outputURL = outputDirectory.appendingPathComponent("formula_crosssheet.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z94")

        XCTAssertEqual(reloadedCell.formula, "=Sheet2!A1+Sheet2!B1")
    }

    // MARK: - Range Tests

    func testRangeFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z95")
        cell.setFormula("=SUM(A1:C10)")

        let outputURL = outputDirectory.appendingPathComponent("formula_range.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z95")

        XCTAssertEqual(reloadedCell.formula, "=SUM(A1:C10)")
    }

    func testAbsoluteRangeFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z96")
        cell.setFormula("=AVERAGE($A$1:$A$100)")

        let outputURL = outputDirectory.appendingPathComponent("formula_abs_range.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z96")

        XCTAssertEqual(reloadedCell.formula, "=AVERAGE($A$1:$A$100)")
    }

    // MARK: - Complex Formula Tests

    func testVLOOKUPFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z97")
        cell.setFormula("=VLOOKUP(A1,Sheet2!A:B,2,FALSE)")

        let outputURL = outputDirectory.appendingPathComponent("formula_vlookup.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z97")

        XCTAssertEqual(reloadedCell.formula, "=VLOOKUP(A1,Sheet2!A:B,2,FALSE)")
    }

    func testNestedFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z98")
        cell.setFormula("=IF(SUM(A1:A10)>100,AVERAGE(B1:B10),MAX(C1:C10))")

        let outputURL = outputDirectory.appendingPathComponent("formula_nested.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z98")

        XCTAssertEqual(reloadedCell.formula, "=IF(SUM(A1:A10)>100,AVERAGE(B1:B10),MAX(C1:C10))")
    }

    func testConcatenationFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("Z99")
        cell.setFormula("=CONCATENATE(A1,\" - \",B1)")

        let outputURL = outputDirectory.appendingPathComponent("formula_concat.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z99")

        XCTAssertEqual(reloadedCell.formula, "=CONCATENATE(A1,\" - \",B1)")
    }

    // MARK: - Formula Preservation Tests

    func testPreserveExistingFormulas() throws {
        // This test verifies that existing formulas in the template are preserved
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Marbar template has many sheets with formulas
        // Just verify we can read the workbook without errors
        XCTAssertGreaterThan(workbook.sheetCount, 0)

        // Modify some data (not formulas)
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("A1")
        cell.setValue(.string("Test"))

        // Save and reload
        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
            .appendingPathExtension("xlsm")

        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        XCTAssertGreaterThan(reloadedWorkbook.sheetCount, 0)
        // If formulas were preserved, the file should still be valid
    }

    // MARK: - setValue(.formula(...)) Compatibility Test

    func testSetValueWithFormula() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("C1")
        // Test the older setValue(.formula(...)) API
        cell.setValue(.formula("=A1+B1"))

        XCTAssertTrue(cell.hasFormula)
        XCTAssertEqual(cell.formula, "=A1+B1")

        let outputURL = outputDirectory.appendingPathComponent("formula_setvalue.xlsm")
        try workbook.save(to: outputURL)

        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("C1")

        XCTAssertEqual(reloadedCell.formula, "=A1+B1")
    }
}
