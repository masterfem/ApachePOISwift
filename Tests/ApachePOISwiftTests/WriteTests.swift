//
// WriteTests.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class WriteTests: XCTestCase {
    var marbarFileURL: URL!
    var outputDirectory: URL!

    override func setUpWithError() throws {
        let testBundle = Bundle.module
        marbarFileURL = testBundle.url(
            forResource: "marbar_template",
            withExtension: "xlsm",
            subdirectory: "TestResources"
        )

        XCTAssertNotNil(marbarFileURL, "marbar_template.xlsm should exist in TestResources")

        // Create temporary output directory
        outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)
    }

    override func tearDownWithError() throws {
        // Don't clean up output directory so we can inspect files
        // if FileManager.default.fileExists(atPath: outputDirectory.path) {
        //     try? FileManager.default.removeItem(at: outputDirectory)
        // }
        print("Test output directory: \(outputDirectory.path)")
    }

    func testSetStringValue() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("A1")
        cell.setValue(.string("Test String"))

        // Verify value was set
        if case .string(let value) = cell.value {
            XCTAssertEqual(value, "Test String")
        } else {
            XCTFail("Expected string value")
        }
    }

    func testSetNumberValue() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("B1")
        cell.setValue(.number(123.45))

        // Verify value was set
        if case .number(let value) = cell.value {
            XCTAssertEqual(value, 123.45, accuracy: 0.001)
        } else {
            XCTFail("Expected number value")
        }
    }

    func testSetBooleanValue() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("C1")
        cell.setValue(.boolean(true))

        // Verify value was set
        if case .boolean(let value) = cell.value {
            XCTAssertTrue(value)
        } else {
            XCTFail("Expected boolean value")
        }
    }

    func testSetFormulaValue() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        let cell = try sheet.cell("D1")
        cell.setValue(.formula("=SUM(A1:A10)"))

        // Verify formula was set
        if case .formula(let formula) = cell.value {
            XCTAssertEqual(formula, "=SUM(A1:A10)")
        } else {
            XCTFail("Expected formula value")
        }
    }

    func testModifyMultipleCells() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Modify multiple cells
        try sheet.cell("A1").setValue(.string("Cell A1"))
        try sheet.cell("A2").setValue(.number(100))
        try sheet.cell("A3").setValue(.boolean(false))
        try sheet.cell("B1").setValue(.formula("=A2*2"))

        // Verify all values
        if case .string(let val) = try sheet.cell("A1").value {
            XCTAssertEqual(val, "Cell A1")
        } else {
            XCTFail("A1 should be string")
        }

        if case .number(let val) = try sheet.cell("A2").value {
            XCTAssertEqual(val, 100)
        } else {
            XCTFail("A2 should be number")
        }

        if case .boolean(let val) = try sheet.cell("A3").value {
            XCTAssertFalse(val)
        } else {
            XCTFail("A3 should be boolean")
        }

        if case .formula(let val) = try sheet.cell("B1").value {
            XCTAssertEqual(val, "=A2*2")
        } else {
            XCTFail("B1 should be formula")
        }
    }

    func testSaveWorkbook() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Modify some cells
        try sheet.cell("A1").setValue(.string("Modified Value"))
        try sheet.cell("B2").setValue(.number(999.99))

        // Save to new file
        let outputURL = outputDirectory.appendingPathComponent("output.xlsm")
        try workbook.save(to: outputURL)

        // Verify file was created
        XCTAssertTrue(FileManager.default.fileExists(atPath: outputURL.path))

        print("Saved modified workbook to: \(outputURL.path)")
    }

    func testSaveAndReload() throws {
        // Step 1: Modify and save
        let workbook1 = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet1 = try workbook1.sheet(at: 0)

        let testString = "Test Data \(Date().timeIntervalSince1970)"
        let testNumber = 42.42

        try sheet1.cell("A1").setValue(.string(testString))
        try sheet1.cell("B1").setValue(.number(testNumber))

        let outputURL = outputDirectory.appendingPathComponent("test_save_reload.xlsm")
        try workbook1.save(to: outputURL)

        // Step 2: Reload and verify
        let workbook2 = try ExcelWorkbook(fileURL: outputURL)
        let sheet2 = try workbook2.sheet(at: 0)

        let cellA1 = try sheet2.cell("A1")
        let cellB1 = try sheet2.cell("B1")

        // Verify values persisted
        if case .string(let value) = cellA1.value {
            XCTAssertEqual(value, testString, "String value should persist after save/reload")
        } else {
            XCTFail("A1 should contain the saved string value")
        }

        if case .number(let value) = cellB1.value {
            XCTAssertEqual(value, testNumber, accuracy: 0.001, "Number value should persist after save/reload")
        } else {
            XCTFail("B1 should contain the saved number value")
        }

        print("Save and reload test passed!")
    }

    func testPreserveMacros() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Verify macros detected
        XCTAssertTrue(workbook.hasVBAMacros, "Original file should have macros")

        // Modify a cell
        let sheet = try workbook.sheet(at: 0)
        try sheet.cell("A1").setValue(.string("Modified"))

        // Save
        let outputURL = outputDirectory.appendingPathComponent("with_macros.xlsm")
        try workbook.save(to: outputURL)

        // Reload and check macros still present
        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        XCTAssertTrue(reloadedWorkbook.hasVBAMacros, "Macros should be preserved after save")

        print("Macros preserved successfully!")
    }

    func testCellCaching() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Get same cell twice
        let cell1 = try sheet.cell("A1")
        let cell2 = try sheet.cell("A1")

        // Should be same object (cached)
        XCTAssertTrue(cell1 === cell2, "Same cell reference should return cached object")

        // Modify through one reference
        cell1.setValue(.string("Cached Test"))

        // Should see change through other reference
        if case .string(let value) = cell2.value {
            XCTAssertEqual(value, "Cached Test", "Cache should ensure same object")
        } else {
            XCTFail("Cell2 should reflect cell1's changes")
        }
    }

    func testModificationTracking() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet1 = try workbook.sheet(at: 0)
        let sheet2 = try workbook.sheet(at: 1)

        // Modify only one sheet
        try sheet1.cell("A1").setValue(.string("Modified"))

        // Only modified sheet should be marked
        XCTAssertTrue(sheet1.isModified, "Modified sheet should be marked")
        XCTAssertFalse(sheet2.isModified, "Unmodified sheet should not be marked")
    }
}
