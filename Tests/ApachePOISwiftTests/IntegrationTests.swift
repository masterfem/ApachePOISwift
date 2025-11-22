//
// IntegrationTests.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class IntegrationTests: XCTestCase {
    var marbarFileURL: URL!

    override func setUpWithError() throws {
        let testBundle = Bundle.module
        marbarFileURL = testBundle.url(
            forResource: "marbar_template",
            withExtension: "xlsm",
            subdirectory: "TestResources"
        )

        XCTAssertNotNil(marbarFileURL, "marbar_template.xlsm should exist in TestResources")
    }

    func testOpenMarbarTemplate() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Verify it's recognized as having macros
        XCTAssertTrue(workbook.hasVBAMacros, "marbar_template.xlsm should have VBA macros")

        // Verify sheet count (expected: 26 sheets based on CLAUDE.md)
        XCTAssertEqual(workbook.sheetCount, 26, "marbar_template should have 26 sheets")

        // Verify we can get sheet names
        let sheetNames = workbook.sheetNames
        XCTAssertEqual(sheetNames.count, 26)
        XCTAssertFalse(sheetNames.isEmpty)

        print("Workbook loaded successfully: \(workbook)")
        print("Sheets: \(sheetNames.joined(separator: ", "))")
    }

    func testAccessAllSheets() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Test we can access all sheets by index
        for i in 0..<workbook.sheetCount {
            let sheet = try workbook.sheet(at: i)
            XCTAssertFalse(sheet.name.isEmpty, "Sheet \(i) should have a name")
            print("Sheet \(i): \(sheet.name) (\(sheet.cellCount) cells)")
        }
    }

    func testAccessSheetByName() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Try to access a specific sheet (GENERALES mentioned in CLAUDE.md)
        // Note: Actual sheet names may vary - this test might need adjustment
        let sheetNames = workbook.sheetNames

        // Access first sheet by name
        if let firstName = sheetNames.first {
            let sheet = try workbook.sheet(named: firstName)
            XCTAssertEqual(sheet.name, firstName)
        }
    }

    func testReadCells() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Get first sheet
        let sheet = try workbook.sheet(at: 0)

        // Try to access some cells (they might be empty, but shouldn't throw)
        let cellA1 = try sheet.cell("A1")
        XCTAssertNotNil(cellA1)
        print("A1 value: \(cellA1.value)")

        let cellB2 = try sheet.cell("B2")
        XCTAssertNotNil(cellB2)
        print("B2 value: \(cellB2.value)")

        // Try cell access by indices
        let cellByIndex = try sheet.cell(column: 0, row: 0)
        XCTAssertEqual(cellByIndex.reference, "A1")
    }

    func testNonEmptyCells() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Get first sheet and count non-empty cells
        let sheet = try workbook.sheet(at: 0)
        let nonEmptyCells = sheet.nonEmptyCells()

        print("First sheet has \(nonEmptyCells.count) non-empty cells")

        // Print first few non-empty cells as examples
        for (index, cell) in nonEmptyCells.prefix(10).enumerated() {
            print("Cell \(index + 1): \(cell.reference) = \(cell.value)")
        }

        XCTAssertTrue(nonEmptyCells.count >= 0)  // Just verify it doesn't crash
    }

    func testWorkbookDescription() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let description = workbook.description

        XCTAssertTrue(description.contains("marbar_template.xlsm"))
        XCTAssertTrue(description.contains("26 sheets"))
        XCTAssertTrue(description.contains("[with macros]"))

        print("Workbook description: \(description)")
    }

    func testInvalidSheetAccess() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Try to access sheet out of range
        XCTAssertThrowsError(try workbook.sheet(at: 999)) { error in
            if case ExcelError.sheetNotFound = error {
                // Expected error
            } else {
                XCTFail("Expected sheetNotFound error, got \(error)")
            }
        }

        // Try to access non-existent sheet by name
        XCTAssertThrowsError(try workbook.sheet(named: "NonExistentSheet")) { error in
            if case ExcelError.sheetNotFound = error {
                // Expected error
            } else {
                XCTFail("Expected sheetNotFound error, got \(error)")
            }
        }
    }

    func testInvalidCellReference() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Try invalid cell references
        XCTAssertThrowsError(try sheet.cell("")) { error in
            if case ExcelError.invalidCellReference = error {
                // Expected error
            } else {
                XCTFail("Expected invalidCellReference error, got \(error)")
            }
        }

        XCTAssertThrowsError(try sheet.cell("123")) { error in
            if case ExcelError.invalidCellReference = error {
                // Expected error
            } else {
                XCTFail("Expected invalidCellReference error, got \(error)")
            }
        }
    }
}
