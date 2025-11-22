//
// CellReferenceTests.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class CellReferenceTests: XCTestCase {

    func testSimpleReference() throws {
        let ref = try CellReference("A1")
        XCTAssertEqual(ref.column, 0)
        XCTAssertEqual(ref.row, 0)
        XCTAssertFalse(ref.absoluteColumn)
        XCTAssertFalse(ref.absoluteRow)
    }

    func testAbsoluteReference() throws {
        let ref = try CellReference("$B$2")
        XCTAssertEqual(ref.column, 1)
        XCTAssertEqual(ref.row, 1)
        XCTAssertTrue(ref.absoluteColumn)
        XCTAssertTrue(ref.absoluteRow)
    }

    func testMixedAbsoluteReference() throws {
        let ref1 = try CellReference("$C3")
        XCTAssertEqual(ref1.column, 2)
        XCTAssertEqual(ref1.row, 2)
        XCTAssertTrue(ref1.absoluteColumn)
        XCTAssertFalse(ref1.absoluteRow)

        let ref2 = try CellReference("D$4")
        XCTAssertEqual(ref2.column, 3)
        XCTAssertEqual(ref2.row, 3)
        XCTAssertFalse(ref2.absoluteColumn)
        XCTAssertTrue(ref2.absoluteRow)
    }

    func testLargeColumn() throws {
        let ref = try CellReference("AA10")
        XCTAssertEqual(ref.column, 26)  // AA = 26 (0-based)
        XCTAssertEqual(ref.row, 9)
    }

    func testVeryLargeColumn() throws {
        let ref = try CellReference("XFD1")
        XCTAssertEqual(ref.column, 16383)  // Maximum Excel column
        XCTAssertEqual(ref.row, 0)
    }

    func testToExcelNotation() throws {
        let ref1 = try CellReference("A1")
        XCTAssertEqual(ref1.toExcelNotation(), "A1")

        let ref2 = try CellReference("$B$2")
        XCTAssertEqual(ref2.toExcelNotation(), "$B$2")

        let ref3 = try CellReference("AA10")
        XCTAssertEqual(ref3.toExcelNotation(), "AA10")
    }

    func testInvalidReference() {
        XCTAssertThrowsError(try CellReference(""))
        XCTAssertThrowsError(try CellReference("123"))
        XCTAssertThrowsError(try CellReference("A"))
        XCTAssertThrowsError(try CellReference("1A"))
    }

    func testInitWithIndices() throws {
        let ref = try CellReference(column: 0, row: 0)
        XCTAssertEqual(ref.toExcelNotation(), "A1")

        let ref2 = try CellReference(column: 26, row: 9)
        XCTAssertEqual(ref2.toExcelNotation(), "AA10")
    }

    func testEquality() throws {
        let ref1 = try CellReference("A1")
        let ref2 = try CellReference("A1")
        let ref3 = try CellReference("B2")

        XCTAssertEqual(ref1, ref2)
        XCTAssertNotEqual(ref1, ref3)
    }
}
