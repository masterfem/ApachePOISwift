//
// FormulaEvaluationTests.swift
// ApachePOISwiftTests
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class FormulaEvaluationTests: XCTestCase {
    var workbook: ExcelWorkbook!
    var sheet: ExcelSheet!

    override func setUp() {
        super.setUp()

        // Use existing test workbook
        let bundle = Bundle.module
        guard let url = bundle.url(forResource: "test_workbook", withExtension: "xlsx", subdirectory: "TestResources") else {
            // If test_workbook doesn't exist, use marbar template which we know exists
            guard let marbarURL = bundle.url(forResource: "marbar_template", withExtension: "xlsm", subdirectory: "TestResources") else {
                XCTFail("No test workbook found")
                return
            }
            workbook = try? ExcelWorkbook(fileURL: marbarURL)
            sheet = try? workbook.sheet(at: 0)
            return
        }

        workbook = try? ExcelWorkbook(fileURL: url)
        sheet = try? workbook.sheet(at: 0)
    }

    // MARK: - Arithmetic Operations

    func testSimpleAddition() throws {
        let result = try workbook.evaluateFormula("=1+2", in: sheet)
        XCTAssertEqual(result, .number(3.0))
    }

    func testSimpleSubtraction() throws {
        let result = try workbook.evaluateFormula("=10-3", in: sheet)
        XCTAssertEqual(result, .number(7.0))
    }

    func testSimpleMultiplication() throws {
        let result = try workbook.evaluateFormula("=5*4", in: sheet)
        XCTAssertEqual(result, .number(20.0))
    }

    func testSimpleDivision() throws {
        let result = try workbook.evaluateFormula("=20/4", in: sheet)
        XCTAssertEqual(result, .number(5.0))
    }

    func testDivisionByZero() throws {
        let result = try workbook.evaluateFormula("=10/0", in: sheet)
        XCTAssertEqual(result, .error(.divideByZero))
    }

    func testPower() throws {
        let result = try workbook.evaluateFormula("=2^8", in: sheet)
        XCTAssertEqual(result, .number(256.0))
    }

    func testOperatorPrecedence() throws {
        let result = try workbook.evaluateFormula("=2+3*4", in: sheet)
        XCTAssertEqual(result, .number(14.0))  // Not 20
    }

    func testParentheses() throws {
        let result = try workbook.evaluateFormula("=(2+3)*4", in: sheet)
        XCTAssertEqual(result, .number(20.0))
    }

    func testComplexExpression() throws {
        let result = try workbook.evaluateFormula("=10+2*3^2-5", in: sheet)
        XCTAssertEqual(result, .number(23.0))  // 10 + 2*9 - 5 = 23
    }

    // MARK: - String Operations

    func testStringConcatenation() throws {
        let result = try workbook.evaluateFormula("=\"Hello\" & \" \" & \"World\"", in: sheet)
        XCTAssertEqual(result, .string("Hello World"))
    }

    // MARK: - Comparison Operations

    func testEqual() throws {
        let result = try workbook.evaluateFormula("=5=5", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    func testNotEqual() throws {
        let result = try workbook.evaluateFormula("=5<>3", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    func testLessThan() throws {
        let result = try workbook.evaluateFormula("=3<5", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    func testGreaterThan() throws {
        let result = try workbook.evaluateFormula("=10>5", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    // MARK: - Mathematical Functions

    func testSUM() throws {
        let result = try workbook.evaluateFormula("=SUM(1, 2, 3, 4, 5)", in: sheet)
        XCTAssertEqual(result, .number(15.0))
    }

    func testAVERAGE() throws {
        let result = try workbook.evaluateFormula("=AVERAGE(10, 20, 30)", in: sheet)
        XCTAssertEqual(result, .number(20.0))
    }

    func testCOUNT() throws {
        let result = try workbook.evaluateFormula("=COUNT(1, 2, 3)", in: sheet)
        XCTAssertEqual(result, .number(3.0))
    }

    func testMIN() throws {
        let result = try workbook.evaluateFormula("=MIN(5, 2, 8, 1, 9)", in: sheet)
        XCTAssertEqual(result, .number(1.0))
    }

    func testMAX() throws {
        let result = try workbook.evaluateFormula("=MAX(5, 2, 8, 1, 9)", in: sheet)
        XCTAssertEqual(result, .number(9.0))
    }

    func testABS() throws {
        let result = try workbook.evaluateFormula("=ABS(-42)", in: sheet)
        XCTAssertEqual(result, .number(42.0))
    }

    func testROUND() throws {
        let result = try workbook.evaluateFormula("=ROUND(3.14159, 2)", in: sheet)
        XCTAssertEqual(result, .number(3.14))
    }

    func testINT() throws {
        let result = try workbook.evaluateFormula("=INT(7.8)", in: sheet)
        XCTAssertEqual(result, .number(7.0))
    }

    // MARK: - Logical Functions

    func testIF_true() throws {
        let result = try workbook.evaluateFormula("=IF(5>3, \"Yes\", \"No\")", in: sheet)
        XCTAssertEqual(result, .string("Yes"))
    }

    func testIF_false() throws {
        let result = try workbook.evaluateFormula("=IF(2>5, \"Yes\", \"No\")", in: sheet)
        XCTAssertEqual(result, .string("No"))
    }

    func testAND_true() throws {
        let result = try workbook.evaluateFormula("=AND(TRUE, TRUE, TRUE)", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    func testAND_false() throws {
        let result = try workbook.evaluateFormula("=AND(TRUE, FALSE, TRUE)", in: sheet)
        XCTAssertEqual(result, .boolean(false))
    }

    func testOR_true() throws {
        let result = try workbook.evaluateFormula("=OR(FALSE, TRUE, FALSE)", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    func testOR_false() throws {
        let result = try workbook.evaluateFormula("=OR(FALSE, FALSE, FALSE)", in: sheet)
        XCTAssertEqual(result, .boolean(false))
    }

    func testNOT() throws {
        let result = try workbook.evaluateFormula("=NOT(FALSE)", in: sheet)
        XCTAssertEqual(result, .boolean(true))
    }

    // MARK: - Text Functions

    func testCONCATENATE() throws {
        let result = try workbook.evaluateFormula("=CONCATENATE(\"Hello\", \" \", \"World\")", in: sheet)
        XCTAssertEqual(result, .string("Hello World"))
    }

    func testLEFT() throws {
        let result = try workbook.evaluateFormula("=LEFT(\"Excel\", 3)", in: sheet)
        XCTAssertEqual(result, .string("Exc"))
    }

    func testRIGHT() throws {
        let result = try workbook.evaluateFormula("=RIGHT(\"Excel\", 3)", in: sheet)
        XCTAssertEqual(result, .string("cel"))
    }

    func testMID() throws {
        let result = try workbook.evaluateFormula("=MID(\"Excel\", 2, 3)", in: sheet)
        XCTAssertEqual(result, .string("xce"))
    }

    func testLEN() throws {
        let result = try workbook.evaluateFormula("=LEN(\"Hello World\")", in: sheet)
        XCTAssertEqual(result, .number(11.0))
    }

    func testUPPER() throws {
        let result = try workbook.evaluateFormula("=UPPER(\"hello\")", in: sheet)
        XCTAssertEqual(result, .string("HELLO"))
    }

    func testLOWER() throws {
        let result = try workbook.evaluateFormula("=LOWER(\"HELLO\")", in: sheet)
        XCTAssertEqual(result, .string("hello"))
    }

    func testTRIM() throws {
        let result = try workbook.evaluateFormula("=TRIM(\"  Hello   World  \")", in: sheet)
        XCTAssertEqual(result, .string("Hello World"))
    }

    // MARK: - Nested Functions

    func testNestedFunctions() throws {
        let result = try workbook.evaluateFormula("=SUM(1, 2, MAX(3, 4, 5))", in: sheet)
        XCTAssertEqual(result, .number(8.0))  // 1 + 2 + 5
    }

    func testComplexNested() throws {
        let result = try workbook.evaluateFormula("=IF(SUM(1,2,3)>5, \"High\", \"Low\")", in: sheet)
        XCTAssertEqual(result, .string("High"))
    }

    // MARK: - Type Coercion

    func testNumberToString() throws {
        let result = try workbook.evaluateFormula("=CONCATENATE(\"Value: \", 42)", in: sheet)
        XCTAssertEqual(result, .string("Value: 42"))
    }

    func testStringToNumber() throws {
        // "5" should be coerced to number in arithmetic
        let result = try workbook.evaluateFormula("=SUM(1, 2, 3)", in: sheet)
        XCTAssertEqual(result, .number(6.0))
    }

    // MARK: - Edge Cases

    func testEmptyFormula() throws {
        XCTAssertThrowsError(try workbook.evaluateFormula("=", in: sheet))
    }

    func testInvalidFunction() throws {
        let result = try workbook.evaluateFormula("=NONEXISTENT()", in: sheet)
        XCTAssertEqual(result, .error(.name))
    }

    func testUnbalancedParentheses() throws {
        XCTAssertThrowsError(try workbook.evaluateFormula("=(1+2", in: sheet))
    }
}
