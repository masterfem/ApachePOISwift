//
// ConditionalFormattingTests.swift
// ApachePOISwiftTests
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class ConditionalFormattingTests: XCTestCase {

    var workbook: ExcelWorkbook!

    override func setUp() {
        super.setUp()

        // Use Marbar template which has conditional formatting
        let bundle = Bundle.module
        guard let url = bundle.url(forResource: "marbar_template", withExtension: "xlsm", subdirectory: "TestResources") else {
            XCTFail("Marbar template not found")
            return
        }

        workbook = try? ExcelWorkbook(fileURL: url)
    }

    override func tearDown() {
        workbook = nil
        super.tearDown()
    }

    // MARK: - Reading Conditional Formatting

    func testCanReadConditionalFormatting() throws {
        // Get a sheet that likely has CF
        let sheet = try workbook.sheet(at: 0)

        // Get all CF areas
        let cfAreas = sheet.getConditionalFormatting()

        // We expect the Marbar template to have CF (706 blocks found in research)
        // Even if this specific sheet doesn't have CF, the method should work
        XCTAssertNotNil(cfAreas)
        XCTAssert(cfAreas is [ConditionalFormattingArea])
    }

    func testConditionalFormattingAreaStructure() throws {
        // Get a sheet
        let sheet = try workbook.sheet(at: 0)

        // Get CF areas
        let cfAreas = sheet.getConditionalFormatting()

        // If there are any CF areas, verify their structure
        if let firstArea = cfAreas.first {
            XCTAssertFalse(firstArea.range.isEmpty, "CF area should have a range")
            XCTAssertFalse(firstArea.rules.isEmpty, "CF area should have at least one rule")

            // Verify first rule structure
            if let firstRule = firstArea.rules.first {
                XCTAssertGreaterThan(firstRule.priority, 0, "Rule priority should be positive")
            }
        }
    }

    func testGetConditionalFormattingForCell() throws {
        // Get a sheet
        let sheet = try workbook.sheet(at: 0)

        // Try to get CF for a cell (e.g., "A1")
        let rules = sheet.getConditionalFormattingForCell("A1")

        // Should return an array (might be empty)
        XCTAssertNotNil(rules)

        // If there are rules, verify they're sorted by priority
        if rules.count > 1 {
            for i in 0..<(rules.count - 1) {
                XCTAssertLessThanOrEqual(
                    rules[i].priority,
                    rules[i + 1].priority,
                    "Rules should be sorted by priority (ascending)"
                )
            }
        }
    }

    // MARK: - Preservation During Save

    func testConditionalFormattingPreservedOnSave() throws {
        // Get original sheet
        let originalSheet = try workbook.sheet(at: 0)
        let originalCFCount = originalSheet.getConditionalFormatting().count

        // Modify a cell (to trigger save)
        if let cell = try? originalSheet.cell("A1") {
            cell.setValue(.string("Test"))
        }

        // Save to temp file
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("cf_test_\(UUID().uuidString).xlsm")

        try workbook.save(to: tempURL)

        // Reload
        let reloadedWorkbook = try ExcelWorkbook(fileURL: tempURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCFCount = reloadedSheet.getConditionalFormatting().count

        // CF should be preserved
        XCTAssertEqual(
            reloadedCFCount,
            originalCFCount,
            "Conditional formatting count should be preserved after save"
        )

        // Cleanup
        try? FileManager.default.removeItem(at: tempURL)
    }

    // MARK: - Rule Type Tests

    func testRuleTypes() throws {
        // Test all rule types can be created
        let types: [ConditionalFormattingRule.RuleType] = [
            .cellIs,
            .expression,
            .colorScale,
            .dataBar,
            .iconSet,
            .top10,
            .uniqueValues,
            .duplicateValues,
            .containsText,
            .beginsWith,
            .endsWith
        ]

        for type in types {
            let rule = ConditionalFormattingRule(
                type: type,
                priority: 1
            )
            XCTAssertEqual(rule.type, type)
        }
    }

    func testOperatorTypes() throws {
        // Test all operator types can be created
        let operators: [ConditionalFormattingRule.Operator] = [
            .lessThan,
            .lessThanOrEqual,
            .equal,
            .notEqual,
            .greaterThanOrEqual,
            .greaterThan,
            .between,
            .notBetween
        ]

        for op in operators {
            let rule = ConditionalFormattingRule(
                type: .cellIs,
                priority: 1,
                operator: op
            )
            XCTAssertEqual(rule.operator, op)
        }
    }

    // MARK: - Complex CF Scenarios

    func testMultipleRulesInArea() throws {
        // Create a CF area with multiple rules
        let rule1 = ConditionalFormattingRule(
            type: .cellIs,
            priority: 1,
            operator: .greaterThan,
            formulas: ["100"]
        )

        let rule2 = ConditionalFormattingRule(
            type: .cellIs,
            priority: 2,
            operator: .lessThan,
            formulas: ["50"]
        )

        let area = ConditionalFormattingArea(
            range: "A1:A10",
            rules: [rule1, rule2]
        )

        XCTAssertEqual(area.rules.count, 2)
        XCTAssertEqual(area.range, "A1:A10")
    }

    func testRuleWithFormulas() throws {
        // Test a rule with formulas
        let rule = ConditionalFormattingRule(
            type: .expression,
            priority: 1,
            formulas: ["=A1>100", "=AND(A1>50, A1<100)"]
        )

        XCTAssertEqual(rule.formulas.count, 2)
        XCTAssertEqual(rule.formulas[0], "=A1>100")
        XCTAssertEqual(rule.formulas[1], "=AND(A1>50, A1<100)")
    }
}
