//
// MergedCellsTests.swift
// ApachePOISwiftTests
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import XCTest
@testable import ApachePOISwift

final class MergedCellsTests: XCTestCase {
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

    func testReadMergedCells() throws {
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Marbar template has 25 out of 26 sheets with merged cells
        let sheet = try workbook.sheet(at: 0)

        // Access merged cells (internal property)
        let mergedCellCount = sheet.mergedCells.count

        print("Sheet '\(sheet.name)' has \(mergedCellCount) merged cell ranges")

        // Marbar template should have at least some merged cells
        XCTAssertGreaterThan(mergedCellCount, 0, "Sheet should have merged cells")
    }

    func testPreserveMergedCells() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        // Load workbook
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)
        let sheet = try workbook.sheet(at: 0)

        // Get original merged cell count
        let originalMergedCount = sheet.mergedCells.count

        print("Original merged cells: \(originalMergedCount)")
        if originalMergedCount > 0 {
            print("First merged cell: \(sheet.mergedCells[0])")
        }

        // Modify some data (not in merged cells)
        let cell = try sheet.cell("Z99")
        cell.setValue(.string("Test Value"))

        // Save
        let outputURL = outputDirectory.appendingPathComponent("merged_cells_test.xlsm")
        try workbook.save(to: outputURL)

        // Reload
        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)

        // Verify merged cells are preserved
        let reloadedMergedCount = reloadedSheet.mergedCells.count

        print("Reloaded merged cells: \(reloadedMergedCount)")

        XCTAssertEqual(reloadedMergedCount, originalMergedCount,
                       "Merged cells should be preserved after save/reload")

        // Verify the actual ranges match
        if originalMergedCount > 0 {
            XCTAssertEqual(reloadedSheet.mergedCells[0], sheet.mergedCells[0],
                           "First merged cell range should match")
        }
    }

    func testPreserveChartsAndDrawings() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        // Load Marbar template (has 30 charts, 25 drawings)
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        // Modify some data
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("A1")
        cell.setValue(.string("Modified"))

        // Save
        let outputURL = outputDirectory.appendingPathComponent("charts_test.xlsm")
        try workbook.save(to: outputURL)

        // Extract and verify charts/drawings exist
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)

        // Unzip the saved file
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
        process.arguments = ["-q", outputURL.path, "-d", tempDir.path]
        try process.run()
        process.waitUntilExit()

        // Check for charts directory
        let chartsDir = tempDir.appendingPathComponent("xl/charts")
        let chartsExist = FileManager.default.fileExists(atPath: chartsDir.path)

        XCTAssertTrue(chartsExist, "Charts directory should be preserved")

        if chartsExist {
            let chartFiles = try FileManager.default.contentsOfDirectory(atPath: chartsDir.path)
            let chartCount = chartFiles.filter { $0.hasSuffix(".xml") }.count
            print("Preserved \(chartCount) chart files")
            XCTAssertGreaterThan(chartCount, 0, "Should preserve chart files")
        }

        // Check for drawings directory
        let drawingsDir = tempDir.appendingPathComponent("xl/drawings")
        let drawingsExist = FileManager.default.fileExists(atPath: drawingsDir.path)

        XCTAssertTrue(drawingsExist, "Drawings directory should be preserved")

        if drawingsExist {
            let drawingFiles = try FileManager.default.contentsOfDirectory(atPath: drawingsDir.path)
            let drawingCount = drawingFiles.filter { $0.hasSuffix(".xml") }.count
            print("Preserved \(drawingCount) drawing files")
            XCTAssertGreaterThan(drawingCount, 0, "Should preserve drawing files")
        }

        // Cleanup
        try? FileManager.default.removeItem(at: tempDir)
    }

    func testVBAMacroPreservation() throws {
        let outputDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)
        try FileManager.default.createDirectory(at: outputDirectory, withIntermediateDirectories: true)

        // Load Marbar template (has VBA macros)
        let workbook = try ExcelWorkbook(fileURL: marbarFileURL)

        XCTAssertTrue(workbook.hasVBAMacros, "Template should have VBA macros")

        // Modify some data
        let sheet = try workbook.sheet(at: 0)
        let cell = try sheet.cell("A1")
        cell.setValue(.string("Modified"))

        // Save
        let outputURL = outputDirectory.appendingPathComponent("vba_test.xlsm")
        try workbook.save(to: outputURL)

        // Reload and verify macros still exist
        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)

        XCTAssertTrue(reloadedWorkbook.hasVBAMacros,
                      "VBA macros should be preserved after save/reload")
    }
}
