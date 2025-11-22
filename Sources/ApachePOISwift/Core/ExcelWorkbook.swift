//
// ExcelWorkbook.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents an Excel workbook (.xlsx or .xlsm file)
public class ExcelWorkbook {
    /// URL of the original Excel file
    private let fileURL: URL

    /// Temporary directory where the archive is extracted
    private var extractedDirectory: URL?

    /// All sheets in the workbook
    private var sheets: [ExcelSheet] = []

    /// Shared strings table (indexed strings used across the workbook)
    internal var sharedStrings: [String] = []

    /// Whether the workbook contains VBA macros (.xlsm)
    public private(set) var hasVBAMacros: Bool = false

    /// Initialize from an Excel file
    /// - Parameter fileURL: URL of the .xlsx or .xlsm file
    /// - Throws: ExcelError if the file cannot be opened or parsed
    public init(fileURL: URL) throws {
        self.fileURL = fileURL
        try load()
    }

    /// Initialize from an Excel file in the bundle
    /// - Parameters:
    ///   - name: Name of the file (without extension)
    ///   - bundle: Bundle containing the file (defaults to main bundle)
    ///   - extension: File extension (defaults to "xlsx")
    /// - Throws: ExcelError if the file is not found or cannot be parsed
    public convenience init(name: String, bundle: Bundle = .main, extension fileExtension: String = "xlsx") throws {
        guard let url = bundle.url(forResource: name, withExtension: fileExtension) else {
            throw ExcelError.fileNotFound(URL(fileURLWithPath: "\(name).\(fileExtension)"))
        }
        try self.init(fileURL: url)
    }

    /// Load and parse the Excel file
    private func load() throws {
        // 1. Extract ZIP archive
        extractedDirectory = try ZIPHandler.extractArchive(at: fileURL)

        guard let extractedDir = extractedDirectory else {
            throw ExcelError.invalidZIPArchive(fileURL)
        }

        // 2. Check for VBA macros
        let vbaProjectURL = extractedDir.appendingPathComponent("xl/vbaProject.bin")
        hasVBAMacros = FileManager.default.fileExists(atPath: vbaProjectURL.path)

        // 3. Parse shared strings (if exists)
        let sharedStringsURL = extractedDir.appendingPathComponent("xl/sharedStrings.xml")
        if FileManager.default.fileExists(atPath: sharedStringsURL.path) {
            let data = try Data(contentsOf: sharedStringsURL)
            let parser = SharedStringsParser()
            sharedStrings = try parser.parse(data: data)
        }

        // 4. Parse workbook.xml to get sheet list
        let workbookURL = extractedDir.appendingPathComponent("xl/workbook.xml")
        guard FileManager.default.fileExists(atPath: workbookURL.path) else {
            throw ExcelError.invalidWorkbookStructure("Missing xl/workbook.xml")
        }

        let workbookData = try Data(contentsOf: workbookURL)
        let workbookParser = WorkbookXMLParser()
        let sheetInfos = try workbookParser.parse(data: workbookData)

        guard !sheetInfos.isEmpty else {
            throw ExcelError.invalidWorkbookStructure("No sheets found in workbook")
        }

        // 5. Parse each sheet
        for (index, sheetInfo) in sheetInfos.enumerated() {
            let sheetURL = extractedDir.appendingPathComponent("xl/worksheets/sheet\(index + 1).xml")

            guard FileManager.default.fileExists(atPath: sheetURL.path) else {
                throw ExcelError.invalidWorkbookStructure("Missing sheet file: sheet\(index + 1).xml")
            }

            let sheetData = try Data(contentsOf: sheetURL)
            let sheetParser = SheetXMLParser()
            let cells = try sheetParser.parse(data: sheetData)

            let sheet = ExcelSheet(sheetInfo: sheetInfo, cells: cells, workbook: self)
            sheets.append(sheet)
        }
    }

    /// Get a sheet by its index (zero-based)
    /// - Parameter index: Sheet index (0 = first sheet)
    /// - Returns: ExcelSheet object
    /// - Throws: ExcelError if the index is out of range
    public func sheet(at index: Int) throws -> ExcelSheet {
        guard index >= 0 && index < sheets.count else {
            throw ExcelError.sheetNotFound("Index \(index) out of range (0-\(sheets.count - 1))")
        }
        return sheets[index]
    }

    /// Get a sheet by its name
    /// - Parameter name: Sheet name (case-sensitive)
    /// - Returns: ExcelSheet object
    /// - Throws: ExcelError if no sheet with that name exists
    public func sheet(named name: String) throws -> ExcelSheet {
        guard let sheet = sheets.first(where: { $0.name == name }) else {
            throw ExcelError.sheetNotFound(name)
        }
        return sheet
    }

    /// Get all sheets in the workbook
    public var allSheets: [ExcelSheet] {
        return sheets
    }

    /// Get all sheet names
    public var sheetNames: [String] {
        return sheets.map { $0.name }
    }

    /// Get the number of sheets in the workbook
    public var sheetCount: Int {
        return sheets.count
    }

    /// Clean up temporary files
    deinit {
        if let extractedDir = extractedDirectory {
            ZIPHandler.cleanupTempDirectory(at: extractedDir)
        }
    }
}

extension ExcelWorkbook: CustomStringConvertible {
    public var description: String {
        let macroIndicator = hasVBAMacros ? " [with macros]" : ""
        return "Workbook: \(fileURL.lastPathComponent) (\(sheetCount) sheets)\(macroIndicator)"
    }
}
