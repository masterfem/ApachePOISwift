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

    /// Style data (fonts, fills, borders, number formats, cell styles)
    internal var stylesData: StylesData?

    /// Whether the workbook contains VBA macros (.xlsm)
    public private(set) var hasVBAMacros: Bool = false

    /// Whether the workbook has been modified
    private var isModified: Bool = false

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

        // 4. Parse styles (if exists)
        let stylesURL = extractedDir.appendingPathComponent("xl/styles.xml")
        if FileManager.default.fileExists(atPath: stylesURL.path) {
            let data = try Data(contentsOf: stylesURL)
            let parser = StylesXMLParser()
            stylesData = try parser.parse(data: data)
        }

        // 5. Parse workbook.xml to get sheet list
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

        // 6. Parse each sheet
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
        // Always cleanup temp directory
        if let extractedDir = extractedDirectory {
            ZIPHandler.cleanupTempDirectory(at: extractedDir)
        }
    }

    // MARK: - Write Support (Phase 2)

    /// Mark the workbook as modified
    internal func markAsModified() {
        isModified = true
    }

    /// Save the workbook to a new file
    /// - Parameter url: Destination URL for the saved file
    /// - Throws: ExcelError if saving fails
    public func save(to url: URL) throws {
        guard let extractedDir = extractedDirectory else {
            throw ExcelError.invalidWorkbookStructure("No extracted directory available")
        }

        // Write modified sheets back to XML
        for (index, sheet) in sheets.enumerated() {
            if sheet.isModified {
                try writeSheet(sheet, at: index + 1, to: extractedDir)
            }
        }

        // Create new archive from extracted directory
        // Note: This must happen before deinit cleans up extractedDirectory
        try ZIPHandler.createArchive(from: extractedDir, to: url)

        // Mark sheets as no longer modified after successful save
        for sheet in sheets {
            sheet.isModified = false
        }
        isModified = false
    }

    /// Write a sheet's XML file
    private func writeSheet(_ sheet: ExcelSheet, at sheetNumber: Int, to directory: URL) throws {
        let sheetURL = directory.appendingPathComponent("xl/worksheets/sheet\(sheetNumber).xml")

        // Read original XML to preserve structure
        let originalXML = try String(contentsOf: sheetURL, encoding: .utf8)

        // Generate new sheet data
        let newSheetData = try generateSheetDataXML(for: sheet)

        // Replace <sheetData>...</sheetData> section
        let updatedXML = try replaceSheetData(in: originalXML, with: newSheetData)

        // Write to file
        try updatedXML.write(to: sheetURL, atomically: true, encoding: .utf8)
    }

    /// Replace the sheetData section in the XML
    private func replaceSheetData(in xml: String, with newData: String) throws -> String {
        // Find <sheetData> start and end
        guard let sheetDataStart = xml.range(of: "<sheetData>"),
              let sheetDataEnd = xml.range(of: "</sheetData>") else {
            throw ExcelError.parsingError("Cannot find <sheetData> tags in sheet XML")
        }

        // Replace the content
        var result = xml
        let rangeToReplace = sheetDataStart.upperBound..<sheetDataEnd.lowerBound
        result.replaceSubrange(rangeToReplace, with: newData)

        return result
    }

    /// Generate XML for sheet data only (cells and rows)
    private func generateSheetDataXML(for sheet: ExcelSheet) throws -> String {
        var xml = "\n"

        // Get all cells grouped by row
        let cells = sheet.getAllCellsForWriting()
        var cellsByRow: [Int: [ExcelCell]] = [:]

        for cell in cells {
            if let cellRef = cell.cellReference {
                if cellsByRow[cellRef.row] == nil {
                    cellsByRow[cellRef.row] = []
                }
                cellsByRow[cellRef.row]?.append(cell)
            }
        }

        // Write rows in order
        for rowIndex in cellsByRow.keys.sorted() {
            guard let rowCells = cellsByRow[rowIndex] else { continue }

            xml += "<row r=\"\(rowIndex + 1)\">\n"

            // Sort cells by column
            for cell in rowCells.sorted(by: { ($0.cellReference?.column ?? 0) < ($1.cellReference?.column ?? 0) }) {
                xml += try generateCellXML(for: cell)
            }

            xml += "</row>\n"
        }

        return xml
    }

    /// Generate XML for a single cell
    private func generateCellXML(for cell: ExcelCell) throws -> String {
        let cellData = cell.getCellData()

        // Skip truly empty cells (no value, formula, or style)
        if cellData.value == nil && cellData.formula == nil && cellData.styleIndex == nil {
            return ""
        }

        var cellXML = "<c r=\"\(cell.reference)\""

        // Add type attribute if present
        if let type = cellData.type {
            cellXML += " t=\"\(type.rawValue)\""
        }

        // Add style attribute if present
        if let styleIndex = cellData.styleIndex {
            cellXML += " s=\"\(styleIndex)\""
        }

        cellXML += ">"

        // Add formula if present
        if let formula = cellData.formula {
            cellXML += "<f>\(xmlEscape(formula))</f>"
        }

        // Add value if present
        if let value = cellData.value {
            if cellData.type == .inlineString {
                // Inline string needs special formatting
                cellXML += "<is><t>\(xmlEscape(value))</t></is>"
            } else {
                cellXML += "<v>\(xmlEscape(value))</v>"
            }
        }

        cellXML += "</c>\n"

        return cellXML
    }

    /// Escape special XML characters
    private func xmlEscape(_ string: String) -> String {
        return string
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
            .replacingOccurrences(of: "'", with: "&apos;")
    }
}

extension ExcelWorkbook: CustomStringConvertible {
    public var description: String {
        let macroIndicator = hasVBAMacros ? " [with macros]" : ""
        return "Workbook: \(fileURL.lastPathComponent) (\(sheetCount) sheets)\(macroIndicator)"
    }
}
