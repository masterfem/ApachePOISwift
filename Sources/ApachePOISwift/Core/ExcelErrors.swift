//
// ExcelErrors.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Errors that can occur when working with Excel files
public enum ExcelError: Error {
    /// The specified file was not found
    case fileNotFound(URL)

    /// The file is not a valid ZIP archive
    case invalidZIPArchive(URL)

    /// The XML structure is invalid or malformed
    case invalidXMLStructure(String)

    /// The requested sheet was not found
    case sheetNotFound(String)

    /// The requested cell was not found
    case cellNotFound(String)

    /// The cell reference is invalid (e.g., not in A1 notation)
    case invalidCellReference(String)

    /// The file format is not supported
    case unsupportedFileFormat(String)

    /// A general parsing error occurred
    case parsingError(String)

    /// Failed to create or modify a file
    case fileWriteError(String)

    /// The workbook structure is invalid
    case invalidWorkbookStructure(String)
}

extension ExcelError: LocalizedError {
    public var errorDescription: String? {
        switch self {
        case .fileNotFound(let url):
            return "File not found: \(url.path)"
        case .invalidZIPArchive(let url):
            return "Invalid ZIP archive: \(url.path)"
        case .invalidXMLStructure(let message):
            return "Invalid XML structure: \(message)"
        case .sheetNotFound(let name):
            return "Sheet not found: \(name)"
        case .cellNotFound(let reference):
            return "Cell not found: \(reference)"
        case .invalidCellReference(let reference):
            return "Invalid cell reference: \(reference)"
        case .unsupportedFileFormat(let message):
            return "Unsupported file format: \(message)"
        case .parsingError(let message):
            return "Parsing error: \(message)"
        case .fileWriteError(let message):
            return "File write error: \(message)"
        case .invalidWorkbookStructure(let message):
            return "Invalid workbook structure: \(message)"
        }
    }
}
