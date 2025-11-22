//
// SheetXMLParser.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Data from a cell in a worksheet
struct CellData {
    let reference: String       // Cell reference (e.g., "A1")
    let type: CellType?         // Cell type (string, number, etc.)
    let value: String?          // Cell value
    let formula: String?        // Cell formula (if any)
}

/// Parses Excel worksheet XML files (sheet1.xml, sheet2.xml, etc.)
class SheetXMLParser: NSObject, XMLParserDelegate {
    private var cells: [String: CellData] = [:]
    private var currentCellReference: String = ""
    private var currentCellType: CellType?
    private var currentValue: String?
    private var currentFormula: String?
    private var currentElement = ""
    private var currentText = ""

    /// Parse sheet XML data
    /// - Parameter data: XML data from xl/worksheets/sheetN.xml
    /// - Returns: Dictionary of cell data keyed by cell reference
    /// - Throws: ExcelError if parsing fails
    func parse(data: Data) throws -> [String: CellData] {
        cells = [:]
        currentCellReference = ""
        currentCellType = nil
        currentValue = nil
        currentFormula = nil
        currentElement = ""
        currentText = ""

        let parser = XMLParser(data: data)
        parser.delegate = self

        guard parser.parse() else {
            if let error = parser.parserError {
                throw ExcelError.parsingError("Sheet: \(error.localizedDescription)")
            }
            throw ExcelError.parsingError("Sheet: Unknown parsing error")
        }

        return cells
    }

    // MARK: - XMLParserDelegate

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String] = [:]
    ) {
        currentElement = elementName

        if elementName == "c" {
            // Cell element: <c r="A1" t="s">
            currentCellReference = attributeDict["r"] ?? ""
            currentCellType = attributeDict["t"].flatMap { CellType(rawValue: $0) }
            currentValue = nil
            currentFormula = nil
            currentText = ""
        } else if elementName == "v" || elementName == "f" {
            // Value or formula element
            currentText = ""
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        currentText += string
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        if elementName == "v" {
            // End of value element
            currentValue = currentText.trimmingCharacters(in: .whitespacesAndNewlines)
        } else if elementName == "f" {
            // End of formula element
            currentFormula = currentText.trimmingCharacters(in: .whitespacesAndNewlines)
        } else if elementName == "c" {
            // End of cell - store cell data
            if !currentCellReference.isEmpty {
                let cellData = CellData(
                    reference: currentCellReference,
                    type: currentCellType,
                    value: currentValue,
                    formula: currentFormula
                )
                cells[currentCellReference] = cellData
            }

            // Reset state
            currentCellReference = ""
            currentCellType = nil
            currentValue = nil
            currentFormula = nil
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        // Error will be caught in parse() method
    }
}
