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
    let styleIndex: Int?        // Style index (references cellXfs in styles.xml)
}

/// Data from a worksheet including cells and merged cell ranges
struct SheetData {
    let cells: [String: CellData]
    let mergedCells: [String]  // Array of merged cell ranges (e.g., ["A1:B2", "C3:D5"])
}

/// Parses Excel worksheet XML files (sheet1.xml, sheet2.xml, etc.)
class SheetXMLParser: NSObject, XMLParserDelegate {
    private var cells: [String: CellData] = [:]
    private var mergedCells: [String] = []
    private var currentCellReference: String = ""
    private var currentCellType: CellType?
    private var currentValue: String?
    private var currentFormula: String?
    private var currentStyleIndex: Int?
    private var currentElement = ""
    private var currentText = ""
    private var isParsingInlineString = false

    /// Parse sheet XML data
    /// - Parameter data: XML data from xl/worksheets/sheetN.xml
    /// - Returns: SheetData containing cells and merged cell ranges
    /// - Throws: ExcelError if parsing fails
    func parse(data: Data) throws -> SheetData {
        cells = [:]
        mergedCells = []
        currentCellReference = ""
        currentCellType = nil
        currentValue = nil
        currentFormula = nil
        currentStyleIndex = nil
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

        return SheetData(cells: cells, mergedCells: mergedCells)
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
            // Cell element: <c r="A1" t="s" s="5">
            currentCellReference = attributeDict["r"] ?? ""
            currentCellType = attributeDict["t"].flatMap { CellType(rawValue: $0) }
            currentStyleIndex = attributeDict["s"].flatMap { Int($0) }
            currentValue = nil
            currentFormula = nil
            currentText = ""
            isParsingInlineString = false
        } else if elementName == "v" || elementName == "f" {
            // Value or formula element
            currentText = ""
        } else if elementName == "is" {
            // Inline string element: <is><t>text</t></is>
            isParsingInlineString = true
            currentText = ""
        } else if elementName == "t" && isParsingInlineString {
            // Text element inside inline string
            currentText = ""
        } else if elementName == "mergeCell" {
            // Merged cell element: <mergeCell ref="A1:B2"/>
            if let ref = attributeDict["ref"] {
                mergedCells.append(ref)
            }
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
        } else if elementName == "t" && isParsingInlineString {
            // End of text element in inline string
            currentValue = currentText.trimmingCharacters(in: .whitespacesAndNewlines)
        } else if elementName == "is" {
            // End of inline string element
            isParsingInlineString = false
        } else if elementName == "c" {
            // End of cell - store cell data
            if !currentCellReference.isEmpty {
                let cellData = CellData(
                    reference: currentCellReference,
                    type: currentCellType,
                    value: currentValue,
                    formula: currentFormula,
                    styleIndex: currentStyleIndex
                )
                cells[currentCellReference] = cellData
            }

            // Reset state
            currentCellReference = ""
            currentCellType = nil
            currentValue = nil
            currentFormula = nil
            currentStyleIndex = nil
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        // Error will be caught in parse() method
    }
}
