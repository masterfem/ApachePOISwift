//
// WorkbookXMLParser.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Information about a sheet in the workbook
struct SheetInfo {
    let name: String
    let sheetId: String
    let relationshipId: String  // r:id attribute
}

/// Parses Excel's workbook.xml file to extract sheet information
class WorkbookXMLParser: NSObject, XMLParserDelegate {
    private var sheets: [SheetInfo] = []

    /// Parse workbook XML data
    /// - Parameter data: XML data from xl/workbook.xml
    /// - Returns: Array of SheetInfo objects
    /// - Throws: ExcelError if parsing fails
    func parse(data: Data) throws -> [SheetInfo] {
        sheets = []

        let parser = XMLParser(data: data)
        parser.delegate = self

        guard parser.parse() else {
            if let error = parser.parserError {
                throw ExcelError.parsingError("Workbook: \(error.localizedDescription)")
            }
            throw ExcelError.parsingError("Workbook: Unknown parsing error")
        }

        return sheets
    }

    // MARK: - XMLParserDelegate

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String] = [:]
    ) {
        if elementName == "sheet" {
            // Extract sheet information
            guard let name = attributeDict["name"],
                  let sheetId = attributeDict["sheetId"] else {
                return
            }

            // The relationship ID can be r:id or id depending on namespace
            let rId = attributeDict["r:id"] ?? attributeDict["id"] ?? ""

            let sheetInfo = SheetInfo(
                name: name,
                sheetId: sheetId,
                relationshipId: rId
            )

            sheets.append(sheetInfo)
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        // Error will be caught in parse() method
    }
}
