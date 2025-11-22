//
// SharedStringsParser.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Parses Excel's sharedStrings.xml file
/// Shared strings are stored in a separate file and referenced by index in cells
class SharedStringsParser: NSObject, XMLParserDelegate {
    private var strings: [String] = []
    private var currentString = ""
    private var isParsingText = false
    private var currentElement = ""

    /// Parse shared strings XML data
    /// - Parameter data: XML data from sharedStrings.xml
    /// - Returns: Array of strings indexed by their position
    /// - Throws: ExcelError if parsing fails
    func parse(data: Data) throws -> [String] {
        strings = []
        currentString = ""
        isParsingText = false

        let parser = XMLParser(data: data)
        parser.delegate = self

        guard parser.parse() else {
            if let error = parser.parserError {
                throw ExcelError.parsingError("SharedStrings: \(error.localizedDescription)")
            }
            throw ExcelError.parsingError("SharedStrings: Unknown parsing error")
        }

        return strings
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

        if elementName == "si" {
            // Start of a shared string item
            currentString = ""
        } else if elementName == "t" {
            // Start of text element
            isParsingText = true
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if isParsingText {
            currentString += string
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        if elementName == "t" {
            // End of text element
            isParsingText = false
        } else if elementName == "si" {
            // End of shared string item - add to array
            strings.append(currentString)
            currentString = ""
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        // Error will be caught in parse() method
    }
}
