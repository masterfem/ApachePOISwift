//
// StylesXMLParser.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Container for all parsed style data
public struct StylesData {
    var fonts: [Font] = []
    var fills: [Fill] = []
    var borders: [Border] = []
    var numberFormats: [NumberFormat] = []
    var cellStyles: [CellStyle] = []
}

/// Parses Excel styles.xml file
class StylesXMLParser: NSObject, XMLParserDelegate {
    private var stylesData = StylesData()

    // Current parsing state
    private var currentElement = ""
    private var currentText = ""

    // Font parsing
    private var currentFont: Font?
    private var fontIndex = 0

    // Fill parsing
    private var currentFill: Fill?
    private var fillIndex = 0
    private var currentPatternType: PatternType?

    // Border parsing
    private var currentBorder: Border?
    private var borderIndex = 0
    private var currentBorderEdge: BorderEdge?
    private var currentBorderPosition = ""

    // Number format parsing
    private var numberFormats: [NumberFormat] = []

    // Cell style (cellXfs) parsing
    private var currentCellStyle: CellStyle?
    private var cellStyleIndex = 0
    private var inCellXfs = false  // Track if we're in cellXfs vs cellStyleXfs

    // Alignment parsing
    private var parsingAlignment = false

    /// Parse styles XML data
    /// - Parameter data: XML data from xl/styles.xml
    /// - Returns: StylesData containing all parsed styles
    /// - Throws: ExcelError if parsing fails
    func parse(data: Data) throws -> StylesData {
        stylesData = StylesData()
        fontIndex = 0
        fillIndex = 0
        borderIndex = 0
        cellStyleIndex = 0
        inCellXfs = false

        let parser = XMLParser(data: data)
        parser.delegate = self

        guard parser.parse() else {
            if let error = parser.parserError {
                throw ExcelError.parsingError("Styles: \(error.localizedDescription)")
            }
            throw ExcelError.parsingError("Styles: Unknown parsing error")
        }

        return stylesData
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
        currentText = ""

        switch elementName {
        // Font elements
        case "font":
            currentFont = Font()

        case "sz":
            if let val = attributeDict["val"], let size = Double(val) {
                currentFont?.size = size
            }

        case "name":
            if let val = attributeDict["val"] {
                currentFont?.name = val
            }

        case "b":
            currentFont?.bold = true

        case "i":
            currentFont?.italic = true

        case "strike":
            currentFont?.strikethrough = true

        case "u":
            let val = attributeDict["val"] ?? "single"
            currentFont?.underline = UnderlineStyle(rawValue: val)

        case "color":
            if currentFont != nil {
                // Font color
                if let rgb = attributeDict["rgb"] {
                    currentFont?.color = rgb
                } else if let theme = attributeDict["theme"] {
                    // Theme colors - for now just store the theme index
                    currentFont?.color = "theme:\(theme)"
                }
            } else if currentBorderEdge != nil {
                // Border color
                if let rgb = attributeDict["rgb"] {
                    currentBorderEdge?.color = rgb
                } else if let theme = attributeDict["theme"] {
                    currentBorderEdge?.color = "theme:\(theme)"
                }
            } else if currentPatternType != nil {
                // Fill color
                if let rgb = attributeDict["rgb"] {
                    if currentElement == "fgColor" || attributeDict["rgb"] != nil {
                        currentFill?.foregroundColor = rgb
                    }
                } else if let theme = attributeDict["theme"] {
                    currentFill?.foregroundColor = "theme:\(theme)"
                }
            }

        case "family":
            if let val = attributeDict["val"], let family = Int(val) {
                currentFont?.family = family
            }

        case "charset":
            if let val = attributeDict["val"], let charset = Int(val) {
                currentFont?.charset = charset
            }

        // Fill elements
        case "fill":
            currentFill = Fill()
            currentPatternType = nil

        case "patternFill":
            if let patternType = attributeDict["patternType"] {
                currentPatternType = PatternType(rawValue: patternType)
                currentFill?.patternType = currentPatternType ?? .none
            }

        case "fgColor":
            if let rgb = attributeDict["rgb"] {
                currentFill?.foregroundColor = rgb
            } else if let theme = attributeDict["theme"] {
                currentFill?.foregroundColor = "theme:\(theme)"
            }

        case "bgColor":
            if let rgb = attributeDict["rgb"] {
                currentFill?.backgroundColor = rgb
            } else if let theme = attributeDict["theme"] {
                currentFill?.backgroundColor = "theme:\(theme)"
            }

        // Border elements
        case "border":
            currentBorder = Border()
            if let diagonalUp = attributeDict["diagonalUp"], diagonalUp == "1" {
                currentBorder?.diagonalUp = true
            }
            if let diagonalDown = attributeDict["diagonalDown"], diagonalDown == "1" {
                currentBorder?.diagonalDown = true
            }

        case "left", "right", "top", "bottom", "diagonal":
            currentBorderPosition = elementName
            if let style = attributeDict["style"], let borderStyle = BorderStyle(rawValue: style) {
                currentBorderEdge = BorderEdge(style: borderStyle)
            } else {
                currentBorderEdge = nil
            }

        // Number format elements
        case "numFmt":
            if let formatId = attributeDict["numFmtId"], let id = Int(formatId) {
                let formatCode = attributeDict["formatCode"]
                numberFormats.append(NumberFormat(formatId: id, formatCode: formatCode))
            }

        // Track when we enter cellXfs section
        case "cellXfs":
            inCellXfs = true
            cellStyleIndex = 0  // Reset counter when entering cellXfs

        // Cell style (cellXfs) elements
        case "xf":
            // Only process if we're in cellXfs section (not cellStyleXfs)
            if inCellXfs {
                var fontId: Int?
                var fillId: Int?
                var borderId: Int?
                var numFmtId: Int?
                var applyFont = false
                var applyFill = false
                var applyBorder = false
                var applyNumberFormat = false
                var applyAlignment = false

                if let val = attributeDict["fontId"], let id = Int(val) {
                    fontId = id
                }
                if let val = attributeDict["fillId"], let id = Int(val) {
                    fillId = id
                }
                if let val = attributeDict["borderId"], let id = Int(val) {
                    borderId = id
                }
                if let val = attributeDict["numFmtId"], let id = Int(val) {
                    numFmtId = id
                }
                if attributeDict["applyFont"] == "1" {
                    applyFont = true
                }
                if attributeDict["applyFill"] == "1" {
                    applyFill = true
                }
                if attributeDict["applyBorder"] == "1" {
                    applyBorder = true
                }
                if attributeDict["applyNumberFormat"] == "1" {
                    applyNumberFormat = true
                }
                if attributeDict["applyAlignment"] == "1" {
                    applyAlignment = true
                }

                currentCellStyle = CellStyle(
                    index: cellStyleIndex,
                    fontId: fontId,
                    fillId: fillId,
                    borderId: borderId,
                    numberFormatId: numFmtId,
                    applyFont: applyFont,
                    applyFill: applyFill,
                    applyBorder: applyBorder,
                    applyNumberFormat: applyNumberFormat,
                    applyAlignment: applyAlignment
                )
            }

        case "alignment":
            parsingAlignment = true
            if let horizontal = attributeDict["horizontal"] {
                currentCellStyle?.horizontalAlignment = HorizontalAlignment(rawValue: horizontal)
            }
            if let vertical = attributeDict["vertical"] {
                currentCellStyle?.verticalAlignment = VerticalAlignment(rawValue: vertical)
            }
            if attributeDict["wrapText"] == "1" {
                currentCellStyle?.wrapText = true
            }
            if let rotation = attributeDict["textRotation"], let rot = Int(rotation) {
                currentCellStyle?.textRotation = rot
            }
            if let indent = attributeDict["indent"], let ind = Int(indent) {
                currentCellStyle?.indent = ind
            }

        default:
            break
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
        switch elementName {
        case "font":
            if let font = currentFont {
                stylesData.fonts.append(font)
                fontIndex += 1
            }
            currentFont = nil

        case "fill":
            if let fill = currentFill {
                stylesData.fills.append(fill)
                fillIndex += 1
            }
            currentFill = nil
            currentPatternType = nil

        case "border":
            if let border = currentBorder {
                stylesData.borders.append(border)
                borderIndex += 1
            }
            currentBorder = nil

        case "left":
            if let edge = currentBorderEdge {
                currentBorder?.left = edge
            }
            currentBorderEdge = nil

        case "right":
            if let edge = currentBorderEdge {
                currentBorder?.right = edge
            }
            currentBorderEdge = nil

        case "top":
            if let edge = currentBorderEdge {
                currentBorder?.top = edge
            }
            currentBorderEdge = nil

        case "bottom":
            if let edge = currentBorderEdge {
                currentBorder?.bottom = edge
            }
            currentBorderEdge = nil

        case "diagonal":
            if let edge = currentBorderEdge {
                currentBorder?.diagonal = edge
            }
            currentBorderEdge = nil

        case "numFmts":
            // End of number formats section
            stylesData.numberFormats = numberFormats

        case "cellXfs":
            // End of cellXfs section
            inCellXfs = false

        case "xf":
            if let cellStyle = currentCellStyle {
                stylesData.cellStyles.append(cellStyle)
                cellStyleIndex += 1
            }
            currentCellStyle = nil

        case "alignment":
            parsingAlignment = false

        default:
            break
        }

        currentElement = ""
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        // Error will be caught in parse() method
    }
}
