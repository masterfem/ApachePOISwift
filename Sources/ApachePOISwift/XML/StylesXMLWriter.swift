//
// StylesXMLWriter.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Writes Excel styles.xml file
class StylesXMLWriter {

    /// Generate styles.xml content
    /// - Parameter stylesData: The styles data to write
    /// - Returns: XML string for styles.xml
    func generateXML(stylesData: StylesData) -> String {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        xml += "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"

        // Number formats
        if !stylesData.numberFormats.isEmpty {
            xml += generateNumberFormatsXML(stylesData.numberFormats)
        }

        // Fonts
        xml += generateFontsXML(stylesData.fonts)

        // Fills
        xml += generateFillsXML(stylesData.fills)

        // Borders
        xml += generateBordersXML(stylesData.borders)

        // Cell style formats (cellStyleXfs) - required element
        xml += "<cellStyleXfs count=\"1\">\n"
        xml += "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n"
        xml += "</cellStyleXfs>\n"

        // Cell formats (cellXfs) - the main styles array
        xml += generateCellStylesXML(stylesData.cellStyles)

        // Cell styles - required element
        xml += "<cellStyles count=\"1\">\n"
        xml += "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n"
        xml += "</cellStyles>\n"

        xml += "</styleSheet>"

        return xml
    }

    // MARK: - Private Helpers

    private func generateNumberFormatsXML(_ formats: [NumberFormat]) -> String {
        var xml = "<numFmts count=\"\(formats.count)\">\n"
        for format in formats {
            xml += "<numFmt numFmtId=\"\(format.formatId)\""
            if let code = format.formatCode {
                xml += " formatCode=\"\(xmlEscape(code))\""
            }
            xml += "/>\n"
        }
        xml += "</numFmts>\n"
        return xml
    }

    private func generateFontsXML(_ fonts: [Font]) -> String {
        var xml = "<fonts count=\"\(fonts.count)\">\n"
        for font in fonts {
            xml += "<font>\n"

            if font.bold {
                xml += "<b/>\n"
            }
            if font.italic {
                xml += "<i/>\n"
            }
            if font.strikethrough {
                xml += "<strike/>\n"
            }
            if let underline = font.underline {
                xml += "<u val=\"\(underline.rawValue)\"/>\n"
            }
            if let size = font.size {
                xml += "<sz val=\"\(size)\"/>\n"
            }
            if let color = font.color {
                if color.hasPrefix("theme:") {
                    let theme = String(color.dropFirst(6))
                    xml += "<color theme=\"\(theme)\"/>\n"
                } else {
                    xml += "<color rgb=\"\(color)\"/>\n"
                }
            }
            if let name = font.name {
                xml += "<name val=\"\(xmlEscape(name))\"/>\n"
            }
            if let family = font.family {
                xml += "<family val=\"\(family)\"/>\n"
            }
            if let charset = font.charset {
                xml += "<charset val=\"\(charset)\"/>\n"
            }

            xml += "</font>\n"
        }
        xml += "</fonts>\n"
        return xml
    }

    private func generateFillsXML(_ fills: [Fill]) -> String {
        var xml = "<fills count=\"\(fills.count)\">\n"
        for fill in fills {
            xml += "<fill>\n"
            xml += "<patternFill patternType=\"\(fill.patternType.rawValue)\""

            if fill.patternType == .none {
                xml += "/>\n"
            } else {
                xml += ">\n"

                if let fgColor = fill.foregroundColor {
                    if fgColor.hasPrefix("theme:") {
                        let theme = String(fgColor.dropFirst(6))
                        xml += "<fgColor theme=\"\(theme)\"/>\n"
                    } else {
                        xml += "<fgColor rgb=\"\(fgColor)\"/>\n"
                    }
                }

                if let bgColor = fill.backgroundColor {
                    if bgColor.hasPrefix("theme:") {
                        let theme = String(bgColor.dropFirst(6))
                        xml += "<bgColor theme=\"\(theme)\"/>\n"
                    } else {
                        xml += "<bgColor rgb=\"\(bgColor)\"/>\n"
                    }
                }

                xml += "</patternFill>\n"
            }

            xml += "</fill>\n"
        }
        xml += "</fills>\n"
        return xml
    }

    private func generateBordersXML(_ borders: [Border]) -> String {
        var xml = "<borders count=\"\(borders.count)\">\n"
        for border in borders {
            xml += "<border"
            if border.diagonalUp {
                xml += " diagonalUp=\"1\""
            }
            if border.diagonalDown {
                xml += " diagonalDown=\"1\""
            }
            xml += ">\n"

            xml += generateBorderEdgeXML("left", border.left)
            xml += generateBorderEdgeXML("right", border.right)
            xml += generateBorderEdgeXML("top", border.top)
            xml += generateBorderEdgeXML("bottom", border.bottom)
            xml += generateBorderEdgeXML("diagonal", border.diagonal)

            xml += "</border>\n"
        }
        xml += "</borders>\n"
        return xml
    }

    private func generateBorderEdgeXML(_ position: String, _ edge: BorderEdge?) -> String {
        guard let edge = edge else {
            return "<\(position)/>\n"
        }

        var xml = "<\(position) style=\"\(edge.style.rawValue)\">\n"
        if let color = edge.color {
            if color.hasPrefix("theme:") {
                let theme = String(color.dropFirst(6))
                xml += "<color theme=\"\(theme)\"/>\n"
            } else {
                xml += "<color rgb=\"\(color)\"/>\n"
            }
        }
        xml += "</\(position)>\n"
        return xml
    }

    private func generateCellStylesXML(_ cellStyles: [CellStyle]) -> String {
        var xml = "<cellXfs count=\"\(cellStyles.count)\">\n"
        for style in cellStyles {
            xml += "<xf"

            if let fontId = style.fontId {
                xml += " fontId=\"\(fontId)\""
            } else {
                xml += " fontId=\"0\""
            }

            if let fillId = style.fillId {
                xml += " fillId=\"\(fillId)\""
            } else {
                xml += " fillId=\"0\""
            }

            if let borderId = style.borderId {
                xml += " borderId=\"\(borderId)\""
            } else {
                xml += " borderId=\"0\""
            }

            if let numFmtId = style.numberFormatId {
                xml += " numFmtId=\"\(numFmtId)\""
            } else {
                xml += " numFmtId=\"0\""
            }

            xml += " xfId=\"0\""

            if style.applyFont {
                xml += " applyFont=\"1\""
            }
            if style.applyFill {
                xml += " applyFill=\"1\""
            }
            if style.applyBorder {
                xml += " applyBorder=\"1\""
            }
            if style.applyNumberFormat {
                xml += " applyNumberFormat=\"1\""
            }
            if style.applyAlignment {
                xml += " applyAlignment=\"1\""
            }

            // Check if we need alignment element
            let hasAlignment = style.horizontalAlignment != nil ||
                             style.verticalAlignment != nil ||
                             style.wrapText ||
                             style.textRotation != nil ||
                             style.indent != nil

            if hasAlignment {
                xml += ">\n"
                xml += "<alignment"

                if let horizontal = style.horizontalAlignment {
                    xml += " horizontal=\"\(horizontal.rawValue)\""
                }
                if let vertical = style.verticalAlignment {
                    xml += " vertical=\"\(vertical.rawValue)\""
                }
                if style.wrapText {
                    xml += " wrapText=\"1\""
                }
                if let rotation = style.textRotation {
                    xml += " textRotation=\"\(rotation)\""
                }
                if let indent = style.indent {
                    xml += " indent=\"\(indent)\""
                }

                xml += "/>\n"
                xml += "</xf>\n"
            } else {
                xml += "/>\n"
            }
        }
        xml += "</cellXfs>\n"
        return xml
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
