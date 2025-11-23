# Excel Conditional Formatting Research & Implementation Guide

**Date**: November 23, 2024  
**Project**: ApachePOISwift - Excel Library for Swift  
**Research Scope**: XML structure, rule types, and implementation strategy for conditional formatting

## Executive Summary

This research document details Excel conditional formatting implementation based on:
1. ECMA-376 (ISO/IEC 29500) specification analysis
2. Real-world examples from Marbar template (706 conditional formatting blocks, 716 rules)
3. Microsoft Learn official documentation
4. Apache POI (Java) reference implementation patterns

**Key Finding**: The Marbar template uses primarily **expression-based rules** (95%) and **data bars** (5%). No color scales, icon sets, or other advanced types are currently used, but full support should be implemented for production readiness.

---

## 1. Excel File Structure: Where Conditional Formatting Stored

### Location in .xlsx/.xlsm Archive

```
my-workbook.xlsm/
└── xl/
    └── worksheets/
        ├── sheet1.xml          <-- CONDITIONAL FORMATTING HERE
        ├── sheet2.xml
        └── ...
```

Conditional formatting is stored **directly in worksheet XML files** at the end of the `<sheetData>` element, before the closing `</worksheet>` tag.

### Position in Worksheet XML

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ...>
  <sheetData>
    <row r="1">
      <c r="A1"><v>123</v></c>
    </row>
    <!-- ... more rows ... -->
  </sheetData>
  
  <!-- CONDITIONAL FORMATTING GOES HERE -->
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
  
</worksheet>
```

### Important Notes

- **Multiple formats**: A worksheet can have multiple `<conditionalFormatting>` blocks
- **Ranges**: The `sqref` (sequence reference) attribute specifies affected cell range
- **Priority**: Rules are evaluated in priority order (1 = highest)
- **Format Application**: The `dxfId` attribute references styles in `xl/styles.xml` (Differential XF)
- **No Direct Cell Modification**: Conditional formatting doesn't modify cell values; it's display-only

---

## 2. Conditional Formatting XML Structure

### Root Element: conditionalFormatting

```xml
<conditionalFormatting 
    sqref="A1:A10"              <!-- Cell range(s) affected -->
    pivot="0">                  <!-- (Optional) Part of pivot table -->
  <!-- cfRule elements -->
</conditionalFormatting>
```

**Attributes**:
- `sqref`: Space-separated or colon-separated cell ranges
  - Single range: `A1:A10`
  - Multiple ranges: `A1:A10 C5:C15 E1:E5`
- `pivot`: (Optional) Boolean, indicates if this applies to pivot table

### Rule Element: cfRule

```xml
<cfRule 
    type="cellIs"               <!-- Rule type: cellIs, expression, colorScale, dataBar, iconSet, top10, etc. -->
    dxfId="5"                   <!-- Index into dxf array (differential formats) in styles.xml -->
    priority="1"                <!-- Evaluation order (1 is highest priority) -->
    operator="greaterThan"      <!-- Comparison operator (for cellIs type) -->
    text="value"                <!-- (Optional) Text to match (for containsText type) -->
    timePeriod="today"          <!-- (Optional) Time period (for timePeriod type) -->
    rank="10"                   <!-- (Optional) Rank for top10 type -->
    percent="0"                 <!-- (Optional) Percentage for top10 type -->
    bottom="0"                  <!-- (Optional) Bottom instead of top for top10 type -->
    aboveAverage="0"            <!-- (Optional) Above/below average indicator -->
    stopIfTrue="0">             <!-- (Optional) Stop evaluating rules if this matches -->
  <!-- Rule-specific content -->
</cfRule>
```

**Key Attributes**:
- `type`: Determines the rule format (see section 3)
- `dxfId`: Links to differential format in styles.xml
- `priority`: Numeric priority for rule evaluation order
- `stopIfTrue`: If 1, stops evaluating further rules for matching cells

---

## 3. Conditional Formatting Rule Types

### 3.1 expression (Formula-Based) - MOST COMMON

Used in: **95% of Marbar template rules**

```xml
<conditionalFormatting sqref="F6:CF6">
  <cfRule type="expression" dxfId="1107" priority="18">
    <formula>F$2>0</formula>
  </cfRule>
</conditionalFormatting>
```

**Characteristics**:
- Evaluates an arbitrary formula
- Formula can reference the cell being formatted using relative references
- Cell references can use absolute anchors: `$A$1`, `$A1`, `A$1`, `A1`
- Formula returns TRUE (apply format) or FALSE (don't apply)

**Real Examples from Marbar**:
```xml
<!-- Example 1: Check if specific column has data -->
<formula>E$2>0</formula>

<!-- Example 2: Check if cell is empty -->
<formula>$F135=""</formula>

<!-- Example 3: Complex formula -->
<formula>E$2</formula>  <!-- evaluates truthiness of cell E2 -->
```

### 3.2 cellIs (Value Comparison)

```xml
<conditionalFormatting sqref="E3:E9">
  <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
    <formula>0.5</formula>
  </cfRule>
</conditionalFormatting>
```

**Valid Operators**:
- `equal` - =
- `notEqual` - !=
- `greaterThan` - >
- `lessThan` - <
- `greaterThanOrEqual` - >=
- `lessThanOrEqual` - <=
- `between` - value is between two numbers
- `notBetween` - value is not between two numbers
- `beginsWith` - text begins with
- `endsWith` - text ends with
- `containsText` - text contains
- `doesNotContain` - text doesn't contain

**Note**: Multiple formulas can be used for `between`/`notBetween`:
```xml
<cfRule type="cellIs" operator="between">
  <formula>10</formula>
  <formula>20</formula>
</cfRule>
```

### 3.3 dataBar (Progress Bars) - USED IN MARBAR

Used in: **~5% of Marbar template rules**

```xml
<conditionalFormatting sqref="Q226:Q236">
  <cfRule type="dataBar" priority="42">
    <dataBar showValue="0">
      <cfvo type="num" val="2"/>
      <cfvo type="num" val="4.2"/>
      <color theme="2" tint="-0.249977111117893"/>
    </dataBar>
    <extLst>
      <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
        <x14:id>{3CBC2E01-B88E-4ADD-8FC4-9786FCF518DD}</x14:id>
      </ext>
    </extLst>
  </cfRule>
</conditionalFormatting>
```

**Elements**:
- `<dataBar>`: Container for data bar settings
  - `showValue`: "0" = hide value, "1" = show value
- `<cfvo>` (Conditional Formatting Value Object): Min/max points
  - `type`: "num" (number), "percent", "percentile", "min", "max"
  - `val`: The value (ignored for min/max types)
- `<color>`: Bar color
  - `theme`: Theme color index
  - `tint`: Tint modifier (range: -1 to 1)
  - OR `rgb`: Direct RGB color (e.g., "FF638EC6")

**Color Examples from Marbar**:
```xml
<color theme="2" tint="-0.499984740745262"/>
<color theme="2" tint="-0.249977111117893"/>
<color theme="2" tint="-0.749992370372631"/>
```

### 3.4 colorScale (Color Gradient)

```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="percentile" val="50"/>
      <cfvo type="max"/>
      <color rgb="F8696B"/>      <!-- Red -->
      <color rgb="FFEB84"/>      <!-- Yellow -->
      <color rgb="63BE7B"/>      <!-- Green -->
    </colorScale>
  </cfRule>
</conditionalFormatting>
```

**Variants**:
- **2-color scale**: 2 cfvo + 2 colors (min to max)
- **3-color scale**: 3 cfvo + 3 colors (min, middle, max)

**cfvo Types**:
- `num`: Specific number
- `percent`: Percentage of min-max range
- `percentile`: Percentile of data
- `min`: Minimum value in range
- `max`: Maximum value in range
- `formula`: Custom formula

### 3.5 iconSet (Icons/Arrows/Indicators)

```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="iconSet" priority="1">
    <iconSet iconSet="3Arrows" reverse="0">
      <cfvo type="percentile" val="0"/>
      <cfvo type="percentile" val="33"/>
      <cfvo type="percentile" val="67"/>
    </iconSet>
  </cfRule>
</conditionalFormatting>
```

**Available Icon Sets**:
- **3-point sets**: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2
- **4-point sets**: 4Arrows, 4ArrowsGray, 4RedToBlack, 4TrafficLights
- **5-point sets**: 5Arrows, 5ArrowsGray, 5Quarters, 5Rating

### 3.6 top10 (Top/Bottom Values)

```xml
<conditionalFormatting sqref="C3:C8">
  <cfRule type="top10" dxfId="1" priority="3" rank="2"/>
</conditionalFormatting>
```

**Attributes**:
- `rank`: Number of items (default 10)
- `percent`: "1" = treat rank as percentage
- `bottom`: "1" = bottom values instead of top

**Examples**:
```xml
<cfRule type="top10" rank="5"/>              <!-- Top 5 values -->
<cfRule type="top10" rank="10" percent="1"/> <!-- Top 10% -->
<cfRule type="top10" rank="5" bottom="1"/>   <!-- Bottom 5 values -->
```

### 3.7 Other Rule Types

**aboveAverage**:
```xml
<cfRule type="aboveAverage" dxfId="1" priority="1"/>
<!-- No formula needed, uses cell range to calculate average -->
```

**uniqueValues**:
```xml
<cfRule type="uniqueValues" dxfId="1" priority="1"/>
<!-- Highlights unique values in the range -->
```

**duplicateValues**:
```xml
<cfRule type="duplicateValues" dxfId="1" priority="1"/>
<!-- Highlights duplicate values in the range -->
```

**containsText** (cellIs variant):
```xml
<cfRule type="cellIs" operator="containsText" text="error" dxfId="1" priority="1">
  <formula>SEARCH("error",A1)</formula>
</cfRule>
```

**timePeriod**:
```xml
<cfRule type="timePeriod" timePeriod="today" dxfId="1" priority="1"/>
```

**timePeriod Values**:
- `today`, `yesterday`, `tomorrow`
- `last7Days`, `lastMonth`, `nextMonth`, `nextWeek`, `thisMonth`, `thisWeek`, `thisYear`

---

## 4. Differential Formats (dxfId in styles.xml)

Conditional formatting rules reference a **dxfId** (differential XF index) from `styles.xml`.

### Location in styles.xml

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts>...</fonts>
  <fills>...</fills>
  <borders>...</borders>
  <cellStyleXfs>...</cellStyleXfs>
  <cellXfs>...</cellXfs>
  
  <!-- DIFFERENTIAL FORMATS FOR CONDITIONAL FORMATTING -->
  <dxfs count="1000">
    <dxf>
      <font>
        <b/>                    <!-- Bold -->
        <color rgb="FF0000"/>   <!-- Red text -->
      </font>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFFF00"/>  <!-- Yellow background -->
        </patternFill>
      </fill>
    </dxf>
    <!-- ... many more dxf elements (0-indexed) -->
  </dxfs>
  
</styleSheet>
```

### dxf (Differential Format) Structure

```xml
<dxf>
  <font>
    <b/>                        <!-- Bold -->
    <i/>                        <!-- Italic -->
    <color rgb="FFFF0000"/>     <!-- Color (ARGB) -->
  </font>
  <numFmt numFmtId="164" formatCode="0.00"/>  <!-- Number format -->
  <fill>
    <patternFill patternType="solid">
      <fgColor rgb="FFFF0000"/> <!-- Foreground (text) color -->
      <bgColor rgb="FFFFFF00"/> <!-- Background color -->
    </patternFill>
  </fill>
  <border>
    <left style="thin"><color rgb="FFFF0000"/></left>
    <right style="thin"><color rgb="FFFF0000"/></right>
    <top style="thin"><color rgb="FFFF0000"/></top>
    <bottom style="thin"><color rgb="FFFF0000"/></bottom>
    <diagonal style="thin"><color rgb="FFFF0000"/></diagonal>
  </border>
  <alignment horizontal="center" vertical="center" wrapText="1"/>
</dxf>
```

**What dxf Can Contain**:
- Font styling (bold, italic, size, color, underline, strikethrough)
- Fill/background color
- Border styling
- Number formatting
- Cell alignment
- Text effects

**Note**: Regular cell styles use `cellXfs`, but conditional formatting uses `dxfs` (differential formats). They serve the same purpose but are indexed separately.

---

## 5. Real-World Analysis: Marbar Template

### Statistics

- **Total Conditional Formatting Blocks**: 706
- **Total Rules**: 716
- **Sheets with Conditional Formatting**: 20 of 26 sheets
- **Sheets with Most Rules**: sheet23.xml (148 blocks, 156 rules)

### Rule Type Distribution

| Type | Count | Percentage | Sheets |
|------|-------|-----------|--------|
| expression | 696 | 97.2% | 20 sheets |
| dataBar | 20 | 2.8% | 2 sheets |
| **Total** | **716** | **100%** | - |

### Expression Formula Patterns Found

```excel
E$2>0              <!-- Check column header exists -->
$F135=""           <!-- Check if cell is empty -->
E2                 <!-- Check if cell value is truthy -->
IF(E2=0,"-",E3+1)  <!-- Complex nested logic -->
```

### Data Bar Examples

```xml
<!-- Example 1: Min=2, Max=4.2, Dark blue tint -->
<dataBar showValue="0">
  <cfvo type="num" val="2"/>
  <cfvo type="num" val="4.2"/>
  <color theme="2" tint="-0.249977111117893"/>
</dataBar>

<!-- Example 2: Min=0, Max=1, Very dark blue tint -->
<dataBar showValue="0">
  <cfvo type="num" val="0"/>
  <cfvo type="num" val="1"/>
  <color theme="2" tint="-0.749992370372631"/>
</dataBar>
```

### dxfId Pattern

- Range: dxfId 25 to 1107
- Marbar template uses 1,083 unique dxfId values
- Each conditional formatting rule applies different styling

---

## 6. Implementation Architecture for ApachePOISwift

### Proposed File Structure

```
ApachePOISwift/
├── Sources/ApachePOISwift/
│   ├── ConditionalFormatting/
│   │   ├── ConditionalFormat.swift         # Main conditional format model
│   │   ├── ConditionalFormattingRule.swift # cfRule abstraction
│   │   ├── RuleType.swift                 # Enum for rule types
│   │   ├── Operator.swift                 # Enum for cellIs operators
│   │   ├── DataBar.swift                  # DataBar-specific model
│   │   ├── ColorScale.swift               # ColorScale-specific model
│   │   ├── IconSet.swift                  # IconSet-specific model
│   │   ├── ConditionalFormattingParser.swift    # XML parsing
│   │   └── ConditionalFormattingWriter.swift    # XML generation
│   └── XML/
│       └── SheetXMLParser.swift           # (UPDATE) Add CF parsing
│
└── Tests/ApachePOISwiftTests/
    └── ConditionalFormattingTests.swift
```

### Core Model: ConditionalFormat

```swift
/// Represents a <conditionalFormatting> element
public struct ConditionalFormat {
    /// Cell ranges affected (e.g., "A1:A10 C5:C15")
    public let sqref: String
    
    /// Rules to apply (in priority order)
    public var rules: [ConditionalFormattingRule]
    
    /// Whether this applies to pivot table
    public let pivot: Bool
    
    public init(sqref: String, rules: [ConditionalFormattingRule], pivot: Bool = false) {
        self.sqref = sqref
        self.rules = rules.sorted { $0.priority < $1.priority }
        self.pivot = pivot
    }
}
```

### Core Model: ConditionalFormattingRule

```swift
/// Represents a <cfRule> element
public struct ConditionalFormattingRule {
    /// Type of rule (expression, cellIs, dataBar, etc.)
    public let type: RuleType
    
    /// Index into differential formats in styles.xml
    public let dxfId: Int?
    
    /// Evaluation priority (1 is highest)
    public let priority: Int
    
    /// Stop evaluating further rules if this matches
    public let stopIfTrue: Bool
    
    /// Rule-specific configuration
    public let configuration: RuleConfiguration
}
```

### RuleType Enum

```swift
public enum RuleType: String {
    case expression = "expression"
    case cellIs = "cellIs"
    case colorScale = "colorScale"
    case dataBar = "dataBar"
    case iconSet = "iconSet"
    case top10 = "top10"
    case aboveAverage = "aboveAverage"
    case uniqueValues = "uniqueValues"
    case duplicateValues = "duplicateValues"
    case timePeriod = "timePeriod"
}
```

### RuleConfiguration (Protocol-based)

```swift
public protocol RuleConfiguration {
    var ruleType: RuleType { get }
}

/// Expression-based rule (custom formula)
public struct ExpressionRule: RuleConfiguration {
    public let formula: String
    public let ruleType: RuleType = .expression
}

/// Value comparison rule
public struct CellIsRule: RuleConfiguration {
    public let `operator`: CellIsOperator
    public let formulas: [String]  // 1 for most operators, 2 for between
    public let ruleType: RuleType = .cellIs
}

/// Data bar visualization
public struct DataBarRule: RuleConfiguration {
    public let minValue: DataBarValue
    public let maxValue: DataBarValue
    public let color: ExcelColor
    public let showValue: Bool
    public let ruleType: RuleType = .dataBar
}

/// Color gradient visualization
public struct ColorScaleRule: RuleConfiguration {
    public let values: [ColorScalePoint]  // 2 or 3 points
    public let ruleType: RuleType = .colorScale
}

/// Icon set visualization
public struct IconSetRule: RuleConfiguration {
    public let iconSet: IconSetType
    public let values: [DataBarValue]
    public let reverse: Bool
    public let ruleType: RuleType = .iconSet
}

/// Top/Bottom N values
public struct Top10Rule: RuleConfiguration {
    public let rank: Int
    public let percent: Bool
    public let bottom: Bool
    public let ruleType: RuleType = .top10
}

/// Above/below average
public struct AboveAverageRule: RuleConfiguration {
    public let bottom: Bool
    public let equalAverage: Bool
    public let ruleType: RuleType = .aboveAverage
}

/// Time-based rule
public struct TimePeriodRule: RuleConfiguration {
    public let timePeriod: TimePeriod
    public let ruleType: RuleType = .timePeriod
}

/// Unique values
public struct UniqueValuesRule: RuleConfiguration {
    public let ruleType: RuleType = .uniqueValues
}

/// Duplicate values
public struct DuplicateValuesRule: RuleConfiguration {
    public let ruleType: RuleType = .duplicateValues
}
```

### Supporting Types

```swift
/// Conditional Formatting Value Object types
public enum DataBarValueType: String {
    case num = "num"
    case percent = "percent"
    case percentile = "percentile"
    case min = "min"
    case max = "max"
    case formula = "formula"
}

public struct DataBarValue {
    public let type: DataBarValueType
    public let value: String?  // nil for min/max
}

public enum CellIsOperator: String {
    case equal = "equal"
    case notEqual = "notEqual"
    case greaterThan = "greaterThan"
    case lessThan = "lessThan"
    case greaterThanOrEqual = "greaterThanOrEqual"
    case lessThanOrEqual = "lessThanOrEqual"
    case between = "between"
    case notBetween = "notBetween"
    case containsText = "containsText"
    case doesNotContain = "doesNotContain"
    case beginsWith = "beginsWith"
    case endsWith = "endsWith"
}

public enum TimePeriod: String {
    case today = "today"
    case yesterday = "yesterday"
    case tomorrow = "tomorrow"
    case last7Days = "last7Days"
    case lastMonth = "lastMonth"
    case nextMonth = "nextMonth"
    case nextWeek = "nextWeek"
    case thisMonth = "thisMonth"
    case thisWeek = "thisWeek"
    case thisYear = "thisYear"
}

public enum IconSetType: String {
    // 3-point
    case threeArrows = "3Arrows"
    case threeArrowsGray = "3ArrowsGray"
    case threeFlags = "3Flags"
    case threeTrafficLights1 = "3TrafficLights1"
    case threeTrafficLights2 = "3TrafficLights2"
    case threeSymbols = "3Symbols"
    case threeSymbols2 = "3Symbols2"
    
    // 4-point
    case fourArrows = "4Arrows"
    case fourArrowsGray = "4ArrowsGray"
    case fourRedToBlack = "4RedToBlack"
    case fourTrafficLights = "4TrafficLights"
    
    // 5-point
    case fiveArrows = "5Arrows"
    case fiveArrowsGray = "5ArrowsGray"
    case fiveQuarters = "5Quarters"
    case fiveRating = "5Rating"
}

public struct ColorScalePoint {
    public let value: DataBarValue
    public let color: ExcelColor
}
```

---

## 7. XML Parsing & Writing

### Parsing Flow

```
SheetXMLParser
  └── readConditionalFormatting()
      └── for each <conditionalFormatting>
          ├── extract sqref
          ├── extract pivot (optional)
          └── for each <cfRule>
              ├── determine type
              ├── extract base attributes (dxfId, priority, stopIfTrue)
              └── parse type-specific content
                  ├── <formula> → ExpressionRule
                  ├── <dataBar> → DataBarRule
                  ├── <colorScale> → ColorScaleRule
                  ├── <iconSet> → IconSetRule
                  └── operator + formula → CellIsRule
```

### Writing Flow

```
ConditionalFormattingWriter
  └── writeConditionalFormatting(formats: [ConditionalFormat]) -> String
      └── for each ConditionalFormat
          ├── <conditionalFormatting sqref="..." pivot="...">
          └── for each rule (sorted by priority)
              ├── <cfRule type="..." dxfId="..." priority="..." ...>
              ├── write rule-specific content
              │   ├── <formula> for expression
              │   ├── <dataBar> for data bars
              │   ├── <colorScale> for color scales
              │   └── <iconSet> for icon sets
              └── </cfRule>
```

### Parser Pseudocode

```swift
func parseConditionalFormatting(element: XMLElement) throws -> ConditionalFormat {
    guard let sqref = element.attribute(forName: "sqref") else {
        throw ExcelError.parsingError("Missing sqref in conditionalFormatting")
    }
    
    let pivot = element.attribute(forName: "pivot")?.stringValue == "1"
    var rules: [ConditionalFormattingRule] = []
    
    for ruleElement in element.elements(forName: "cfRule") {
        let rule = try parseRule(ruleElement)
        rules.append(rule)
    }
    
    return ConditionalFormat(sqref: sqref.stringValue, rules: rules, pivot: pivot)
}

func parseRule(_ element: XMLElement) throws -> ConditionalFormattingRule {
    guard let typeStr = element.attribute(forName: "type")?.stringValue,
          let type = RuleType(rawValue: typeStr) else {
        throw ExcelError.parsingError("Invalid or missing rule type")
    }
    
    let dxfId = element.attribute(forName: "dxfId").flatMap { Int($0.stringValue) }
    let priority = Int(element.attribute(forName: "priority")?.stringValue ?? "0") ?? 0
    let stopIfTrue = element.attribute(forName: "stopIfTrue")?.stringValue == "1"
    
    let configuration: RuleConfiguration
    
    switch type {
    case .expression:
        configuration = try parseExpressionRule(element)
    case .cellIs:
        configuration = try parseCellIsRule(element)
    case .dataBar:
        configuration = try parseDataBarRule(element)
    case .colorScale:
        configuration = try parseColorScaleRule(element)
    case .iconSet:
        configuration = try parseIconSetRule(element)
    case .top10:
        configuration = try parseTop10Rule(element)
    // ... other cases
    }
    
    return ConditionalFormattingRule(
        type: type,
        dxfId: dxfId,
        priority: priority,
        stopIfTrue: stopIfTrue,
        configuration: configuration
    )
}
```

---

## 8. Integration with Existing ApachePOISwift Code

### Updates to SheetXMLParser

Add to `SheetXMLParser.swift`:

```swift
class SheetXMLParser: NSObject, XMLParserDelegate {
    // ... existing code ...
    
    private var conditionalFormats: [ConditionalFormat] = []
    
    func parse(data: Data) throws -> SheetData {
        // ... existing parsing code ...
        
        let parser = XMLParser(data: data)
        parser.delegate = self
        
        guard parser.parse() else {
            throw ExcelError.parsingError("Sheet: \(parser.parserError?.localizedDescription ?? "Unknown error")")
        }
        
        // Return extended SheetData with conditional formatting
        return SheetData(
            cells: cells,
            mergedCells: mergedCells,
            conditionalFormats: conditionalFormats
        )
    }
    
    func parser(_ parser: XMLParser, 
                didStartElement elementName: String,
                namespaceURI: String?,
                qualifiedName qName: String?,
                attributes attributeDict: [String: String] = [:]) {
        
        if elementName == "conditionalFormatting" {
            // Parse conditional formatting block
            let format = try? parseConditionalFormattingElement(attributeDict)
            // Handle parsing
        }
        
        // ... existing code ...
    }
}
```

### Updates to SheetData

```swift
struct SheetData {
    let cells: [String: CellData]
    let mergedCells: [String]
    let conditionalFormats: [ConditionalFormat] = []  // NEW
}
```

### Updates to ExcelSheet

```swift
public class ExcelSheet {
    // ... existing code ...
    
    public var conditionalFormats: [ConditionalFormat] = []
    
    /// Add conditional formatting rule
    public func addConditionalFormat(_ format: ConditionalFormat) {
        conditionalFormats.append(format)
    }
    
    /// Remove conditional formatting from range
    public func removeConditionalFormats(sqref: String) {
        conditionalFormats.removeAll { $0.sqref.contains(sqref) }
    }
}
```

---

## 9. Testing Strategy

### Unit Tests: Parsing

```swift
func testParseExpressionRule() throws {
    let xml = """
    <conditionalFormatting sqref="A1:A10">
      <cfRule type="expression" dxfId="5" priority="1">
        <formula>A1>100</formula>
      </cfRule>
    </conditionalFormatting>
    """
    
    let format = try ConditionalFormattingParser.parse(xml)
    
    XCTAssertEqual(format.sqref, "A1:A10")
    XCTAssertEqual(format.rules.count, 1)
    
    let rule = format.rules[0]
    XCTAssertEqual(rule.type, .expression)
    XCTAssertEqual(rule.dxfId, 5)
    XCTAssertEqual(rule.priority, 1)
    
    if let exprRule = rule.configuration as? ExpressionRule {
        XCTAssertEqual(exprRule.formula, "A1>100")
    } else {
        XCTFail("Expected ExpressionRule")
    }
}

func testParseDataBarRule() throws {
    let xml = """
    <conditionalFormatting sqref="Q226:Q236">
      <cfRule type="dataBar" priority="42">
        <dataBar showValue="0">
          <cfvo type="num" val="2"/>
          <cfvo type="num" val="4.2"/>
          <color theme="2" tint="-0.25"/>
        </dataBar>
      </cfRule>
    </conditionalFormatting>
    """
    
    let format = try ConditionalFormattingParser.parse(xml)
    let rule = format.rules[0]
    
    if let dataBar = rule.configuration as? DataBarRule {
        XCTAssertEqual(dataBar.minValue.type, .num)
        XCTAssertEqual(dataBar.maxValue.value, "4.2")
    } else {
        XCTFail("Expected DataBarRule")
    }
}
```

### Integration Tests: Marbar Template

```swift
func testMarbarTemplateConditionalFormatting() throws {
    let workbook = try ExcelWorkbook(fileURL: marbarTemplateURL)
    
    // Sheet5 should have 12 conditionalFormatting blocks
    let sheet5 = try workbook.sheet(at: 4)
    XCTAssertEqual(sheet5.conditionalFormats.count, 12)
    
    // Verify first rule
    let format = sheet5.conditionalFormats[0]
    XCTAssertEqual(format.sqref, "F6:CF6")
    
    let rule = format.rules[0]
    XCTAssertEqual(rule.type, .expression)
    XCTAssertEqual(rule.priority, 18)
    XCTAssertEqual(rule.dxfId, 1107)
}

func testRoundTripConditionalFormatting() throws {
    // Open Marbar template
    let workbook = try ExcelWorkbook(fileURL: marbarTemplateURL)
    let originalSheet = try workbook.sheet(at: 0)
    let originalFormatCount = originalSheet.conditionalFormats.count
    
    // Save to new file
    let outputURL = tempDirectory.appendingPathComponent("marbar_output.xlsm")
    try workbook.save(to: outputURL)
    
    // Reopen and verify
    let reopened = try ExcelWorkbook(fileURL: outputURL)
    let reopenedSheet = try reopened.sheet(at: 0)
    
    XCTAssertEqual(reopenedSheet.conditionalFormats.count, originalFormatCount)
    
    // Verify specific rules are preserved
    let originalFormats = Set(originalSheet.conditionalFormats.map { $0.sqref })
    let reopenedFormats = Set(reopenedSheet.conditionalFormats.map { $0.sqref })
    
    XCTAssertEqual(originalFormats, reopenedFormats)
}
```

---

## 10. Usage Examples

### Example 1: Create Expression-Based Conditional Format

```swift
import ApachePOISwift

// Create conditional format: highlight if column header has data
let rule = ConditionalFormattingRule(
    type: .expression,
    dxfId: 5,      // References dxf[5] in styles.xml
    priority: 1,
    stopIfTrue: false,
    configuration: ExpressionRule(formula: "F$2>0")
)

let format = ConditionalFormat(
    sqref: "F6:CF6",
    rules: [rule]
)

sheet.addConditionalFormat(format)
```

### Example 2: Create Data Bar

```swift
let minValue = DataBarValue(type: .num, value: "0")
let maxValue = DataBarValue(type: .num, value: "100")
let color = ExcelColor(theme: 2, tint: "-0.25")

let dataBarConfig = DataBarRule(
    minValue: minValue,
    maxValue: maxValue,
    color: color,
    showValue: true
)

let rule = ConditionalFormattingRule(
    type: .dataBar,
    dxfId: nil,
    priority: 1,
    stopIfTrue: false,
    configuration: dataBarConfig
)

let format = ConditionalFormat(sqref: "A1:A100", rules: [rule])
sheet.addConditionalFormat(format)
```

### Example 3: Create Color Scale

```swift
let points = [
    ColorScalePoint(
        value: DataBarValue(type: .min, value: nil),
        color: ExcelColor(rgb: "F8696B")  // Red
    ),
    ColorScalePoint(
        value: DataBarValue(type: .percentile, value: "50"),
        color: ExcelColor(rgb: "FFEB84")  // Yellow
    ),
    ColorScalePoint(
        value: DataBarValue(type: .max, value: nil),
        color: ExcelColor(rgb: "63BE7B")  // Green
    )
]

let colorScaleConfig = ColorScaleRule(values: points)

let rule = ConditionalFormattingRule(
    type: .colorScale,
    dxfId: nil,
    priority: 1,
    stopIfTrue: false,
    configuration: colorScaleConfig
)

let format = ConditionalFormat(sqref: "A1:A100", rules: [rule])
sheet.addConditionalFormat(format)
```

### Example 4: Create Top 10 Rule

```swift
let top10Config = Top10Rule(
    rank: 10,
    percent: false,
    bottom: false
)

let rule = ConditionalFormattingRule(
    type: .top10,
    dxfId: 3,
    priority: 1,
    stopIfTrue: false,
    configuration: top10Config
)

let format = ConditionalFormat(sqref: "A1:A100", rules: [rule])
sheet.addConditionalFormat(format)
```

---

## 11. Phase Implementation Plan

### Phase 1: Infrastructure (Week 1)
- Create ConditionalFormatting models (types, enums, structs)
- Implement parser for expression and cellIs rules
- Implement parser for dataBar rules (most common in Marbar)
- Add parsing to SheetXMLParser
- Add conditional formats to SheetData and ExcelSheet
- Unit tests for parsing

### Phase 2: Writing Support (Week 2)
- Implement ConditionalFormattingWriter
- Write conditionalFormatting XML generation
- Update WorkbookSaver to include CF in output
- Integration tests with Marbar template
- Round-trip tests (read → modify → write → read)

### Phase 3: Advanced Types (Week 3)
- Implement colorScale parser and writer
- Implement iconSet parser and writer
- Implement top10, aboveAverage, uniqueValues, duplicateValues
- Implement timePeriod rules
- Comprehensive tests

### Phase 4: Polish & Optimization (Week 4)
- Performance optimization
- Differential format (dxfId) validation
- Edge case handling
- Documentation and examples
- Integration with Solids Control app

---

## 12. Key References

### Official Standards
- ECMA-376 (ISO/IEC 29500) - Office Open XML File Formats
  - Part 1: Fundamentals and Markup Language Reference
  - Section: SpreadsheetML - Conditional Formatting

### Microsoft Documentation
- [Working with Conditional Formatting](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting)

### Apache POI Reference
- XSSFConditionalFormatting: https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/usermodel/XSSFConditionalFormatting.html
- ConditionalFormattingRule: https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/ConditionalFormattingRule.html

### Real-World Data (Marbar Template)
- Location: `/Users/masterfem/SolidsControlNqn/ios/solidscontrolapp/solidscontrolapp/ExcelTemplate/marbar_template.xlsm`
- Statistics: 706 conditional formatting blocks, 716 rules
- Rule types: 97.2% expression, 2.8% dataBar
- dxfId range: 25-1107

---

## 13. Summary & Recommendations

### Key Takeaways

1. **Conditional formatting is stored in worksheet XML**, separate from cell data
2. **Two main rule types dominate Marbar template**: expression (97%) and dataBar (3%)
3. **dxfId links to differential formats** in styles.xml (distinct from regular cellXfs)
4. **Rule evaluation order matters**: priority attribute determines sequence
5. **Complex formulas are supported**: any formula that evaluates to TRUE/FALSE

### Implementation Priority

1. **MUST HAVE** (Phase 1-2):
   - Expression-based rules (formula evaluation)
   - DataBar rules (visualization)
   - Basic parser and writer

2. **SHOULD HAVE** (Phase 3):
   - ColorScale rules
   - Top10/bottom rules
   - IconSet rules

3. **NICE TO HAVE** (Phase 4):
   - AboveAverage, UniqueValues, DuplicateValues
   - TimePeriod rules
   - Advanced formula evaluation

### Compatibility Notes

- Excel 2007+ supports all rule types
- Excel 97-2003 doesn't support conditional formatting
- Some icon set types may not render in older Excel versions
- Color gradients (colorScale) require Excel 2007+
- DataBar visualization is consistent across Excel versions

---

**Research Completed**: November 23, 2024
**Next Step**: Begin Phase 1 implementation of ConditionalFormatting models and parser
