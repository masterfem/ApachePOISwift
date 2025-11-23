# Critical Bug: Excel OOXML cellXfs vs cellStyleXfs Parsing Issue

**Date**: November 22, 2024
**Severity**: Critical - Causes complete style corruption
**Status**: Fixed in commit 368e665
**Platforms**: Any Excel OOXML (.xlsx/.xlsm) parser implementation

## Executive Summary

When parsing Excel `styles.xml` files, failing to distinguish between `<cellStyleXfs>` and `<cellXfs>` sections causes all cell style indices to be shifted, resulting in cells displaying completely wrong fonts, colors, borders, and formatting after save/reload cycles.

## Keywords for Search

- Excel OOXML styles.xml parsing
- cellXfs cellStyleXfs difference
- Swift Excel parser wrong colors after reload
- Apache POI style index mismatch
- Excel XML cell formatting wrong after save
- styles.xml xf element parsing
- OpenXML cell styles corruption
- Excel fillId fontId incorrect after reload

## The Bug

### Symptoms

1. **Cell gets wrong colors after save/reload**
   - You set a cell's background to RED (`FFFF0000`)
   - Cell correctly shows red before saving
   - After save/reload, cell shows GREEN (`FF79A63B`) or other wrong color

2. **Style indices shift unexpectedly**
   - Cell has styleIndex=1254 which should reference fillId=27
   - After reload, styleIndex=1254 references fillId=30 (off by 3)
   - The written XML is correct, but parsing produces wrong mappings

3. **Fonts, borders, and all formatting gets mixed up**
   - Not limited to colors - affects all style properties
   - The shift is consistent (always off by N indices)

### Root Cause

Excel `styles.xml` has **two separate collections** of `<xf>` elements:

```xml
<styleSheet>
  <!-- ... -->

  <!-- 1. Cell STYLE Xfs (templates) - NOT used directly by cells -->
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>

  <!-- 2. Cell Xfs (actual cell styles) - cells reference THESE -->
  <cellXfs count="1255">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>  <!-- cellStyles[0] -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>  <!-- cellStyles[1] -->
    <!-- ... -->
    <xf numFmtId="0" fontId="27" fillId="0" borderId="0"/> <!-- cellStyles[1254] -->
  </cellXfs>
</styleSheet>
```

**The Problem**: Both sections use the same element name `<xf>`. A naive parser that just counts all `<xf>` elements will:

1. Count the 1 `<xf>` in `cellStyleXfs` as cellStyles[0]
2. Then count the first `<xf>` in `cellXfs` as cellStyles[1]

But cells actually reference `cellXfs` indices starting from 0!

### Example

**Written XML** (correct):
```xml
<cellStyleXfs count="1">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>  <!-- NOT indexed -->
</cellStyleXfs>

<cellXfs count="1255">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>  <!-- Index 0 -->
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>  <!-- Index 1 -->
  <!-- ... -->
  <xf numFmtId="0" fontId="0" fillId="27" borderId="0"/> <!-- Index 1254 -->
</cellXfs>
```

**Cell in sheet.xml**:
```xml
<c r="Z93" s="1254">  <!-- s="1254" references cellXfs[1254] -->
  <v>Red Background</v>
</c>
```

**Buggy Parser Result**: cellStyles array contains:
- Index 0: The cellStyleXfs element ❌
- Index 1: First cellXfs element (should be index 0) ❌
- Index 2: Second cellXfs element (should be index 1) ❌
- ...
- Index 1254: The 1254th cellXfs element (should be index 1253) ❌

When cell references style 1254, it gets cellXfs[1253] instead of cellXfs[1254]!

## The Fix

### Solution: Track Parser State

Add a flag to track which section you're currently parsing:

```swift
// Swift example
class StylesXMLParser: NSObject, XMLParserDelegate {
    private var cellStyleIndex = 0
    private var inCellXfs = false  // ← Add this flag

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        attributes: [String: String]
    ) {
        switch elementName {
        case "cellXfs":
            inCellXfs = true          // ← Set flag when entering cellXfs
            cellStyleIndex = 0        // ← Reset counter!

        case "xf":
            // Only process if we're in cellXfs (not cellStyleXfs)
            if inCellXfs {            // ← Check flag before processing
                let style = CellStyle(
                    index: cellStyleIndex,
                    fontId: attributes["fontId"],
                    fillId: attributes["fillId"],
                    // ...
                )
                cellStyles.append(style)
                cellStyleIndex += 1
            }
            // If not in cellXfs, ignore this <xf> element

        // ...
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String
    ) {
        if elementName == "cellXfs" {
            inCellXfs = false         // ← Reset flag
        }
    }
}
```

### Alternative Solutions

**1. Use a different counter for each section:**
```swift
private var cellStyleXfsIndex = 0
private var cellXfsIndex = 0
private var currentSection = ""

// In didStartElement:
if elementName == "cellStyleXfs" {
    currentSection = "cellStyleXfs"
} else if elementName == "cellXfs" {
    currentSection = "cellXfs"
} else if elementName == "xf" {
    if currentSection == "cellXfs" {
        // Process for cellStyles array
    }
    // Ignore if currentSection == "cellStyleXfs"
}
```

**2. Parse sections separately (two-pass):**
```swift
// First pass: extract cellXfs section as substring
let cellXfsXML = extractBetween("<cellXfs", "</cellXfs>", from: xml)

// Second pass: parse only cellXfs
parseCellXfs(cellXfsXML)
```

**3. Use XPath (if available):**
```swift
// XPath to select only cellXfs > xf elements
let cellStyles = xml.xpath("//cellXfs/xf")
```

## Technical Background

### Why Does Excel Have Both?

According to ECMA-376 (Office Open XML spec):

- **`cellStyleXfs`**: Master formatting records (templates)
  - Used by named cell styles (like "Normal", "Heading 1")
  - Cells don't reference these directly

- **`cellXfs`**: Extended formatting records (actual cell formats)
  - Each cell's `s` attribute references an index into this array
  - These inherit from cellStyleXfs via `xfId` attribute

### The Inheritance Chain

```xml
<cellStyleXfs count="1">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>  <!-- Master "Normal" -->
</cellStyleXfs>

<cellXfs count="2">
  <!-- This cell style inherits from cellStyleXfs[0] -->
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>

  <!-- This cell style overrides font, inherits rest from cellStyleXfs[0] -->
  <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
```

Most parsers can **ignore cellStyleXfs** entirely for basic read/write operations.

## Debugging Tips

### How to Detect This Bug

1. **Write a test that applies a style and reloads:**
   ```swift
   let cell = sheet.cell("A1")
   cell.setBackgroundColor("FFFF0000")  // Red
   workbook.save(to: tempURL)

   let reloaded = ExcelWorkbook(fileURL: tempURL)
   let reloadedCell = reloaded.sheet(at: 0).cell("A1")

   assert(reloadedCell.fill?.foregroundColor == "FFFF0000")  // Fails if bug exists
   ```

2. **Check the offset:**
   ```swift
   // Add debug output
   print("cellStyleXfs count:", cellStyleXfsCount)
   print("Expected style index:", expectedIndex)
   print("Actual style index:", actualIndex)
   print("Offset:", actualIndex - expectedIndex)
   ```

   If offset equals cellStyleXfs count, you have this bug!

3. **Examine the XML directly:**
   ```bash
   # Extract the xlsx
   unzip -q output.xlsx -d /tmp/excel

   # Count cellStyleXfs entries
   grep -c "<xf" /tmp/excel/xl/styles.xml | head -1

   # Find the cell's style reference
   grep 'r="A1"' /tmp/excel/xl/worksheets/sheet1.xml
   # Output: <c r="A1" s="1254">

   # Find what style 1254 references
   awk '/<cellXfs/,/<\/cellXfs>/' /tmp/excel/xl/styles.xml | \
     grep -n "<xf" | sed -n '1255p'
   ```

### Common Variations of This Bug

1. **Including cellStyleXfs in the count but not the data:**
   ```swift
   // BUG: Count includes cellStyleXfs but array doesn't
   let totalStyles = cellStyleXfsCount + cellXfsCount  // Wrong!
   ```

2. **Resetting index at wrong time:**
   ```swift
   // BUG: Index is never reset between sections
   case "xf":
       styles.append(parseXf())
       index += 1  // Keeps incrementing through both sections
   ```

3. **Using line numbers instead of element count:**
   ```swift
   // BUG: Line 150 in XML doesn't mean element 150
   if parser.lineNumber == 150 {  // Wrong!
   ```

## Testing Strategy

### Comprehensive Test

```swift
func testStyleIndicesAfterReload() throws {
    let workbook = try ExcelWorkbook(fileURL: templateURL)
    let sheet = try workbook.sheet(at: 0)

    // Use an empty cell to avoid template styles
    let cell = try sheet.cell("Z99")

    // Apply unique style
    cell.setBackgroundColor("FFFF0000")  // Red

    let outputURL = tempDirectory.appendingPathComponent("test.xlsx")
    try workbook.save(to: outputURL)

    // Reload
    let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
    let reloadedCell = try reloadedWorkbook.sheet(at: 0).cell("Z99")

    // Verify exact color match
    XCTAssertEqual(reloadedCell.fill?.foregroundColor, "FFFF0000",
                   "Style index must match after reload")

    // Also verify pattern type
    XCTAssertEqual(reloadedCell.fill?.patternType, .solid)
}
```

### Edge Cases to Test

1. **Empty cellStyleXfs:**
   ```xml
   <cellStyleXfs count="0"/>
   <cellXfs count="10">...</cellXfs>
   ```

2. **Multiple cellStyleXfs:**
   ```xml
   <cellStyleXfs count="5">...</cellStyleXfs>
   <cellXfs count="100">...</cellXfs>
   ```

3. **cellStyleXfs after cellXfs (malformed but might exist):**
   ```xml
   <cellXfs count="10">...</cellXfs>
   <cellStyleXfs count="1">...</cellStyleXfs>  <!-- Parser should handle -->
   ```

## Related Issues

This same pattern occurs in other Excel OOXML sections:

- **Fonts**: No dual section, safe
- **Fills**: No dual section, safe
- **Borders**: No dual section, safe
- **Number Formats**: Has built-in vs custom, but different element names
- **Cell Styles** (`<cellStyles>`): Different element name, safe

## References

- **ECMA-376 Part 1**: Office Open XML File Formats
  - Section 18.8.10: `cellStyleXfs` (Cell Style Formats)
  - Section 18.8.9: `cellXfs` (Cell Formats)

- **Microsoft [MS-XLSX]**: Excel (.xlsx) Extensions to Office Open XML
  - Section 2.4.54: Styles Part

- **Apache POI**: Reference implementation (Java)
  - `StylesTable.java`: Shows correct parsing
  - https://github.com/apache/poi/blob/trunk/poi-ooxml/src/main/java/org/apache/poi/xssf/model/StylesTable.java

## Impact

### Who This Affects

- ✅ Any custom Excel OOXML parser implementation
- ✅ Libraries parsing styles.xml (Swift, Rust, Go, Python, etc.)
- ✅ Tools that modify Excel files programmatically
- ❌ Libraries that only read (styles might be wrong but no corruption)
- ❌ Excel itself (handles this correctly)

### Severity Assessment

**Critical** because:
1. Silent data corruption (no error thrown)
2. Affects ALL styled cells after first save
3. Difficult to debug (colors appear random)
4. Users blame the entire library, not just styles
5. May not be caught by basic integration tests

## Fix Verification

After applying the fix, verify with:

```bash
# Run style creation tests
swift test --filter StyleCreationTests

# All 10 tests should pass:
# ✓ testApplyFont
# ✓ testMakeBold
# ✓ testMakeItalic
# ✓ testSetBackgroundColor
# ✓ testApplyFill
# ✓ testSetBorder
# ✓ testApplyBorder
# ✓ testApplyCompleteStyle
# ✓ testMultipleCellsStyling
# ✓ testNumberFormat
```

---

**Last Updated**: November 22, 2024
**Fixed In**: ApachePOISwift commit 368e665
**Reported By**: Claude Code during Phase 4B implementation

If you found this helpful, please star the repo: https://github.com/YourUsername/ApachePOISwift
