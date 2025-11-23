# Excel Conditional Formatting - Quick Reference

**Location**: Worksheet XML (`xl/worksheets/sheet1.xml`)  
**Element**: `<conditionalFormatting>` after `<sheetData>`  
**Standards**: ECMA-376, ISO/IEC 29500

## Marbar Template Statistics

- **Total Blocks**: 706
- **Total Rules**: 716
- **Expression Rules**: 696 (97.2%)
- **DataBar Rules**: 20 (2.8%)
- **Unique dxfId Range**: 25-1107
- **Affected Sheets**: 20 of 26

---

## XML Structure Overview

### Basic Template

```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="expression" dxfId="5" priority="1" stopIfTrue="0">
    <formula>CONDITION_HERE</formula>
  </cfRule>
</conditionalFormatting>
```

### Attributes

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `sqref` | string | Yes | Cell range(s): "A1:A10" or "A1:A10 C5:C15" |
| `pivot` | 0/1 | No | Pivot table flag (default: 0) |
| `type` | string | Yes | expression, cellIs, dataBar, colorScale, iconSet, top10, aboveAverage, uniqueValues, duplicateValues, timePeriod |
| `dxfId` | int | No | Index into dxf array in styles.xml |
| `priority` | int | Yes | Evaluation order (1 = highest) |
| `stopIfTrue` | 0/1 | No | Stop evaluating further rules if match |
| `operator` | string | No | cellIs only: equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, between, notBetween, beginsWith, endsWith, containsText, doesNotContain |

---

## Rule Types Cheat Sheet

### 1. expression (Formula-Based) - 97% of Marbar

```xml
<cfRule type="expression" dxfId="1107" priority="18">
  <formula>F$2>0</formula>
</cfRule>
```
- Any formula that evaluates to TRUE/FALSE
- Cell references use relative notation
- Can use absolute anchors: $A$1, $A1, A$1, A1

### 2. cellIs (Value Comparison)

```xml
<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
  <formula>100</formula>
</cfRule>
```
- Operators: equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, between, notBetween, beginsWith, endsWith, containsText, doesNotContain
- Multiple `<formula>` tags for between/notBetween

### 3. dataBar (Progress Bars) - 3% of Marbar

```xml
<cfRule type="dataBar" priority="42">
  <dataBar showValue="0">
    <cfvo type="num" val="2"/>
    <cfvo type="num" val="4.2"/>
    <color theme="2" tint="-0.25"/>
  </dataBar>
</cfRule>
```
- cfvo type: num, percent, percentile, min, max, formula
- Color: theme + tint, or rgb="RRGGBB"

### 4. colorScale (Gradient)

```xml
<cfRule type="colorScale" priority="1">
  <colorScale>
    <cfvo type="min"/>
    <cfvo type="percentile" val="50"/>
    <cfvo type="max"/>
    <color rgb="F8696B"/>
    <color rgb="FFEB84"/>
    <color rgb="63BE7B"/>
  </colorScale>
</cfRule>
```
- 2-color or 3-color scales
- cfvo types: num, percent, percentile, min, max, formula

### 5. iconSet (Icons/Arrows)

```xml
<cfRule type="iconSet" priority="1">
  <iconSet iconSet="3Arrows" reverse="0">
    <cfvo type="percentile" val="0"/>
    <cfvo type="percentile" val="33"/>
    <cfvo type="percentile" val="67"/>
  </iconSet>
</cfRule>
```
- Icon sets: 3Arrows, 3ArrowsGray, 3Flags, 3Symbols, 4Arrows, 5Arrows, etc.

### 6. top10 (Top/Bottom Values)

```xml
<cfRule type="top10" dxfId="1" priority="3" rank="10" percent="0" bottom="0"/>
```
- rank: number of items
- percent: "1" for percentage
- bottom: "1" for bottom instead of top

### 7. aboveAverage

```xml
<cfRule type="aboveAverage" dxfId="1" priority="1"/>
```

### 8. uniqueValues

```xml
<cfRule type="uniqueValues" dxfId="1" priority="1"/>
```

### 9. duplicateValues

```xml
<cfRule type="duplicateValues" dxfId="1" priority="1"/>
```

### 10. timePeriod

```xml
<cfRule type="timePeriod" timePeriod="today" dxfId="1" priority="1"/>
```
- Values: today, yesterday, tomorrow, last7Days, lastMonth, nextMonth, nextWeek, thisMonth, thisWeek, thisYear

---

## Differential Formats (dxf in styles.xml)

Conditional formatting references styling via **dxfId** (differential format index).

```xml
<dxfs count="1000">
  <dxf>
    <font>
      <b/>
      <color rgb="FF0000"/>
    </font>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFF00"/>
      </patternFill>
    </fill>
  </dxf>
  <!-- ... 999 more dxf elements ... -->
</dxfs>
```

Can include:
- Font: bold, italic, color, size, underline, strikethrough
- Fill: background color, pattern
- Border: style, color
- Number format
- Alignment

---

## Key Implementation Notes

1. **Parsing**: Handle each `<conditionalFormatting>` block separately
2. **Priority**: Rules evaluated in priority order (1 = first)
3. **stopIfTrue**: When "1", stops evaluating further rules if this rule matches
4. **Multiple ranges**: sqref can contain space-separated ranges
5. **dxfId optional**: Some rules (dataBar, colorScale) may omit dxfId
6. **Formula flexibility**: expression type accepts any Excel formula

---

## Real Examples from Marbar

### Expression: Column has data
```xml
<cfRule type="expression" dxfId="1107" priority="18">
  <formula>F$2>0</formula>
</cfRule>
```

### Expression: Cell is empty
```xml
<cfRule type="expression" dxfId="945" priority="87">
  <formula>E$2>0</formula>
</cfRule>
```

### DataBar: Numeric range visualization
```xml
<cfRule type="dataBar" priority="42">
  <dataBar showValue="0">
    <cfvo type="num" val="0"/>
    <cfvo type="num" val="100"/>
    <color theme="2" tint="-0.249977111117893"/>
  </dataBar>
</cfRule>
```

---

## File Layout in .xlsx

```
workbook.xlsm
├── xl/
│   ├── worksheets/
│   │   ├── sheet1.xml        <-- Contains <conditionalFormatting> blocks
│   │   ├── sheet2.xml
│   │   └── ...
│   └── styles.xml            <-- Contains <dxfs> referenced by dxfId
```

Conditional formatting XML comes **after `<sheetData>`** in worksheet files.

---

## Testing Checklist

- [ ] Parse expression rules with formula
- [ ] Parse dataBar with cfvo elements
- [ ] Parse colorScale with 2 and 3 colors
- [ ] Parse iconSet with different types
- [ ] Handle multiple ranges in sqref
- [ ] Preserve rule priority order
- [ ] Verify dxfId references valid styles
- [ ] Write conditionalFormatting back to XML
- [ ] Round-trip: read → save → read (verify identical)
- [ ] Marbar template: 706 blocks, 716 rules preserved

---

**Full documentation**: See `ConditionalFormatting-Research.md`
