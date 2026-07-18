# Output Excel of Form — Reverse Engineering Investigation

**Date:** 2026-07-16  
**Status:** Investigation Complete  
**Scope:** 100% investigation — no code changes  
**Confidence Level:** See per-section ratings below

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Architecture Overview](#2-architecture-overview)
3. [Does It Modify the Original Workbook?](#3-does-it-modify-the-original-workbook)
4. [What Is Preserved in the Output Workbook](#4-what-is-preserved-in-the-output-workbook)
5. [Where Are Field Configurations Stored?](#5-where-are-field-configurations-stored)
6. [Data Flow: Designer → Export → Workbook](#6-data-flow-designer--export--workbook)
7. [Workbook Internals (XLSX Structure)](#7-workbook-internals-xlsx-structure)
8. [Hidden Sheets and Named Ranges](#8-hidden-sheets-and-named-ranges)
9. [Storage Strategy Summary](#9-storage-strategy-summary)
10. [Current Implementation Gaps](#10-current-implementation-gaps)
11. [Appendix: Key Files and Their Roles](#11-appendix-key-files-and-their-roles)

---

## 1. Executive Summary

The "Output Excel of Form" feature in the old PaperLess/ConMas system **creates a completely new workbook** from scratch. It does NOT modify the original uploaded workbook.

The generated workbook is produced by the `WorkbookGenerator.cs` class, which uses **Excel COM Interop** to:

1. Create a brand new workbook (`Workbooks.Add()`)
2. Configure sheets, page settings, print area, row heights, column widths, merged cells, freeze panes, and cell styles
3. Save as `.xlsx`

**Key finding:** The current `WorkbookGenerator` preserves the structural layout (dimensions, styling, merged cells) but does **NOT** embed field configuration metadata (types, readOnly, required, placeholder, etc.) into the workbook. Field metadata is only stored in the XML output and database objects.

However, in the old ConMas system, **Excel cell comments** were the primary carrier of field configuration data. The comment format was:
- Line 0: Field name
- Line 1: Field type (e.g., `KeyboardText`, `InputNumeric`)
- Line 2: ClusterIndex (sort order)
- Lines 3+: Input parameters (e.g., `Required=0;Lines=1;InputRestriction=None;MaxLength=0;...`)

| Question | Answer | Confidence |
|----------|--------|:----------:|
| Modifies original workbook? | **No** — creates entirely new workbook | **100%** |
| Preserves formatting? | **Partially** — page setup, row/col dimensions, merged cells, styles preserved. **Not preserved:** images, shapes, hidden rows/cols, header/footer, cell comments | **90%** |
| Field config stored where? | **Cell comments** in original (old system), **database** + **XML** + **runtime.json** in current system. Generated workbook has NO field config embedded | **100%** |
| Hidden sheets used? | **Unknown** — no evidence found in current codebase. Old system stored metadata in database, not hidden sheets | **70%** |
| Named ranges used? | **No** — no named range creation found in WorkbookGenerator or old system analysis | **90%** |
| Custom XML parts used? | **No** — standard OOXML structure with no custom XML parts found | **90%** |

---

## 2. Architecture Overview

### 2.1 Current System Architecture

```
Uploaded Excel (.xlsx)
       │
       ▼
  ┌──────────────────────────┐
  │  FormSaveService         │  ← Services/FormSaveService.cs
  │  SaveAsync()             │
  └──────┬──────────┬────────┘
         │          │
         ▼          ▼
  ┌──────────┐  ┌──────────────┐
  │ XML      │  │ Database     │
  │ Generator│  │ Generator    │
  │ .xml     │  │ .NET objects │
  └──────────┘  └──────┬───────┘
                       │
         ┌─────────────┼─────────────┐
         ▼             ▼             ▼
  ┌──────────┐  ┌──────────┐  ┌──────────┐
  │ Workbook │  │ Preview  │  │   PDF    │
  │ Generator│  │ Generator│  │ Generator│
  │  .xlsx   │  │   .png   │  │   .pdf   │
  └──────────┘  └──────────┘  └──────────┘
```

### 2.2 Old ConMas System Architecture

```
Excel Workbook (with cell comments)
       │
       ▼
  ┌──────────────────────────────┐
  │  ConMasClient.exe Publish    │
  └──────┬──────────┬────────────┘
         │          │
         ▼          ▼
  ┌──────────┐  ┌──────────────┐
  │  XML     │  │  PostgreSQL  │
  │  string  │  │  def_top     │
  │          │  │  def_sheet   │
  │          │  │  def_cluster │
  └──────────┘  └──────────────┘
                       │
         ┌─────────────┤
         ▼             ▼
  ┌──────────┐  ┌──────────┐
  │ Background│  │ Original │
  │ PDF bytes  │  │ Excel    │
  │ (in DB)   │  │ (saved)  │
  └──────────┘  └──────────┘
```

### 2.3 The "Output Excel of Form" in the Old System

The old system had a `finishOutput` / `excelOutput` configuration in the XML:

```xml
<finishOutput>1</finishOutput>
<excelOutput>1</excelOutput>
```

This controlled whether the system could generate an output Excel file after form submission. The output workbook was generated with user-entered values filled in, not as a blank template.

**Key files in current codebase:**

| File | Role |
|------|------|
| `Generators/WorkbookGenerator.cs` | Creates the output XLSX via Excel COM |
| `Generators/DatabaseGenerator.cs` | Converts FormDefinition to DB objects |
| `Generators/XmlGenerator.cs` | Generates XML from FormDefinition |
| `Services/FormSaveService.cs` | Orchestrates the full save pipeline |
| `Models/FormDefinition.cs` | The complete form data model |

---

## 3. Does It Modify the Original Workbook?

### Answer: **No — Creates Completely New Workbook**

**Evidence** (from `WorkbookGenerator.cs`):

```csharp
public string Generate(FormDefinition form, string outputPath)
{
    // ...
    excelApp = new Application { Visible = false, DisplayAlerts = false };
    workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);  // ← NEW workbook
    // ...
    workbook.SaveAs(outputPath);  // ← Saved to a NEW path
    // ...
}
```

The original uploaded workbook is never opened or modified. A brand new workbook is created from `xlWBATWorksheet` template, configured according to the `FormDefinition`, and saved to a separate output path.

**Confidence: 100%** — Source code directly confirms this.

### What About the Original Uploaded Workbook?

The original uploaded XLSX is preserved as-is in the `Forms/` directory (persistent copy) and the `Upload/` directory (temporary). It is never modified during any export operation.

---

## 4. What Is Preserved in the Output Workbook

### 4.1 Currently Preserved

| Feature | Preserved? | Implementation |
|---------|:----------:|----------------|
| Sheet names | ✅ | `ws.Name = sheetDef.Name` |
| Orientation (portrait/landscape) | ✅ | `ws.PageSetup.Orientation = XlPageOrientation.xlLandscape` |
| Margins (left, top, right, bottom) | ✅ | `ws.PageSetup.LeftMargin, .TopMargin, .RightMargin, .BottomMargin` |
| Center horizontally/vertically | ✅ | `ws.PageSetup.CenterHorizontally, .CenterVertically` |
| Zoom | ✅ | `ws.PageSetup.Zoom` |
| Fit to pages (wide/tall) | ✅ | `ws.PageSetup.FitToPagesWide, .FitToPagesTall` |
| Print area | ✅ | `ws.PageSetup.PrintArea` |
| Row heights | ✅ | `ws.Rows[row].RowHeight` |
| Column widths | ✅ | `ws.Columns[col].ColumnWidth` |
| Merged cells | ✅ | `range.Merge()` + apply cell style |
| Freeze panes | ✅ | `ws.Application.ActiveWindow.FreezePanes = true` |
| Cell styles | ✅ | Font name, size, bold, italic, underline, color, fill, alignment, wrap text |

### 4.2 NOT Currently Preserved

| Feature | Preserved? | Notes | Priority |
|---------|:----------:|-------|:--------:|
| Cell values | ❌ | No cell values are written | Low |
| Cell comments (field config) | ❌ | No comments are created | **High** — needed for tablet compatibility |
| Images / logos | ❌ | No drawing/image writing | Medium |
| Shapes | ❌ | No shape writing | Medium |
| Hidden rows/columns | ❌ | Not tracked in FormDefinition model | Low |
| PageSetup header/footer | ❌ | Not applied | Low |
| Background image | ❌ | Background rendered separately as PNG | Low |
| Border styles (cell borders) | ❌ | In CellStyleInfo model but not applied in WorkbookGenerator | Medium |
| Data validation | ❌ | Not applied | Low |
| Conditional formatting | ❌ | Not applied | Low |

**Confidence: 90%** — Based on source code analysis of every method in `WorkbookGenerator.cs`.

### 4.3 Preservation Comparison: Old System vs Current System

| Feature | Old ConMas | Current System |
|---------|:----------:|:--------------:|
| Page setup | ✅ Stored in DB + XML | ✅ Applied via COM |
| Column widths | ✅ Read from COM | ✅ Read from FormDefinition |
| Row heights | ✅ Read from COM | ✅ Read from FormDefinition |
| Merged cells | ✅ Preserved | ✅ Preserved |
| Cell styles | ✅ Preserved | ✅ Preserved |
| Cell comments | ✅ **Created by designer** | ❌ Not written |
| Field coordinates | ✅ Stored in DB | ✅ Stored in runtime.json + DB objects |
| Background image | ✅ PDF in DB | ✅ PNG at /forms/ or /preview/ |
| Original Excel file | ✅ Saved to Forms/ | ✅ Saved to Forms/ |
| Hidden sheets | ❌ Not used | ❌ Not supported |
| Named ranges | ❌ Not used | ❌ Not supported |
| Custom XML parts | ❌ Not used | ❌ Not supported |
| VBA macros | ❌ Stripped | ❌ Not supported |

---

## 5. Where Are Field Configurations Stored?

### 5.1 In the Old ConMas System

**Primary storage: Cell comments** (in the Excel workbook)

Comment format (from decompiled `ClusterInfo._texts`):
```
Line 0: ClusterName        (e.g., "samples", "Machine")
Line 1: ClusterTypeKey     (e.g., "KeyboardText", "InputNumeric")
Line 2: ClusterIndex       (numeric sort order, -1 if empty)
Line 3: InputParameters    (e.g., "Required=0;Lines=1;InputRestriction=None;MaxLength=0;...")
Lines 4+: Additional metadata
```

**Secondary storage: PostgreSQL database** (after publish)

| Table | Column | What It Stores |
|-------|--------|----------------|
| `def_top` | `def_top_name` | Form name |
| `def_top` | `xml_data` | Full XML definition |
| `def_top` | `background_image_file` | PDF binary |
| `def_sheet` | `def_sheet_name` | Sheet name |
| `def_sheet` | `width/height` | Page dimensions |
| `def_cluster` | `cluster_name` | Field name |
| `def_cluster` | `cluster_type` | Field type |
| `def_cluster` | `input_parameter` | Input parameters string |
| `def_cluster` | `left_position` etc. | Normalized coordinates |
| `def_cluster` | `cell_addr` | Cell/merge range address |
| `def_cluster` | `read_only` | Read-only flag |

**Tertiary storage: XML** (embedded in the definition)

The publish pipeline generates a comprehensive XML document containing ALL configuration data for every field, sheet, and the top-level form.

### 5.2 In the Current System

**Primary storage: `runtime.json` files** (in `Forms/` directory)

From `RuntimeCoordinateGenerator.SaveMetadata()`:

```json
{
  "version": "1.0",
  "capturedAt": "2026-07-13T09:15:44.6697944Z",
  "workbookName": "[V3.1_Sample]アンケート用紙.xlsx",
  "dpi": 300,
  "scaleX": 4.166667,
  "scaleY": 4.166667,
  "pageWidthPx": 2550,
  "pageHeightPx": 3299,
  "sheets": [
    {
      "name": "Sheet1",
      "index": 0,
      "pageWidthPx": 2550,
      "pageHeightPx": 3299,
      "fields": [
        {
          "id": "field_I6",
          "cellReference": "I6",
          "leftPx": 1775.6,
          "topPx": 796.6,
          "widthPx": 1442.5,
          "heightPx": 117.5,
          "leftRatio": 0.6963137,
          "topRatio": 0.2414671,
          "widthRatio": 0.5656863,
          "heightRatio": 0.0363747,
          "dataType": "text",
          "isMerged": true,
          "mergeRange": "I6:M6",
          "fontSize": 11,
          "bold": false,
          "readOnly": false,
          "required": false
        }
      ],
      "backgroundImage": "/forms/bg_91752bb875164097a7d0773f161e7d6a.png"
    }
  ]
}
```

**Secondary storage: Database generator objects** (`DatabaseGenerator.cs`)

The `DatabaseGenerator` produces `DefTop`, `DefSheet`, and `DefCluster` objects that mirror the old ConMas database schema:

| Object | Properties |
|--------|------------|
| `DefTop` | Id, Title, Author, Created, Modified, Version, Description, Metadata |
| `DefSheet` | Id, TopId, SheetId, Name, Index, PaperSize, Orientation, WidthPt, HeightPt, Margins, Center settings, Zoom, FitToPages, PrintArea, BackgroundImage, FreezePane, RowHeights, ColumnWidths, MergedCells |
| `DefCluster` | Id, TopId, ClusterId, Name, Type, SheetId, CellAddress, Coordinates (left/right/top/bottom in ratios and points), InputParameters, Visibility, Readonly, Remarks, Functions, Metadata |

**Tertiary storage: XML** (from `Generators/XmlGenerator.cs`)

Generates structured XML with all form data, including sheet settings, clusters with input parameters, and images.

### 5.3 What's Missing in the Generated Workbook

The `WorkbookGenerator` creates the structural workbook but does NOT embed field configuration metadata into the workbook. This means:

- ❌ No cell comments with field type/parameters (as the old system did)
- ❌ No cell values
- ❌ No data validation rules
- ❌ No field-specific formatting beyond what's in CellStyleInfo

**To fully replicate the old system's behavior, the generated workbook needs to:**
1. Write cell values back to the appropriate cells
2. Create cell comments with field configuration metadata (name, type, parameters)
3. Optionally apply data validation for field types (e.g., numeric only)

**Confidence: 100%** — Source code confirmed.

---

## 6. Data Flow: Designer → Export → Workbook

### 6.1 Current Flow

```
FormDefinition (created from uploaded Excel)
       │
       ▼
FormSaveService.SaveAsync()
       │
       ├── 1. XmlGenerator.Generate()        → form_{id}.xml
       ├── 2. DatabaseGenerator.Generate()    → DatabaseResult (DefTop/DefSheet/DefCluster)
       ├── 3. WorkbookGenerator.Generate()    → form_{id}.xlsx  ← THE OUTPUT WORKBOOK
       ├── 4. PreviewGenerator.Generate()     → preview_{id}.png
       └── 5. PdfGenerator.Generate()         → form_{id}.pdf
```

### 6.2 Old ConMas Flow

```
Designer (Excel Add-in)
  User selects cells, clicks "Add Field"
       │
       ▼
  Cell comment created with field metadata
       │
       ▼
  Workbook saved (.xlsx)
       │
       ▼
  ConMasClient.exe Publish
       │
       ├── Read Excel (COM Interop)
       │     ├── Comments → Cluster detection
       │     ├── Column widths, row heights
       │     └── Print area, page setup
       │
       ├── Compute coordinates (ratio normalization)
       │
       ├── Generate XML (comprehensive metadata)
       │
       ├── Export background PDF (ExportAsFixedFormat)
       │
       └── Insert into PostgreSQL database
             ├── def_top
             ├── def_sheet
             └── def_cluster
```

### 6.3 The "Output Excel of Form" Data Flow

When a user submits a completed form in the old system:

```
Report submission (user fills form on tablet)
       │
       ▼
  rep_* tables populated with values
       │
       ▼
  "Output Excel of Form" button
       │
       ▼
  Generate new Excel workbook:
    1. Take the original template workbook structure
    2. Fill in user-entered values at correct cells
    3. Apply field configuration (formatting, validation)
    4. Save as new workbook with user data
```

The current system's `WorkbookGenerator` appears to handle the template structure (step 1), but steps 2 and 3 (filling values and applying field config) are not yet implemented.

---

## 7. Workbook Internals (XLSX Structure)

### 7.1 What the Current WorkbookGenerator Produces

The generated XLSX contains standard OOXML parts:

| Part | Contents | Included? |
|------|----------|:---------:|
| `[Content_Types].xml` | Standard content types | ✅ Auto-generated |
| `xl/workbook.xml` | Workbook-level metadata | ✅ |
| `xl/_rels/workbook.xml.rels` | Relationships | ✅ |
| `xl/worksheets/sheet1.xml` | Worksheet content | ✅ (with merged cells, column widths, row heights) |
| `xl/styles.xml` | Cell styles | ✅ |
| `xl/sharedStrings.xml` | Shared string table | ❌ (no cell values written) |
| `xl/theme/theme1.xml` | Document theme | ✅ Auto-generated |
| `xl/comments1.xml` | Comments | ❌ Not written |
| `xl/drawings/drawing1.xml` | Drawings | ❌ Not written |
| `xl/vbaProject.bin` | VBA | ❌ Not supported |
| `docProps/app.xml` | Application properties | ✅ Auto-generated |
| `docProps/core.xml` | Core properties | ✅ Auto-generated |

### 7.2 What the Old System's Workbook Contains

The old system's template workbooks contain (verified from `Forms/*.xlsx` files):

- Standard OOXML structure
- **Cell comments** (the primary field config storage)
- **Merged cells** (defining field boundaries)
- **Print area** (defining page boundaries)
- Standard styles and formatting
- **No hidden sheets** (in analyzed templates)
- **No custom XML parts** (in analyzed templates)
- **No named ranges** (in analyzed templates)
- **No VBA** (in analyzed templates)

---

## 8. Hidden Sheets and Named Ranges

### 8.1 Hidden Sheets

**Finding: No hidden sheets used in the analyzed templates.**

Examined template workbooks (sample `Forms/*.xlsx` files) show only normal visible worksheets. The `WorkbookGenerator` does not create any hidden sheets.

**Confidence: 70%** — Based on a sample of analyzed workbooks; may vary across templates.

### 8.2 Named Ranges

**Finding: No named ranges used.**

The `WorkbookGenerator` does not create any named ranges. The old system's publish pipeline did not reference named ranges for configuration storage. The only range-based feature used is:
- **Print area** (`_xlnm.Print_Area` — a special built-in defined name)
- **Merge areas** (stored as `mergeCell` elements in the worksheet XML)

**Confidence: 90%** — Source code and DB analysis both support this.

### 8.3 Custom XML Parts

**Finding: Not used.**

The workbook does not contain any custom XML parts (`xl/customXml/`). Field configuration is stored externally (in comments, DB, or JSON), not within the workbook as custom XML.

**Confidence: 90%**

---

## 9. Storage Strategy Summary

### 9.1 Where Everything Is Stored

| Data | Storage Location | Format |
|------|-----------------|--------|
| Original Excel file | `Forms/{guid}.xlsx` | Binary XLSX |
| Background image (PNG) | `Forms/bg_{guid}.png` or `wwwroot/preview/page_{id}.png` | PNG |
| Field coordinates | `Forms/{id}.runtime.json` | JSON |
| Field coordinates (DB view) | `DefCluster` objects in `DatabaseResult` | .NET objects |
| Field configuration | `FormDefinition.Clusters[].InputParameters` | Dictionary (string→string) |
| Form XML definition | `Forms/form_{id}.xml` | XML |
| Output workbook | `Forms/form_{id}.xlsx` | XLSX (no field metadata embedded) |

### 9.2 What the Old System Stored vs Current System

| Item | Old System | Current System | Gap |
|------|:----------:|:--------------:|:----|
| Original XLSX | In Forms/ directory | In Forms/ directory | ✅ Same |
| Background PDF in DB | ✅ | ❌ (PNG on disk) | Different approach |
| Coordinates in runtime.json | ❌ | ✅ | New approach |
| Coordinates in def_cluster | ✅ | ✅ (via DatabaseGenerator) | ✅ Compatible |
| Field config in comments | ✅ | ❌ | **Gap** |
| Cell values in workbook | ✅ (after fill) | ❌ | **Gap** |
| Comprehensive XML | ✅ | ✅ | ✅ Compatible |
| Tablet-ready workbook | ✅ | ❌ | **Gap** |

---

## 10. Current Implementation Gaps

### 10.1 Gap 1: Field Configuration Not Embedded in Workbook

**Status:** ❌ Not implemented  
**Impact:** The generated workbook has correct structure but no field configuration data  
**Priority:** High  

The `WorkbookGenerator` needs to:
1. Write cell comments with field name (line 0), type (line 1), and input parameters (line 3+)
2. Write cell values (default values or user-entered values)
3. Apply data validation where appropriate

**Reference:** The old system's comment format was:
```
Line 0: ClusterName       → "samples", "Machine"
Line 1: ClusterTypeKey    → "KeyboardText", "InputNumeric"  
Line 2: ClusterIndex      → "0" (numeric, for sort order)
Line 3: InputParameters   → "Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;..."
```

### 10.2 Gap 2: User Values Not Written to Cells

**Status:** ❌ Not implemented  
**Impact:** Generated workbook is blank (no filled-in data)  
**Priority:** High (for output-with-values scenario)  

The current `WorkbookGenerator` accepts `FormDefinition` which doesn't contain user-entered values. A separate pipeline for outputting filled forms would need:
1. Sheet definition with cell value mapping
2. A value dictionary mapping field IDs → values → cell addresses

### 10.3 Gap 3: Images and Shapes Not Written

**Status:** ❌ Not implemented  
**Impact:** Generated workbook has no images or shapes  
**Priority:** Medium  

`ImageDefinition` exists in `FormDefinition` but is not consumed by `WorkbookGenerator`.

### 10.4 Gap 4: Cell Border Styles Not Fully Applied

**Status:** ⚠️ Partially implemented  
**Impact:** CellStyleInfo has border properties but `ApplyStyleToRange` doesn't apply them  
**Priority:** Low  

```csharp
// CellStyleInfo has:
public string? BorderTop { get; set; }
public string? BorderBottom { get; set; }
public string? BorderLeft { get; set; }
public string? BorderRight { get; set; }

// But ApplyStyleToRange doesn't use these properties
```

### 10.5 Gap 5: Background Image Not Embedded

**Status:** ⚠️ By design  
**Impact:** Background is a separate PNG, not embedded in the workbook  
**Priority:** Low  

The old system stored the background as PDF in the database. The current system stores it as PNG on disk. Neither embeds it in the generated workbook.

---

## 11. Appendix: Key Files and Their Roles

### Source Code (Current System)

| File | Role |
|------|------|
| `ExcelAPI/ExcelAPI/Generators/WorkbookGenerator.cs` | **Core** — Creates output XLSX via Excel COM |
| `ExcelAPI/ExcelAPI/Generators/DatabaseGenerator.cs` | Creates DB-compatible objects (DefTop, DefSheet, DefCluster) |
| `ExcelAPI/ExcelAPI/Generators/XmlGenerator.cs` | Generates comprehensive XML from FormDefinition |
| `ExcelAPI/ExcelAPI/Generators/PreviewGenerator.cs` | Generates PNG preview of the form |
| `ExcelAPI/ExcelAPI/Generators/PdfGenerator.cs` | Generates PDF from the form |
| `ExcelAPI/ExcelAPI/Services/FormSaveService.cs` | **Orchestrator** — calls all generators |
| `ExcelAPI/ExcelAPI/Models/FormDefinition.cs` | **Data model** — complete form definition structure |
| `ExcelAPI/ExcelAPI/Runtime/FormRuntimeBuilder.cs` | Builds runtime metadata from parsed workbook |
| `ExcelAPI/ExcelAPI/Runtime/RuntimeSerializer.cs` | Serializes runtime metadata to JSON |
| `ExcelAPI/ExcelAPI/Runtime/FieldDetector.cs` | Detects editable fields for runtime |
| `ExcelAPI/ExcelAPI/Runtime/RuntimeCoordinateGenerator.cs` | Saves/loads runtime metadata (runtime.json) |
| `ExcelAPI/ExcelAPI/Services/CoordinateTransformer.cs` | Coordinate transformation service |

### Documentation

| File | Description |
|------|-------------|
| `docs/PublishPipeline.md` | Complete legacy publish pipeline reconstruction |
| `docs/HowConMasLegacyWorks.md` | How the coordinate system works (decompiled DLL analysis) |
| `docs/LegacyGapAnalysis.md` | Gap analysis between current and legacy systems |
| `docs/TabletPipeline.md` | How the tablet renders forms from the database |
| `docs/Phase22_DesignerBlueprint.md` | Designer architecture blueprint |

### Runtime Metadata

| File | Description |
|------|-------------|
| `ExcelAPI/ExcelAPI/Forms/*.runtime.json` | Per-template runtime metadata (coordinates, fields, background URL) |
| `ExcelAPI/ExcelAPI/Forms/*.xlsx` | Per-template uploaded Excel workbooks |

---

## Confidence Level Legend

| Rating | Meaning |
|:------:|---------|
| **100%** | Confirmed by source code, decompiled code, or direct data analysis |
| **90%** | Strong evidence from code + data, reasonable inference |
| **80%** | Multiple sources of evidence support this conclusion |
| **70%** | Likely correct based on available evidence |
| **50%** | Best guess, requires further investigation |
| **<50%** | Unknown — needs original source code or binary analysis |

---

*End of report. This is an investigation-only deliverable — no code was modified.*
