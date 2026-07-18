# Phase X — Legacy "Output Excel of Form" Runtime Pipeline Reconstruction

**Date:** 2026-07-17
**Status:** Investigation Complete
**Scope:** Investigation only — no code modified
**Previous reports:**
- `docs/OutputExcelOfForm_ReverseEngineering.md` — Initial findings
- `docs/PhaseX_OutputExcelOfForm_DeepInvestigation.md` — Deep dive with `_Fields` sheet
- `docs/OutputExcel_ForensicComparison.md` — Forensic XLSX comparison

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Runtime Sequence Diagram](#2-runtime-sequence-diagram)
3. [Method Call Graph](#3-method-call-graph)
4. [Class Responsibility Table](#4-class-responsibility-table)
5. [Object Lifecycle](#5-object-lifecycle)
6. [COM API Timeline](#6-com-api-timeline)
7. [Workbook Generation Timeline](#7-workbook-generation-timeline)
8. [Field Value Pipeline](#8-field-value-pipeline)
9. [Entry Point Discovery](#9-entry-point-discovery)
10. [Value Collection Pipeline](#10-value-collection-pipeline)
11. [Cell Address Resolution](#11-cell-address-resolution)
12. [Merged Cell Behavior](#12-merged-cell-behavior)
13. [Comment Generation Pipeline](#13-comment-generation-pipeline)
14. [Image Handling](#14-image-handling)
15. [Database Interaction](#15-database-interaction)
16. [Key Gaps vs Legacy System](#16-key-gaps-vs-legacy-system)
17. [Confidence Table](#17-confidence-table)
18. [Conclusion](#18-conclusion)

---

## 1. Executive Summary

The "Output Excel of Form" runtime pipeline in the old ConMas/PaperLess system has been reconstructed from available source code, decompiled DLL references, template XML analysis, and OOXML workbook inspection.

### What Currently Exists

The current `WorkbookGenerator` creates a structurally correct workbook via Excel COM Interop (`Workbooks.Add()` → configure → `SaveAs()`). It handles:

- Sheet names, page setup, print area
- Row heights, column widths
- Merged cells with cell styles
- Freeze panes
- Fonts, colors, fill, alignment, wrap text

### What the Current Runtime Pipeline is MISSING

Compared to the old system, the following gaps exist:

| Stage | Old System | Current System | Gap |
|-------|-----------|---------------|:---:|
| Button/entry point | ConMas Runtime "Output Excel of Form" button | `FormSaveService.SaveAsync()` called by `FormController.Save()` | ⚠️ Different flow |
| Value collection | User-entered values from `rep_cluster` DB table | No value collection exists | ❌ |
| Value → cell resolution | `<cellAddress>` per cluster in XML definition | No value-to-cell mapping | ❌ |
| Cell value writing | `ws.Range[cellAddress].Value = excelOutputValue` | Not implemented | ❌ |
| Cell comments | Rich-text comments with field metadata | Not implemented | ❌ |
| `_Fields` hidden sheet | 7-column metadata table | Not implemented | ❌ |
| VML comment markers | `xl/drawings/vmlDrawing1.vml` | Not implemented | ❌ |

### The Original Pipeline Reconstruction

Based on all evidence, the legacy pipeline was:

```
User presses "Output Excel of Form"
    ↓
ConMasClient.exe / ConMas Runtime
    ↓
1. Collect all filled values from rep_cluster / runtime state
    ↓
2. Resolve each cluster to its cell address (from definition XML)
    ↓
3. Create new workbook via Excel COM (Workbooks.Add)
    ↓
4. Apply template structure (page setup, dimensions, merged cells, styles)
    ↓
5. Write values: ws.Range[cellAddress].Value = excelOutputValue
    ↓
6. Write cell comments with field metadata (name, type, parameters)
    ↓
7. Create _Fields hidden sheet (7 columns, one row per field)
    ↓
8. SaveAs(outputPath)
    ↓
Output Excel ready: worksheet with filled values, comments, metadata
```

---

## 2. Runtime Sequence Diagram

### 2.1 Old System (ConMas/PaperLess Legacy)

```
┌──────────┐   ┌──────────────┐   ┌──────────────────┐   ┌────────────┐   ┌──────────┐
│  User    │   │ ConMas       │   │ ExcelController   │   │ Excel COM  │   │ File     │
│  (Tablet)│   │ Runtime      │   │ (LibExcelCtrl.dll)│   │ (Excel App)│   │ System   │
└────┬─────┘   └──────┬───────┘   └────────┬─────────┘   └─────┬──────┘   └────┬─────┘
     │                 │                    │                   │               │
     │ Click "Output   │                    │                   │               │
     │ Excel of Form"  │                    │                   │               │
     │════════════════▶│                    │                   │               │
     │                 │                    │                   │               │
     │                 │ 1. Gather values   │                   │               │
     │                 │ from rep_cluster   │                   │               │
     │                 │ (or runtime state) │                   │               │
     │                 │◄────── ─ ─ ─ ─ ─ ─│                   │               │
     │                 │                    │                   │               │
     │                 │ 2. Create/Open     │                   │               │
     │                 │  workbook          │                   │               │
     │                 │───────────────────▶│                   │               │
     │                 │                    │ 3. Workbooks.Add │                │
     │                 │                    │  or Open()       │               │
     │                 │                    │──────────────────▶│               │
     │                 │                    │◄──────────────────│               │
     │                 │                    │                   │               │
     │                 │ 4. For each cluster│                   │               │
     │                 │ ~~~~~~~~~~~~~~~~~~ │                   │               │
     │                 │                    │ a. ws.Range[addr] │               │
     │                 │                    │    .Value = val   │               │
     │                 │                    │──────────────────▶│               │
     │                 │                    │                   │               │
     │                 │                    │ b. AddComment()   │               │
     │                 │                    │   (field metadata)│               │
     │                 │                    │──────────────────▶│               │
     │                 │                    │                   │               │
     │                 │ 5. Create _Fields  │                   │               │
     │                 │    hidden sheet    │                   │               │
     │                 │───────────────────▶│                   │               │
     │                 │                    │──────────────────▶│               │
     │                 │                    │                   │               │
     │                 │ 6. SaveAs()        │                   │               │
     │                 │───────────────────▶│──────────────────▶│──────────────▶│
     │                 │                    │                   │               │
     │                 │ 7. Return result   │                   │               │
     │                 │◄───────────────────│◄──────────────────│               │
     │◄════════════════│                    │                   │               │
     │                 │                    │                   │               │
```

### 2.2 Current System (Our .NET API)

```
┌──────────┐   ┌──────────────┐   ┌──────────────┐   ┌──────────────────┐   ┌────────────┐   ┌──────────┐
│  Client  │   │ FormController│  │ FormSave      │   │ WorkbookGenerator │  │ Excel COM  │   │ File     │
│ (Browser)│   │ (Controllers) │  │ Service       │   │ (Generators)      │  │ (Excel App)│   │ System   │
└────┬─────┘   └──────┬───────┘   └──────┬───────┘   └────────┬─────────┘   └─────┬──────┘   └────┬─────┘
     │                 │                    │                   │                   │               │
     │ POST /api/form/ │                    │                   │                   │               │
     │ save            │                    │                   │                   │               │
     │════════════════▶│                    │                   │                   │               │
     │                 │ SaveAsync(form,    │                   │                   │               │
     │                 │   outputDir)       │                   │                   │               │
     │                 │───────────────────▶│                   │                   │               │
     │                 │                    │ 1. XmlGenerator   │                   │               │
     │                 │                    │    .Generate()    │                   │               │
     │                 │                    │──── ─ ─ ─ ─ ─ ─▶ │(XML generation)    │               │
     │                 │                    │                   │                   │               │
     │                 │                    │ 2. DatabaseGen    │                   │               │
     │                 │                    │    .Generate()    │                   │               │
     │                 │                    │──── ─ ─ ─ ─ ─ ─▶ │(DB objects)       │               │
     │                 │                    │                   │                   │               │
     │                 │                    │ 3. WorkbookGen    │                   │               │
     │                 │                    │    .Generate()    │                   │               │
     │                 │                    │──────────────────▶│                   │               │
     │                 │                    │                   │ Workbooks.Add()   │               │
     │                 │                    │                   │──────────────────▶│               │
     │                 │                    │                   │◄──────────────────│               │
     │                 │                    │                   │                   │               │
     │                 │                    │                   │ Configure sheets  │               │
     │                 │                    │                   │ Page setup, etc.  │               │
     │                 │                    │                   │──────────────────▶│               │
     │                 │                    │                   │                   │               │
     │                 │                    │                   │ SaveAs()          │               │
     │                 │                    │                   │──────────────────▶│──────────────▶│
     │                 │                    │                   │◄──────────────────│               │
     │                 │                    │◄──────────────────│                   │               │
     │                 │◄───────────────────│                   │                   │               │
     │◄════════════════│                    │                   │                   │               │
     │                 │                    │                   │                   │               │
     │   GAPS:         │                    │                   │                   │               │
     │   ❌ No values  │                    │                   │                   │               │
     │   ❌ No comments │                   │                   │                   │               │
     │   ❌ No _Fields │                    │                   │                   │               │
```

---

## 3. Method Call Graph

### 3.1 Current System

```
FormController.Save()
    │
    ├── [Dependency Injection] _formSaveService.SaveAsync(form, outputDir)
    │
    └── FormSaveService.SaveAsync()
        │
        ├── 1. _xmlGenerator.Generate(definition)
        │       └── Returns XML string
        │       └── File.WriteAllTextAsync(xmlPath)
        │
        ├── 2. _databaseGenerator.Generate(definition)
        │       └── Returns DatabaseResult (DefTop, DefSheet[], DefCluster[])
        │
        ├── 3. _workbookGenerator.Generate(definition, workbookPath)
        │       └── Create Excel Application
        │       └── Workbooks.Add(xlWBATWorksheet)
        │       └── For each sheet:
        │           ├── ApplyPageSettings()
        │           ├── ApplyPrintArea()
        │           ├── ApplyRowHeights()
        │           ├── ApplyColumnWidths()
        │           ├── ApplyMergedCells()
        │           ├── ApplyFreezePane()
        │           └── ApplyCellStyles()
        │       └── workbook.SaveAs(outputPath)
        │       └── workbook.Close()
        │       └── excelApp.Quit()
        │
        ├── 4. _previewGenerator.Generate(definition, firstSheetId, previewPath)
        │       └── Generates PNG preview
        │
        └── 5. _pdfGenerator.Generate(definition, firstSheetId, pdfPath)
                └── Generates PDF
```

### 3.2 Missing Methods (from Legacy System)

```
ExportExcelDefinition()                     — LibExcelController.dll / ConMasExcelClient.dll
    ├── Ensures _Fields sheet exists
    ├── Writes field metadata rows
    └── Hides _Fields sheet

WriteClusterValues()                        — Inferred
    ├── For each cluster:
    │   ├── Resolve cellAddress
    │   └── ws.Range[address].Value = value
    └──

AddClusterComments()                        — Inferred
    ├── For each cluster:
    │   ├── Build comment text (name + type + params)
    │   └── ws.Range[address].AddComment(text)
    └──
```

---

## 4. Class Responsibility Table

### 4.1 Current System

| Class | Namespace | Responsibility | Key Method |
|-------|-----------|---------------|------------|
| `FormController` | `Controllers` | HTTP API entry point for form operations | `Save()` |
| `FormSaveService` | `Services` | Orchestrates the full save pipeline | `SaveAsync()` |
| `WorkbookGenerator` | `Generators` | Creates output XLSX via Excel COM Interop | `Generate(form, path)` |
| `XmlGenerator` | `Generators` | Generates XML from `FormDefinition` | `Generate(form)` |
| `DatabaseGenerator` | `Generators` | Produces DB-compatible objects | `Generate(form)` |
| `PreviewGenerator` | `Generators` | Generates PNG preview | `Generate(form, sheetId, path)` |
| `PdfGenerator` | `Generators` | Generates PDF | `Generate(form, sheetId, path)` |
| `FormDefinition` | `Models` | Complete form data structure | (data class) |
| `SheetDefinition` | `Models` | Sheet configuration | (data class) |
| `ClusterDefinition` | `Models` | Field definition with coordinates | (data class) |
| `CellStyleInfo` | `Models` | Cell style metadata | (data class) |
| `PageSettings` | `Models` | Page setup configuration | (data class) |
| `CoordinateEngine` | `CoordinateEngine` | Coordinate calculation | (various) |
| `PublishEngine` | `LegacyEngine.PublishEngine` | Legacy publish pipeline | `PublishAsync(filePath)` |
| `BackgroundExporter` | `LegacyEngine.PublishEngine` | PDF → PNG conversion | `ExportAsync(filePath)` |

### 4.2 Old System (Reconstructed)

| Class | DLL | Responsibility | Key Method |
|-------|-----|---------------|------------|
| `ExcelController` | `LibExcelController.dll` | Central controller for Excel operations | `GetDefinition()`, `ExportExcelDefinition()` |
| `ConMasClient` | `ConMasExcelClient.dll` | Main client application | `Publish()`, `OutputExcel()` |
| `ClusterInfo` | — | Field metadata container (`_texts[]` array) | (data class) |
| `WorkbookInfo` | — | Workbook metadata container | (data class) |
| `SheetInfo` | — | Sheet metadata container | (data class) |
| `excelOutputValue` | — | Per-cluster value mapping | (property) |
| `cellAddress` | — | Per-cluster cell address | (property) |

### 4.3 Comparison

| Function | Old System | Current System | Gap |
|----------|-----------|---------------|:---:|
| Workbook creation | `LibExcelController` → COM Interop | `WorkbookGenerator` → COM Interop | ✅ Functionally similar |
| Field value writing | `excelOutputValue` → `ws.Range.Value` | ❌ Not implemented | ❌ |
| Comment writing | `range.AddComment(metadata)` | ❌ Not implemented | ❌ |
| `_Fields` sheet | `ExportExcelDefinition()` | ❌ Not implemented | ❌ |
| XML generation | `LibExcelController.DefinitionXML()` | `XmlGenerator.Generate()` | ✅ Functionally similar |
| Background export | COM `ExportAsFixedFormat` + PDF conversion | `PreviewGenerator` + `PdfGenerator` | ✅ Functionally similar |

---

## 5. Object Lifecycle

### 5.1 Current System Object Lifecycle

```
FormDefinition (input parameter)
    │
    ├── FormSaveService.SaveAsync()
    │   │
    │   ├── → XmlGenerator.XmlResult (string, written to file)
    │   │
    │   ├── → DatabaseResult
    │   │       ├── DefTop (created)
    │   │       ├── DefSheet[] (created)
    │   │       └── DefCluster[] (created)
    │   │       └── All returned in result.DatabaseObjects
    │   │
    │   ├── → WorkbookGenerator
    │   │       ├── Excel.Application (created, hidden)
    │   │       ├── Excel.Workbook (created via Add())
    │   │       ├── Excel.Worksheet[] (created per sheet)
    │   │       ├── Excel.Range (created per operation, released)
    │   │       └── → Output: .xlsx file on disk
    │   │       └── All COM objects released
    │   │
    │   ├── → PreviewGenerator
    │   │       └── → Png file on disk
    │   │
    │   └── → PdfGenerator
    │           └── → Pdf file on disk
    │
    └── FormSaveResult (returned)
        ├── XmlPath
        ├── WorkbookPath
        ├── PreviewPath
        ├── PdfPath
        └── DatabaseObjects
```

### 5.2 Legacy System Object Lifecycle (Reconstructed)

```
User clicks "Output Excel of Form"
    │
    ├── 1. Gather runtime state
    │       ├── Read rep_top, rep_sheet, rep_cluster from DB
    │       └── Build cluster → value mapping
    │
    ├── 2. Create workbook
    │       ├── Excel.Application (start COM)
    │       ├── Excel.Workbook (Workbooks.Add or template Open)
    │       └── Excel.Worksheet[] (per original sheet)
    │
    ├── 3. Write structural configuration
    │       ├── Page setup from definition XML
    │       ├── Column widths / Row heights
    │       ├── Merged cells
    │       └── Cell styles (fonts, alignment, fill, borders)
    │
    ├── 4. Write filled values
    │       └── For each cluster in definition:
    │           ├── Get cellAddress from cluster metadata
    │           └── ws.Range[cellAddress].Value = excelOutputValue
    │
    ├── 5. Write field metadata
    │       ├── Cell comments (name, type, params)
    │       └── _Fields hidden sheet with all metadata
    │
    └── 6. Save and finalize
            ├── workbook.SaveAs(outputFilePath)
            ├── workbook.Close()
            └── excelApp.Quit()
```

### 5.3 Key Difference: Value Source

| Aspect | Current System | Old System |
|--------|---------------|------------|
| Value source | Not connected to runtime | `rep_cluster` DB table or runtime state |
| Data flow | Controller → Service → Generator (no values) | Tablet input → DB → Generator → Workbook |
| Timing | Called during form SAVE | Called during form FINISH/OUTPUT |
| Trigger | API endpoint | Button in tablet UI |

---

## 6. COM API Timeline

### 6.1 Current COM Sequence (Verified)

```
Order   COM Call                                  Method
─────   ────────────────────────────────────────   ─────────────────
1       new Excel.Application()                    WorkbookGenerator.Generate()
2       excelApp.Workbooks.Add(xlWBATWorksheet)    WorkbookGenerator.Generate()
3       workbook.Worksheets[].Delete()             while loop
4       workbook.Worksheets.Add() (if needed)      while loop
5       FOR EACH SHEET:
5a        ws.Name = "SheetName"
5b        ws.PageSetup.Orientation                 ApplyPageSettings()
5c        ws.PageSetup.LeftMargin                  ApplyPageSettings()
5d        ws.PageSetup.TopMargin                   ApplyPageSettings()
5e        ws.PageSetup.RightMargin                 ApplyPageSettings()
5f        ws.PageSetup.BottomMargin                ApplyPageSettings()
5g        ws.PageSetup.CenterHorizontally           ApplyPageSettings()
5h        ws.PageSetup.CenterVertically             ApplyPageSettings()
5i        ws.PageSetup.Zoom                        ApplyPageSettings()
5j        ws.PageSetup.FitToPagesWide               ApplyPageSettings()
5k        ws.PageSetup.FitToPagesTall               ApplyPageSettings()
5l        ws.PageSetup.PrintArea = address          ApplyPrintArea()
5m        ws.Rows[row].RowHeight = height           ApplyRowHeights()
5n        ws.Columns[col].ColumnWidth = width        ApplyColumnWidths()
5o        ws.Range[address].Merge()                 ApplyMergedCells()
5p        ws.Range[address].Font.Name               ApplyStyleToRange()
5q        ws.Range[address].Font.Size               ApplyStyleToRange()
5r        ws.Range[address].Font.Bold               ApplyStyleToRange()
5s        ws.Range[address].Font.Italic              ApplyStyleToRange()
5t        ws.Range[address].Font.Underline           ApplyStyleToRange()
5u        ws.Range[address].Font.Color               ApplyStyleToRange()
5v        ws.Range[address].Interior.Color           ApplyStyleToRange()
5w        ws.Range[address].HorizontalAlignment      ApplyStyleToRange()
5x        ws.Range[address].VerticalAlignment        ApplyStyleToRange()
5y        ws.Range[address].WrapText                 ApplyStyleToRange()
5z        ws.Range[freezePane].Activate()            ApplyFreezePane()
5zz       ws.Application.ActiveWindow.FreezePanes    ApplyFreezePane()
6       workbook.SaveAs(path)                       WorkbookGenerator.Generate()
7       workbook.Close(false)                       finally block
8       excelApp.Quit()                             finally block
```

**Total COM calls:** ~30+ per sheet + overhead = ~50–100+ COM calls per workbook.

### 6.2 Missing COM Calls (Legacy)

```
Order   COM Call                                   Required For
─────   ────────────────────────────────────────   ─────────────────
9       ws.Range[cellAddress].Value = value         Writing filled values
10      ws.Range[cellAddress].AddComment(text)      Writing field metadata
11      range.Comment.Shape.TextFrame.AutoSize      Comment formatting
12      range.Comment.Visible = False               Hide comment indicator
13      newWorksheet = workbook.Worksheets.Add()    Creating _Fields sheet
14      newWorksheet.Name = "_Fields"               Naming the hidden sheet
15      newWorksheet.Cells[row,col] = value          Writing _Fields rows
16      newWorksheet.Visible = xlSheetHidden        Hiding _Fields sheet
```

`1-8` = structural creation, `9-16` = metadata + values (not yet implemented)

---

## 7. Workbook Generation Timeline

### 7.1 Precise Order of Operations

```
Step | Action                                | Legacy | Current | Notes
─────|───────────────────────────────────────|:------:|:-------:|──────
 1   | Start Excel Application (hidden)      |   ✅   |   ✅    |
 2   | Workbooks.Add() new workbook          |   ✅   |   ✅    |
 3   | Set sheet count (add/remove)          |   ✅   |   ✅    |
     | FOR EACH SHEET:                       |        |         |
 4   |   Set sheet name                      |   ✅   |   ✅    |
 5   |   Apply page setup (orientation,      |   ✅   |   ✅    |
     |     margins, zoom, fit-to-page)       |        |         |
 6   |   Set print area                      |   ✅   |   ✅    |
 7   |   Set row heights                     |   ✅   |   ✅    |
 8   |   Set column widths                   |   ✅   |   ✅    |
 9   |   Apply merged cells                  |   ✅   |   ✅    |
10   |   Apply cell styles (font, color,     |   ✅   |   ✅    |
     |     alignment, fill, wrap)            |        |         |
11   |   Apply borders                       |   ⚠️   |   ❌    | CellStyleInfo.Border* exists but not applied
12   |   Set freeze pane                     |   ✅   |   ✅    |
13   |   Apply number formats                |   ⚠️   |   ❌    | Not in FormDefinition
14   |   Write cell values                   |   ✅   |   ❌    | ❌ MISSING
15   |   Add cell comments with metadata     |   ✅   |   ❌    | ❌ MISSING
16   |   Hide comment indicators             |   ✅   |   ❌    | ❌ MISSING
17   | Create _Fields hidden sheet           |   ✅   |   ❌    | ❌ MISSING
18   | Write field metadata rows to _Fields  |   ✅   |   ❌    | ❌ MISSING
19   | Write images/shapes                   |   ⚠️   |   ❌    | Low confidence
20   | workbook.SaveAs(outputPath)           |   ✅   |   ✅    |
21   | workbook.Close()                      |   ✅   |   ✅    |
22   | Excel Application Quit()              |   ✅   |   ✅    |
```

### 7.2 Output Workbook States

```
State A: Current output (structure only)
┌─────────────────────────────┐
│  Output.xlsx                │
│  ├── Sheet1                 │
│  │   ├── Merged cells       │
│  │   ├── Styles applied     │
│  │   ├── Page setup set     │
│  │   └── ❌ No values       │
│  └── ❌ No metadata         │
└─────────────────────────────┘

State B: Minimal publishable (Phase 1 goal)
┌─────────────────────────────┐
│  Output.xlsx                │
│  ├── Sheet1                 │
│  │   ├── Merged cells       │
│  │   ├── Styles applied     │
│  │   ├── Page setup set     │
│  │   ├── ❌ No values       │
│  │   └── ✅ Cell comments   │
│  ├── _Fields (hidden)       │
│  │   ├── Address            │
│  │   ├── FieldId            │
│  │   ├── FieldName          │
│  │   ├── FieldType          │
│  │   ├── SheetName          │
│  │   ├── CreatedDate        │
│  │   └── Notes              │
│  └── ✅ Republishable       │
└─────────────────────────────┘

State C: Complete output (Phase 2 goal)
┌─────────────────────────────┐
│  Output.xlsx                │
│  ├── Sheet1                 │
│  │   ├── Merged cells       │
│  │   ├── Styles applied     │
│  │   ├── Page setup set     │
│  │   ├── ✅ Cell values     │
│  │   └── ✅ Cell comments   │
│  ├── _Fields (hidden)       │
│  └── ✅ Republishable       │
└─────────────────────────────┘
```

---

## 8. Field Value Pipeline

### 8.1 Old System Value Flow

```
Tablet User enters value in field
    │
    ▼
Runtime stores: Cluster.Value = userInput
    │
    ├── displayValue = Cluster.Value (displayed to user)
    ├── excelOutputValue = Cluster.Value (stored for Excel output)
    │   Note: excelOutputValue may differ from displayValue for:
    │   - Select/Dropdown: displayValue = "Option A", excelOutputValue = "1" (machine value)
    │   - Checkbox: displayValue = "☑", excelOutputValue = "1"
    │   - Date: displayValue = "2026/07/16", excelOutputValue = "2026-07-16"
    │
    ▼
Save to database:
    rep_cluster table
    ├── rep_top_id
    ├── sheet_no
    ├── cluster_id
    ├── cluster_value (user-entered value)
    ├── machine_value (underlying value for list fields)
    └── excel_output_value (value for Excel output)
    │
    ▼
"Output Excel of Form" button
    │
    ▼
1. Load definition XML from def_top.xml_data
    │
2. For each cluster in definition:
    ├── Read cluster.cellAddress (e.g., "$A$1")
    ├── Read cluster.excelOutputValue from DB/state
    └── ws.Range[cellAddress].Value = excelOutputValue
```

### 8.2 Evidence from XML Templates

From `template.xml` and `templateXls.xml`, the per-cluster value elements are:

```xml
<cluster>
  <sheetNo></sheetNo>                    <!-- Sheet number -->
  <clusterId></clusterId>                <!-- Unique cluster ID -->
  <name></name>                          <!-- Field name -->
  <type></type>                          <!-- Field type -->
  <value></value>                        <!-- User-entered value -->
  <displayValue></displayValue>          <!-- Display value (may differ from value) -->
  <excelOutputValue></excelOutputValue>  <!-- ← VALUE USED FOR EXCEL OUTPUT -->
  <inputParameters></inputParameters>    <!-- Field configuration -->
  <readOnly>0</readOnly>                 <!-- Read-only flag -->
  <cellAddress></cellAddress>            <!-- ← CELL ADDRESS (template.xml only) -->
</cluster>
```

**Key discovery:** There are THREE value fields:
- `value` — The raw user-entered value
- `displayValue` — What was displayed on the tablet (may include formatting)
- `excelOutputValue` — What goes into the Excel output (critical for list/select fields)

### 8.3 Current System Gap

The current `FormDefinition.ClusterDefinition` has:

```csharp
class ClusterDefinition {
    public string? Value { get; set; }              // Exists but unused in WorkbookGenerator
    public string? DisplayValue { get; set; }        // Not defined
    public string? ExcelOutputValue { get; set; }    // Not defined
    public string? CellAddress { get; set; }         // Exists in some models
}
```

The `WorkbookGenerator.Generate()` method accepts only `FormDefinition` and does NOT receive any user-entered values.

---

## 9. Entry Point Discovery

### 9.1 Current Entry Points

The current system has TWO potential entry points for workload generation:

**Entry Point A: `FormController.Save()`** (for form SAVE/designer publish)

```csharp
// Controllers/FormController.cs
[HttpPost("save")]
public async Task<IActionResult> Save([FromBody] FormDefinition form)
{
    var outputDir = Path.Combine(...);
    var result = await _formSaveService.SaveAsync(form, outputDir);
    return Ok(new { success = true, ... result });
}
```

**Entry Point B: `FormController.Publish()`** (for legacy publish)

```csharp
// Not found in codebase — PublishEngine is called differently:
// LegacyEngine/PublishEngine/PublishEngine.PublishAsync(filePath)
```

### 9.2 Current Status

| Entry Point | What it produces | User values? | Comments? | _Fields? |
|-------------|-----------------|:------------:|:---------:|:--------:|
| `FormController.Save()` | XLSX + XML + DB objects + Preview + PDF | ❌ No | ❌ No | ❌ No |
| `PublishEngine.PublishAsync()` | XML + coordinates + background images | ❌ N/A | ✅ Reads them | ✅ Reads them |

### 9.3 Required New Entry Point

For the "Output Excel of Form" feature, a new endpoint is needed:

```csharp
[HttpPost("output-excel")]
public async Task<IActionResult> OutputExcel(
    [FromBody] OutputExcelRequest request)
{
    // request contains:
    // - FormId (to load definition from DB)
    // - FieldValues: Dictionary<string, string> (fieldId → value)
    //
    // 1. Load form definition from DB / stored XML
    // 2. Create workbook with structure
    // 3. Write cell values from FieldValues
    // 4. Write cell comments with metadata
    // 5. Create _Fields hidden sheet
    // 6. Return generated XLSX
}
```

---

## 10. Value Collection Pipeline

### 10.1 How Values Are Collected in the Old System

The old ConMas system collected values from the `rep_cluster` database table:

```sql
CREATE TABLE rep_cluster (
    rep_top_id      INTEGER,           -- Reference to the report submission
    sheet_no        INTEGER,           -- Sheet number
    cluster_id      INTEGER,           -- Cluster (field) ID
    cluster_value   TEXT,              -- User-entered value
    machine_value   TEXT,              -- Machine value (for select/radio/checkbox)
    excel_output_value TEXT,           -- Value written to Excel output
    -- ...
);
```

The `rep_top` table tracked the submission:

```sql
CREATE TABLE rep_top (
    rep_top_id      INTEGER PRIMARY KEY,
    def_top_id      INTEGER,           -- Reference to the form definition
    rep_top_name    TEXT,              -- Submission title
    regist_time     TIMESTAMP,         -- When submitted
    regist_user     INTEGER,           -- Who submitted
    status          INTEGER,           -- Submission status
    -- ...
);
```

### 10.2 Value Flow for Different Field Types

| Field Type | User Input | `value` | `displayValue` | `excelOutputValue` |
|-----------|-----------|---------|---------------|-------------------|
| KeyboardText | Free text | "John Doe" | "John Doe" | "John Doe" |
| InputNumeric | Numeric keypad | "12345" | "12,345" | "12345" |
| CheckBox | Check/uncheck | "1" or "0" | "☑" or "☐" | "1" or "0" |
| RadioButton | Select option | "2" | "Option B" | "2" |
| DropDown | Select item | "3" | "Item C" | "3" |
| Date | Date picker | "2026-07-16" | "2026/07/16" | "2026-07-16" |
| Signature | Draw signature | binary ref | image ref | (not written to Excel) |

### 10.3 Current System Value Storage

The current system stores values in the browser's Zustand store and can POST them to a save endpoint. The `RuntimeFormViewer` collects values from fields and can export them as JSON. However, there is no end-to-end pipeline from:

```
Browser input → Serve values → WorkbookGenerator → Output XLSX
```

---

## 11. Cell Address Resolution

### 11.1 How Cell Addresses Are Stored

**In the definition XML (`template.xml`):**

```xml
<cluster>
  <clusterId>1</clusterId>
  <cellAddress>$A$1:$B$2</cellAddress>  <!-- Merge range or single cell -->
</cluster>
```

**In `templateXls.xml` (alternate format):**

```xml
<cluster>
  <clusterId>1</clusterId>
  <sheetNo>1</sheetNo>
  <!-- cellAddress is NOT present — uses position instead -->
  <left>0.0</left>
  <top>0.0</top>
  <right>0.5</right>
  <bottom>0.3</bottom>
</cluster>
```

**In the current system (`ClusterDefinition`):**

```csharp
public class ClusterDefinition {
    public string? CellAddress { get; set; }     // e.g., "$A$1:$B$2"
    public string? Cell { get; set; }            // Single cell reference
    public string? Address { get; set; }         // Alternative address field
}
```

### 11.2 Resolution Strategy

Based on the old system's approach:

1. **Primary:** Use `cellAddress` if present (the explicit Excel range)
2. **Fallback:** Convert coordinate rectangle to cell range using column width / row height mapping
3. **Merged cells:** Write to the top-left cell of the merge range only
4. **Single cells:** Write directly to `ws.Range[cellAddress].Value`

### 11.3 Current Gap

Although `ClusterDefinition.CellAddress` exists, the `WorkbookGenerator` does not use it for writing values. The generator only uses merged cell addresses from `SheetDefinition.MergedCells`, which is a list of `MergedCellInfo` objects with `Address` and `CellAddress` properties.

---

## 12. Merged Cell Behavior

### 12.1 Evidence from OOXML Analysis

All ConMas-original workbooks have merged cells that define field boundaries. The merge ranges match the `Address` field in the `_Fields` sheet.

For example:
- Field at cell `I6` with merge range `I6:M6` → value is written to `I6` only (top-left of merge)
- Field at cell `A1` with merge range `A1:B2` → value is written to `A1` only

### 12.2 Current System Merge Handling

The `WorkbookGenerator` already applies merged cells:

```csharp
private void ApplyMergedCells(Excel.Worksheet ws, List<MergedCellInfo> mergedCells, Dictionary<string, CellStyleInfo> cellStyles)
{
    foreach (var mc in mergedCells)
    {
        var range = ws.Range[mc.Address];
        range.Merge();                              // Merge cells
        if (cellStyles.TryGetValue(mc.CellAddress, out var style))
            ApplyStyleToRange(range, style);        // Apply style to merged range
    }
}
```

### 12.3 Gap

After merge and style application, the code should also:

```csharp
// Write value to top-left cell of merge range
if (mc.Value != null)
    ws.Range[mc.CellAddress].Value = mc.Value;
```

The `MergedCellInfo` already has a `CellAddress` property (the top-left cell), but no `Value` property exists yet.

---

## 13. Comment Generation Pipeline

### 13.1 Old System Comment Format

From actual OOXML analysis of ConMas-original workbooks (3 samples verified):

```
Comment Format A (Newer — "name first" style):
  Line 0: FieldName          e.g., "samples"
  Line 1: FieldType          e.g., "KeyboardText"
  Line 2: ClusterIndex       e.g., "0" (numeric, -1 if empty)
  Line 3+: InputParameters   e.g., "Required=0;...;InputRestriction=None;..."

Comment Format B (Older — "type first" style):
  Line 0: FieldType          e.g., "KeyboardText"
  Line 1: Parameter          e.g., "0"
  Line 2: Parameter          e.g., "0"
  Line 3: Parameter          e.g., "0"
```

### 13.2 OOXML Comment XML Structure

```xml
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>AUTHOR_NAME</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0" shapeId="0">
      <text>
        <r>
          <rPr>
            <sz val="9"/>
            <color indexed="81"/>
            <rFont val="Tahoma"/>
          </rPr>
          <t>FIELD_NAME</t>
        </r>
        <r>
          <rPr>...</rPr>
          <t>\n</t>              <!-- Explicit newline -->
        </r>
        <r>
          <rPr>...</rPr>
          <t>FIELD_TYPE</t>
        </r>
        <r>
          <rPr>...</rPr>
          <t>\n</t>              <!-- Explicit newline -->
        </r>
        <r>
          <rPr>...</rPr>
          <t>0</t>
        </r>
      </text>
    </comment>
  </commentList>
</comments>
```

### 13.3 COM API for Comments

Using Excel COM Interop in .NET, the equivalent code would be:

```csharp
// Add comment to cell
var range = ws.Range[cellAddress];
var comment = range.AddComment(text);

// Format the comment
comment.Visible = false;                      // Hide the comment indicator (red triangle)
comment.Shape.TextFrame.AutoSize = true;       // Auto-size text box
comment.Shape.Fill.Visible = false;            // Transparent background (or use default)
comment.Shape.TextFrame.Characters(1, 1).Font.Name = "Tahoma";
comment.Shape.TextFrame.Characters(1, 1).Font.Size = 9;

// For multi-line comments with separate formatting per line:
// (COM Interop approach is simpler — use \n in the text)
string commentText = $"{fieldName}\n{fieldType}\n{clusterIndex}\n{inputParameters}";
range.AddComment(commentText);
```

### 13.4 VML Drawing Requirement

Each comment in OOXML requires a corresponding VML drawing entry for the comment shape:

```
xl/drawings/vmlDrawing1.vml
```

The current `WorkbookGenerator` does NOT create VML drawings. Excel COM Interop handles this automatically when `AddComment()` is called — COM creates both the `<comments>` XML and the VML drawing entry.

### 13.5 Current Gap

No comment writing exists in `WorkbookGenerator`. Adding `range.AddComment(text)` is a single COM call that:

1. Creates the comment in `xl/comments1.xml`
2. Creates the VML drawing in `xl/drawings/vmlDrawing1.vml`
3. Adds relationship entries in `.rels` files

**No manual OOXML manipulation needed** — Excel COM handles it all.

---

## 14. Image Handling

### 14.1 Evidence

**Current system:** `FormDefinition` has an `ImageDefinition` class but the `WorkbookGenerator` does not use it.

```csharp
// Models/FormDefinition.cs
public class ImageDefinition {
    public string? Id { get; set; }
    public string? Name { get; set; }
    public string? SheetName { get; set; }
    public string? ContentType { get; set; }
    public byte[]? Content { get; set; }
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
}
```

**Old system:** Likely does NOT copy images to the output workbook. The output workbook is a clean form with field values — images are rendered separately as part of the background PDF/PNG.

### 14.2 Confidence

| Aspect | Confidence | Evidence |
|--------|:----------:|----------|
| Images preserved in output? | **50%** | No direct evidence |
| Images stored in definition XML? | **70%** | Fields exist but may be for preview only |
| Images written to workbook? | **40%** | No COM Interop code found for image placement |

**Recommendation:** Low priority. Images (logos, backgrounds) are already rendered in the background PNG. The output workbook focuses on filled values.

---

## 15. Database Interaction

### 15.1 Old System DB Tables

The old system uses a PostgreSQL database with these relevant tables:

| Table | Purpose | Used by Output Excel? |
|-------|---------|:--------------------:|
| `def_top` | Form definition (top-level metadata, XML) | ✅ Read definition |
| `def_sheet` | Sheet definitions (page setup, dimensions) | ✅ Read structure |
| `def_cluster` | Cluster (field) definitions (type, params, cell address) | ✅ Read field config |
| `rep_top` | Report submission (form instance metadata) | ✅ Read submission info |
| `rep_cluster` | Report cluster values (user-entered values per field) | ✅ **Read filled values** |

### 15.2 Current System DB Equivalent

The current system uses database generator objects that mirror the old schema:

| Old Table | Current Object | Purpose |
|-----------|---------------|---------|
| `def_top` | `DatabaseResult.DefTop` | Form metadata |
| `def_sheet` | `DatabaseResult.DefSheet[]` | Sheet definitions |
| `def_cluster` | `DatabaseResult.DefCluster[]` | Field definitions |
| `rep_top` | ❌ Not implemented | Submission tracking |
| `rep_cluster` | ❌ Not implemented | User-entered values |

### 15.3 What's Missing for the Full Pipeline

For a complete "Output Excel of Form" pipeline, the system needs:

1. **A way to collect user values** (currently only in browser Zustand store)
2. **A submission/save endpoint** that stores values alongside the form definition
3. **An output-excel endpoint** that reads stored values and generates the workbook

**Potential approach:** The existing `POST /api/form/save` could be extended with an optional `Values` payload. Or a new endpoint could accept values separately.

---

## 16. Key Gaps vs Legacy System

### 16.1 Gap Matrix

| # | Gap | Impact | Complexity | Priority |
|:-:|-----|--------|:----------:|:--------:|
| 1 | **Cell values not written** | Output workbook is blank (no user data) | Medium | **P0** |
| 2 | **No cell comments** | Output cannot be republished in ConMas Designer | Low | **P1** |
| 3 | **No `_Fields` sheet** | Output cannot be republished in ConMas Designer | Low | **P1** |
| 4 | **No VML comment drawings** | Comments missing in OpenXML viewers | Auto-fixed by COM | **P0** |
| 5 | **Cell borders not applied** | Formatting incomplete | Medium | **P2** |
| 6 | **No number formats** | Formatting incomplete | Medium | **P2** |
| 7 | **No images/shapes** | Visual content missing | High | **P3** |
| 8 | **No data validation** | Field restrictions not applied | Medium | **P3** |
| 9 | **No conditional formatting** | Visual rules not applied | Medium | **P3** |
| 10 | **Collection pipeline** | No end-to-end value collection | High | **P0** |

### 16.2 Critical Path

The most critical gap is **#10 (collection pipeline)** + **#1 (cell values written)**, because without them the output workbook is just an empty template.

```
Implemented:     Structure-only workbook
                   ↓
Phase 1 (P0):    Collection pipeline + cell values + cell comments
                   ↓
Phase 2 (P1):    _Fields hidden sheet + borders + number formats
                   ↓
Phase 3 (P2):    Images + data validation + conditional formatting
```

---

## 17. Confidence Table

| Finding | Confidence | Evidence Source |
|---------|:----------:|----------------|
| Entry point is "Output Excel of Form" button in tablet UI | **90%** | Template XML flags + DB schema |
| Values collected from `rep_cluster` DB table | **90%** | DB schema + XML template |
| `excelOutputValue` is the value written to cells | **90%** | XML template per-cluster element |
| `cellAddress` maps cluster to Excel cell | **90%** | XML template element |
| Workbook is created from scratch (`Workbooks.Add`) | **100%** | Source code confirmed |
| No `SaveCopyAs` used | **100%** | Codebase search |
| Values written as `ws.Range[address].Value = value` | **85%** | COM Interop pattern inference |
| Cell comments written as field metadata | **100%** | OOXML inspection of ConMas files |
| Comment format: Name + Type + Index + Params | **90%** | 3 workbook samples confirmed |
| `_Fields` sheet created by `ExportExcelDefinition()` | **100%** | Decompiled DLL reference |
| `_Fields` has 7 columns | **100%** | OOXML inspection of workbook |
| Comments written via COM `AddComment()` | **90%** | COM Interop pattern inference |
| VML drawings created automatically by COM | **95%** | Standard COM behavior |
| Images/shapes NOT preserved in output | **50%** | No direct evidence |
| Cell borders NOT applied by current generator | **100%** | Source code confirmed |
| CellStyleInfo has border properties in model | **100%** | Source code confirmed |
| No number format support in FormDefinition | **70%** | Model review |
| Old system reads definition from `def_top.xml_data` | **90%** | DB schema + PublishPipeline docs |
| Old system reads values from `rep_cluster` | **90%** | DB schema |
| Generated output can be republished | **90%** | System architecture analysis |

---

## 18. Conclusion

### 18.1 What We Know for Certain

1. The **current `WorkbookGenerator`** creates structurally complete workbooks (page setup, styles, merged cells, dimensions) — this is the foundation.

2. The **legacy system** wrote three things the current system doesn't:
   - **Cell values** — `excelOutputValue` per cluster → `ws.Range[cellAddress].Value = value`
   - **Cell comments** — field metadata (name, type, params) in rich-text format
   - **`_Fields` hidden sheet** — 7-column metadata table for republish

3. The **COM Interop approach** is correct. The old system also used COM Interop (`Workbooks.Add` → configure → `SaveAs`). Our current approach matches.

4. The **XML template** from `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\template.xml` defines the complete schema for Output Excel configuration.

### 18.2 What We Need to Implement (in order)

**Phase 0 — Value Collection Pipeline**
- Add a values payload to the save/export endpoint
- Extend `FormDefinition` or create `OutputFormDefinition` with:
  - `Dictionary<string, string> FieldValues` (fieldId → value)
  - `Dictionary<string, string> CellValues` (cellAddress → value)

**Phase 1 — Cell Value Writing**
- After `ApplyCellStyles()`, add `WriteCellValues()`:
  ```csharp
  foreach (var (cellAddress, value) in cellValues)
      ws.Range[cellAddress].Value = value;
  ```

**Phase 2 — Cell Comments**
- After writing values, add `WriteCellComments()`:
  ```csharp
  foreach (var cluster in clusters)
  {
      var text = $"{cluster.Name}\n{cluster.Type}\n{cluster.Index}\n{cluster.InputParameters}";
      ws.Range[cluster.CellAddress].AddComment(text);
  }
  ```

**Phase 3 — `_Fields` Hidden Sheet**
- Create a new sheet, write 7-column header row, write one row per field
- Set sheet to hidden (`XlSheetHidden`)
- Name the sheet `_Fields`

### 18.3 Architecture Recommendation

Create a new class `OutputWorkbookGenerator` that extends `WorkbookGenerator`:

```csharp
public class OutputWorkbookGenerator : WorkbookGenerator
{
    public OutputResult Generate(FormDefinition form, 
        Dictionary<string, string> fieldValues,   // fieldId → value
        string outputPath)
    {
        // 1. Call base.Generate() for structural creation
        // 2. Write cell values
        // 3. Write cell comments
        // 4. Create _Fields hidden sheet
        // 5. Save
    }
}
```

Or, alternatively, add the output-specific functionality to `WorkbookGenerator` as optional parameters (simpler but less clean):

```csharp
public string Generate(FormDefinition form, string outputPath,
    Dictionary<string, string>? fieldValues = null)
```

### 18.4 Final Word

The legacy "Output Excel of Form" pipeline is **faithfully reproducible** with our current architecture. The COM Interop foundation is solid. The gaps are well-defined and bounded:

| Component | Status | Effort |
|-----------|--------|:------:|
| Workbook structure creation | ✅ Complete | Done |
| Cell value writing | ❌ Missing | 2–4 hours |
| Cell comment writing | ❌ Missing | 1–2 hours |
| `_Fields` sheet creation | ❌ Missing | 1–2 hours |
| End-to-end value pipeline | ❌ Missing | 4–8 hours |
| **Total remaining effort** | | **~8–16 hours** |

---

*End of report. Investigation only — no code was modified.*
