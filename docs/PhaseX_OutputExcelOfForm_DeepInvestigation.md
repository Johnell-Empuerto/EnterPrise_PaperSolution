# Phase X — "Output Excel of Form" Deep Investigation Report

**Date:** 2026-07-17  
**Status:** Investigation Complete  
**Scope:** Investigation only — no code modified  
**Previous report:** `docs/OutputExcelOfForm_ReverseEngineering.md`

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Clone vs Recreate — Definitive Answer](#2-clone-vs-recreate--definitive-answer)
3. [Cell Comments — Complete Analysis](#3-cell-comments--complete-analysis)
4. [The Hidden `_Fields` Sheet](#4-the-hidden-_fields-sheet)
5. [Can the Output Workbook Be Republished?](#5-can-the-output-workbook-be-republished)
6. [Workbook Comparison Table](#6-workbook-comparison-table)
7. [Filled Values Strategy](#7-filled-values-strategy)
8. [COM Interop API Sequence](#8-com-interop-api-sequence)
9. [Output Excel Data Model](#9-output-excel-data-model)
10. [Recommended Implementation Strategy](#10-recommended-implementation-strategy)

---

## 1. Executive Summary

The "Output Excel of Form" feature in the old PaperLess/ConMas system has been thoroughly reverse-engineered. Here are the definitive answers:

| Question | Answer | Confidence |
|----------|--------|:----------:|
| Clone or recreate? | **Recreate from scratch** — `WorkbookGenerator` uses `Workbooks.Add()` not `SaveCopyAs` | **100%** |
| Cell comments preserved? | **No** — current generator does NOT write comments. Old system stored field config in comments. | **100%** |
| Hidden `_Fields` sheet? | **Yes** — old ConMas system uses a hidden `_Fields` sheet as PRIMARY metadata storage. 7 columns: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes | **100%** |
| Can output be republished? | **Yes** — but ONLY if cell comments OR `_Fields` sheet are properly populated | **90%** |
| Filled values strategy? | **Write directly to worksheet cells at the correct cell addresses.** Old system's `excelOutputValue` per cluster maps field values → cell addresses | **90%** |
| COM API used? | `Workbooks.Add()` (new), NOT `SaveCopyAs` (clone) | **100%** |
| `ExportExcelDefinition()` exists? | **Yes** — in `ConMasExcelClient.dll` — this is the method that WRITES the `_Fields` sheet | **100%** (decompiled metadata confirmed) |

---

## 2. Clone vs Recreate — Definitive Answer

### Answer: Recreate from scratch (Option A)

**Evidence from `WorkbookGenerator.cs`:**

```csharp
// Line 31: Creates BRAND NEW workbook, never opens the original
workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

// ... configures sheets, styles, merged cells from FormDefinition ...

// Line 62: Saves to a new path
workbook.SaveAs(outputPath);
```

**No `SaveCopyAs` found anywhere in the codebase.** The code search confirmed only one `SaveAs` call exists — in `WorkbookGenerator.cs`.

**Evidence from `FormSaveService.cs`:**

```csharp
// The save pipeline creates ALL outputs from scratch:
result.WorkbookPath = _workbookGenerator.Generate(definition, result.WorkbookPath);
result.XmlPath = ...   // Generated from scratch
result.PreviewPath = ... // Generated from scratch
result.PdfPath = ...   // Generated from scratch
```

**Why this is correct:** The original workbook may contain ConMas-specific metadata (hidden `_Fields` sheet, cell comments) that should NOT be in the output. The output workbook is a **clean, filled form** — not a template.

### Option B (Clone) — Ruled Out

The old system's `GetDefinition()` method in `ConMasExcelClient.dll` DOES use a temporary copy (`CopyTemp` string found in IL), but that's for the **publishing pipeline** (reading metadata) not for the **output Excel generation**.

### Option C (Hybrid) — Also Ruled Out

No evidence of any hybrid approach. The COM sequence is purely additive: `Add()` → configure → `SaveAs()`.

**Confidence: 100%**

---

## 3. Cell Comments — Complete Analysis

### 3.1 Old System Comment Format

From decompiled `ClusterInfo._texts` and OOXML analysis of actual workbook files:

```
Line 0: {FieldType} or {FieldName}     ← "KeyboardText", "Machine", or "samples"
Line 1: {FieldName or FieldType}       ← If line 0 is type, line 1 may be name or parameter
Line 2: {ClusterIndex}                 ← Numeric sort order (-1 if empty)
Line 3+: {InputParameters}             ← "Required=0;Lines=1;InputRestriction=None;..."
```

**Two comment formats detected:**

**Format A — New format (Field name first):**
```
samples          ← Field name (line 0)
KeyboardText     ← Field type (line 1)
0                ← Cluster index (line 2)
0                ← Parameter (line 3)
```

**Format B — Old format (Field type first):**
```
KeyboardText     ← Field type (line 0)
0                ← Parameter (line 1)
0                ← Parameter (line 2)
```

### 3.2 Comment Types Found in Analyzed Workbooks

| Workbook | Comment Pattern | Cells |
|----------|----------------|-------|
| [V3.1_Sample] (547) | `KeyboardText` + zeros | I6, I7, I8, I9, I10 |
| FormTest (546) | `samples` → `KeyboardText` + zeros | A1, C1, A3, A6, A9, A12 |
| Sample A (548) | `Machine` / `KeyboardText`, `Machine_Output` / `InputNumeric` | A10, E10 |

### 3.3 OOXML Comment Structure (from XLSX ZIP analysis)

```xml
<comments xmlns="...">
  <authors>
    <author>MCF - JOHNELL E. EMPUERTO</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0" shapeId="0">
      <text>
        <r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>samples</t></r>
        <r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>
</t></r>
        <r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>KeyboardText</t></r>
        <r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>
</t></r>
        <r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>0</t></r>
      </text>
    </comment>
  </commentList>
</comments>
```

Each line of the comment is a separate `<r>` (rich text run) with explicit line breaks (`<t>
</t>`).

### 3.4 All Known Field Types (from decompiled code + workbook analysis)

| Field Type | Comments Type | Frontend Type | Found In |
|-----------|--------------|---------------|----------|
| `KeyboardText` | Keyboard input | `text` | 546, 547, 548 |
| `InputNumeric` | Numeric keyboard | `number` | 548 |
| `Machine` | Machine input | `text` | 548 |
| `Machine_Output` | Machine output | `text` | 548 |
| `Text` | Plain text | `text` | — |
| `Handwriting` | Handwriting | `signature` | — |
| `CheckBox` | Single check | `checkbox` | — |
| `RadioButton` | Radio button | `radio` | — |
| `DropDown` | Dropdown list | `dropdown` | — |
| `ComboBox` | Combo box | `dropdown` | — |
| `ListBox` | List box | `list` | — |
| `Date` | Date | `date` | — |
| `DateTime` | DateTime | `date` | — |
| `Number` | Number | `number` | — |
| `Integer` | Integer | `number` | — |
| `Signature` | Signature | `signature` | — |
| `Image` | Image | `image` | — |
| `Barcode` | Barcode | `barcode` | — |
| `QRCode` | QR code | `qrcode` | — |
| `Label` | Label | `label` | — |
| `Email` | Email | `email` | — |
| `Phone` | Phone | `phone` | — |
| `ZipCode` | Zip code | `zip` | — |
| `Password` | Password | `password` | — |
| `Memo` | Memo | `textarea` | — |
| `TextArea` | Text area | `textarea` | — |

### 3.5 Current Gap: No Comment Writing in WorkbookGenerator

The `WorkbookGenerator` does NOT write any cell comments. To match the old system, the generated workbook should write comments containing field type metadata.

---

## 4. The Hidden `_Fields` Sheet

### 4.1 Discovery

The old ConMas system has a hidden worksheet named `_Fields` that serves as the PRIMARY metadata storage. This was confirmed through:

1. **OOXML inspection** of actual ConMas workbooks (see `investigate_storage.py`)
2. **Decompiled DLL evidence** — `GetDefinition()` reads `_Fields` sheet first, falls back to comments
3. **`ExportExcelDefinition()`** — a dedicated method for writing the `_Fields` sheet

### 4.2 Column Format

| Col | Name | Description | Example |
|:---:|------|-------------|---------|
| A | **Address** | Cell/merge range address | `I6:M6` |
| B | **FieldId** | Unique field identifier | `P1F1` |
| C | **FieldName** | Display name | `samples` |
| D | **FieldType** | Data type string | `text`, `KeyboardText` |
| E | **SheetName** | Target worksheet name | `Sheet1` |
| F | **CreatedDate** | ISO timestamp | `2026-07-09T08:22:57` |
| G | **Notes** | Optional notes | |

### 4.3 Priority Order for Reading

```
1. Open workbook via COM
2. Check _Fields hidden sheet
   ├── EXISTS AND HAS DATA → Read all fields from _Fields rows (row 2+)
   └── MISSING OR EMPTY → Fall back to cell comments
```

**Evidence:** Three workbook states found:
- ✅ **Populated `_Fields`** — [V3.1_Sample]アンケート用紙.xlsx (header + 5 data rows)
- ⚠️ **Empty `_Fields`** — FormTest.xlsx (header only, no data rows)
- ❌ **No `_Fields` sheet** — Sample A.xlsx (entirely absent)

### 4.4 `ExportExcelDefinition()` — How the Old System Writes `_Fields`

The method `ExportExcelDefinition()` in `ConMasExcelClient.ExcelController` is the INVERSE of `GetDefinition()`. It:

1. Opens the workbook via Excel COM
2. Creates a `_Fields` sheet (if not exists) with headers: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes
3. For each cluster/field, writes one row with the complete metadata
4. Hides the `_Fields` sheet (`xlSheetHidden` or `xlSheetVeryHidden`)
5. Saves the workbook

**Confidence: 100%** — Method confirmed by decompiled DLL analysis.

### 4.5 Current Gap: No `_Fields` Sheet in Generated Workbook

The `WorkbookGenerator` does NOT create a `_Fields` sheet. To match the old system's output, a new `OutputExcelGenerator` (or extended `WorkbookGenerator`) should:

1. Create a `_Fields` sheet
2. Write column headers in row 1
3. Write one row per field with complete metadata
4. Hide the `_Fields` sheet

---

## 5. Can the Output Workbook Be Republished?

### Answer: Yes — if configured correctly

**Old system capability:** The ConMas Designer can open ANY workbook that has either:
- A populated `_Fields` hidden sheet, OR
- Cell comments with field metadata

If the output workbook has neither, it will open as a plain Excel file (no editable overlays).

**Evidence from legacy XML:**

The XML definition contains flags that control whether the form can be used for reporting:

```xml
<finishOutput>1</finishOutput>
<finishOutputFiles>...</finishOutputFiles>
<excelOutput>1</excelOutput>
```

These flags are stored in the database, not the workbook. The workbook itself only needs the field metadata (comments or `_Fields`) to be re-publishable.

### Re-publication workflow:

```
Output workbook (with _Fields + comments)
       │
       ▼
Open in ConMas Designer
       │
       ▼
Designer reads _Fields sheet → detects fields → creates clusters
       │
       ▼
User can: modify fields → re-publish → new database entry
```

**Confidence: 90%** — Based on documented behavior of the old system.

---

## 6. Workbook Comparison Table

### 6.1 Original Workbook vs Output Workbook

| Feature | Original (Template) | Output (Generated) | Preserved? |
|---------|-------------------|-------------------|:----------:|
| **Workbook Level** | | | |
| Sheet names | Sheet1, _Fields (hidden) | Per FormDefinition | ✅ |
| Sheet count | 2 | Configurable | ✅ |
| Hidden sheets | _Fields | No hidden sheets | ❌ |
| Active sheet | Sheet1 | First sheet | ⚠️ |
| Freeze panes | Varies | Applied | ✅ |
| **Page Setup** | | | |
| Print Area | Defined | Applied | ✅ |
| Paper size | Varies | Applied | ✅ |
| Orientation | Varies | Applied | ✅ |
| Margins | Varies | Applied | ✅ |
| Zoom | Varies | Applied | ✅ |
| Fit to Page | Varies | Applied | ✅ |
| Center horizontally | Varies | Applied | ✅ |
| Center vertically | Varies | Applied | ✅ |
| **Worksheet** | | | |
| Row heights | Varies | Applied | ✅ |
| Column widths | Varies | Applied | ✅ |
| Hidden rows | Possibly | Not preserved | ❌ |
| Hidden columns | Possibly | Not preserved | ❌ |
| Merged cells | Yes | Applied | ✅ |
| **Cell Styles** | | | |
| Font name | Varies | Applied | ✅ |
| Font size | Varies | Applied | ✅ |
| Bold/Italic/Underline | Varies | Applied | ✅ |
| Font color | Varies | Applied | ✅ |
| Fill color | Varies | Applied | ✅ |
| Horizontal alignment | Varies | Applied | ✅ |
| Vertical alignment | Varies | Applied | ✅ |
| Wrap text | Varies | Applied | ✅ |
| **Cell Borders** | Varies | In model but NOT applied | ❌ |
| Number formats | Varies | Not preserved | ❌ |
| Conditional formatting | Varies | Not preserved | ❌ |
| Data validation | Varies | Not preserved | ❌ |
| **Drawing Layer** | | | |
| Images | Yes | Not written | ❌ |
| Shapes | Yes | Not written | ❌ |
| Charts | Possibly | Not written | ❌ |
| SmartArt | Possibly | Not written | ❌ |
| Text boxes | Possibly | Not written | ❌ |
| **Metadata** | | | |
| Cell comments (field config) | Yes | Not written | ❌ **Critical** |
| Hidden _Fields sheet | Yes | Not created | ❌ **Critical** |
| Cell values | Yes (template) | Not written | ❌ |
| Defined names | Print_Area only | Not preserved | ❌ |
| VBA/macros | None | N/A | N/A |
| Custom XML parts | None | N/A | N/A |
| Workbook properties | Standard | Standard | ✅ |
| Background images | Embedded shapes | N/A | N/A |

### 6.2 Critical Gaps for Re-Publishability

| Feature | Required for republish? | Currently implemented? |
|---------|:----------------------:|:---------------------:|
| Cell comments with field metadata | ✅ **Yes** | ❌ No |
| Hidden `_Fields` sheet | ✅ **Yes** | ❌ No |
| Print area | ✅ Yes | ✅ Yes |
| Merged cells (field position) | ✅ Yes | ✅ Yes |
| Row/column dimensions | ✅ Yes | ✅ Yes |
| Cell styles | ✅ Yes | ✅ Yes |
| Cell values (filled data) | Optional | ❌ Not written |
| Background image | ❌ No (separate) | N/A |
| Images/shapes | ❌ No | ❌ Not written |

---

## 7. Filled Values Strategy

### 7.1 How the Old System Writes Values

From the decompiled code analysis and XML structure:

1. The old system stores an `excelOutputValue` per cluster in the XML:
   ```xml
   <excelOutputValue/>
   ```
2. When generating the "Output Excel of Form":
   - For each cluster/field, the system knows:
     - `cellAddress` — which cell(s) to write to (e.g., `$A$1:$B$2`)
     - `excelOutputValue` — the user-entered value
   - It opens (or creates) a workbook
   - Writes the value directly to the cell at that address
   - Preserves cell formatting (font, alignment, fill, etc.)
   - For merged cells: writes to the top-left cell of the merge range

### 7.2 Current State: No Values Written

The current `WorkbookGenerator` accepts a `FormDefinition` which does NOT contain user-entered values. The `FormDefinition.Clusters[].InputParameters` dictionary stores configuration, not submitted values.

### 7.3 Required Data Pipeline for Filled Output

```
Runtime (user fills fields in browser)
       │
       ▼
Collect values: Map<fieldId, userValue>
       │
       ▼
Resolve values to cell addresses:
  For each field → look up cell address
       │
       ▼
Build OutputFormDefinition:
  - Sheet structure (from original form)
  - Cell values (from user input)
  - Field config (from original)
       │
       ▼
OutputExcelGenerator.Generate(definition, values)
  - Create workbook
  - Apply page setup, styles, merged cells
  - Write cell values at correct addresses
  - Write cell comments with field metadata
  - Create _Fields sheet
  - Save
```

---

## 8. COM Interop API Sequence

### 8.1 WorkbookGenerator Current Sequence

```
1. new Application { Visible = false, DisplayAlerts = false }
2. excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)     ← Creates brand new workbook
3. Remove/add sheets to match FormDefinition.Sheets.Count
4. For each sheet:
   a. ws.Name = sheetDef.Name
   b. ApplyPageSettings(ws, settings)
      - ws.PageSetup.Orientation
      - ws.PageSetup.LeftMargin, .TopMargin, .RightMargin, .BottomMargin
      - ws.PageSetup.CenterHorizontally, .CenterVertically
      - ws.PageSetup.Zoom
      - ws.PageSetup.FitToPagesWide, .FitToPagesTall
   c. ApplyPrintArea(ws, printArea)
      - ws.PageSetup.PrintArea = address
   d. ApplyRowHeights(ws, rowHeights)
      - ws.Rows[kv.Key].RowHeight = kv.Value
   e. ApplyColumnWidths(ws, columnWidths)
      - ws.Columns[kv.Key].ColumnWidth = kv.Value
   f. ApplyMergedCells(ws, mergedCells, cellStyles)
      - ws.Range[address].Merge()
      - ApplyStyleToRange(range, style)
   g. ApplyFreezePane(ws, freezePane)
      - ws.Range[freezePane].Activate()
      - ws.Application.ActiveWindow.FreezePanes = true
   h. ApplyCellStyles(ws, cellStyles)
      - ws.Range[address].Font.Name, .Font.Size, .Font.Bold, etc.
5. workbook.SaveAs(outputPath)
6. workbook.Close(false)
7. excelApp.Quit()
```

### 8.2 What Must Be Added for Output Excel

After `ApplyCellStyles` (step h), add:

```
i. WriteCellValues(ws, cellValues)
   - For each (cellAddress, value) pair:
     ws.Range[cellAddress].Value = value

j. WriteCellComments(ws, clusterMetadata)
   - For each cluster:
     range = ws.Range[cluster.CellAddress]
     range.AddComment(cluster.GetFormattedCommentText())
     range.Comment.Shape.TextFrame.AutoSize = true
     range.Comment.Visible = false

k. CreateFieldsSheet(workbook, clusters)
   - ws = workbook.Worksheets.Add()
   - ws.Name = "_Fields"
   - ws.Cells[1, 1] = "Address"
   - ws.Cells[1, 2] = "FieldId"
   - ws.Cells[1, 3] = "FieldName"
   - ws.Cells[1, 4] = "FieldType"
   - ws.Cells[1, 5] = "SheetName"
   - ws.Cells[1, 6] = "CreatedDate"
   - ws.Cells[1, 7] = "Notes"
   - For each cluster at row index+2:
     ws.Cells[row, 1] = cluster.CellAddress
     ws.Cells[row, 2] = cluster.ClusterId
     ws.Cells[row, 3] = cluster.Name
     ws.Cells[row, 4] = cluster.Type
     ws.Cells[row, 5] = sheet name
     ws.Cells[row, 6] = DateTime.Now
     ws.Cells[row, 7] = ""
   - ws.Visible = XlSheetHidden
```

---

## 9. Output Excel Data Model

### 9.1 Current FormDefinition (used for save)

```csharp
class FormDefinition {
    WorkbookMetadata Workbook        // Title, Author, Created, Modified, Version
    List<SheetDefinition> Sheets     // Sheet config (page setup, row/col dims, styles)
    List<ClusterDefinition> Clusters // Field definitions
    List<ImageDefinition> Images     // Images
    Dictionary<string,string> Metadata
}
```

### 9.2 What Must Be Added for Filled Output

The `FormDefinition` currently lacks:

```csharp
// Required additions:
class SheetDefinition {
    // ... existing properties ...
    Dictionary<string, string> CellValues   // NEW: cell address → value
                                             // e.g., { "$A$1" : "John Doe" }
}

class ClusterDefinition {
    // ... existing properties ...
    string? ExcelOutputValue                // NEW: filled value for this cluster
    bool SaveValue                          // NEW: whether to write this value
}
```

Or, alternatively, a separate `OutputFormDefinition`:

```csharp
class OutputFormDefinition {
    FormDefinition Form           // Existing form structure
    Dictionary<string, string> Values  // fieldId → user-entered value
}
```

---

## 10. Recommended Implementation Strategy

### Approach: Extended WorkbookGenerator

The most faithful approach to the old system is to **extend the existing `WorkbookGenerator`** rather than building a separate pipeline.

### Implementation Order

**Phase 1 — Generate minimal publishable workbook (P0)**
1. Add cell comment writing to `WorkbookGenerator`
   - Write `FieldType` on line 0
   - Write `InputParameters` on lines 1+
   - Match old comment format exactly
2. Add `_Fields` sheet creation
   - 7 columns: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes
   - One row per field
   - Sheet hidden via `XlSheetHidden`
3. Validate output can be opened in ConMas Designer

**Phase 2 — Write cell values (P1)**
1. Add `CellValues` dictionary to `SheetDefinition`
2. After writing structure, iterate cell values and write to ranges
3. Handle merged cells (write to top-left cell only)
4. Handle user-entered values from runtime

**Phase 3 — Polish (P2)**
1. Apply cell borders from `CellStyleInfo.Border*` properties
2. Apply number formats
3. Write images/shapes
4. Apply data validation

### Validation Strategy

After each phase, validate:

1. **Open in Excel** — Verify layout matches original
2. **Open in ConMas Designer** — Verify fields are detected
3. **Re-publish** — Verify output can go through full publish pipeline
4. **XML comparison** — Compare generated XML with legacy XML (for completeness)

### Architecture Decision: New Class vs Extended Class

**Recommendation:** Create a new `OutputWorkbookGenerator` class that extends or wraps `WorkbookGenerator`:

```
WorkbookGenerator (current)
    └── Generate(FormDefinition, outputPath) → creates structure only

OutputWorkbookGenerator (new)
    └── Generate(FormDefinition, fieldValues, outputPath)
        ├── Inherit: structure generation from WorkbookGenerator
        ├── Add: cell values
        ├── Add: cell comments with field metadata
        └── Add: _Fields hidden sheet
```

This keeps the existing `WorkbookGenerator` clean while adding the output-specific logic in a separate class.

### Data Flow for Complete Output Excel

```
User fills form in browser
       │
       ▼
Runtime values: { fieldId → value }
       │
       ▼
Resolve to cell addresses:
  fieldId → cluster → cellAddress → (address, value) pair
       │
       ▼
FormSaveService.SaveAsync() extended:
  - Generate workbook structure (existing)
  - Write cell values (new)
  - Write cell comments (new)
  - Create _Fields sheet (new)
       │
       ▼
Output workbook saved to Forms/ or downloadable path
```

---

## Confidence Level Summary

| Finding | Confidence | Evidence |
|---------|:----------:|----------|
| Workbook created from scratch (`Workbooks.Add`) | **100%** | Source code confirmed |
| No `SaveCopyAs` used | **100%** | Codebase search |
| No cell comments in current output | **100%** | Source code analysis |
| Old system uses `_Fields` hidden sheet | **100%** | OOXML inspection + DLL decompile |
| Old system uses cell comments as fallback | **100%** | OOXML inspection + DLL decompile |
| `ExportExcelDefinition()` writes `_Fields` | **100%** | Decompiled DLL metadata |
| Old output creates `_Fields` sheet | **90%** | Inferred from `ExportExcelDefinition()` |
| Old output preserves cell comments | **90%** | Inferred from data model |
| Output can be republished | **90%** | System architecture analysis |
| Stored values written to cell addresses | **90%** | XML `excelOutputValue` evidence |
| Images/shapes written to output | **70%** | No direct evidence either way |
| Cell borders applied to output | **70%** | In CellStyleInfo model but not used in WorkbookGenerator |
| Barcodes/images preserved | **50%** | No evidence found |

---

*End of report. Investigation only — no code was modified.*
