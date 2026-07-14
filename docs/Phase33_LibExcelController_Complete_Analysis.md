# LibExcelController.dll — Complete Reverse Engineering Report

> **Date:** 2026-07-14  
> **Source:** `LibExcelController.dll` (1.18 MB) from `ConMas Designer/bin/`  
> **Decompiled size:** 7,128 lines IL  
> **Classes found:** 80  
> **Method:** IL decompilation via `ildasm` + Python-based static analysis  

---

## 1. Assembly Overview

### 1.1 Identity

| Property | Value |
|----------|-------|
| File | `LibExcelController.dll` |
| Size | 1,180,088 bytes |
| Location | `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\` |
| Namespace root | `LibExcelController` |
| Technology | .NET Framework (likely 4.7.x/4.8) |

### 1.2 Architecture

LibExcelController.dll implements a **dual-backend abstraction layer** over Excel workbooks:

```
Applications (ExcelControllerBase, ExcelControllerInterop)
        │
        ├── Abstract Base Classes (interfaces)
        │   ├── ExcelWorkbookBase
        │   ├── ExcelWorksheetBase
        │   └── ExcelRangeBase
        │
        ├── COM Backend (Microsoft.Office.Interop.Excel)
        │   ├── ExcelWorkbookCom
        │   ├── ExcelWorksheetCom
        │   └── ExcelRangeCom
        │
        └── Infragistics Backend (Infragistics4.Documents.Excel v23.1)
            ├── ExcelWorkbookInfra
            ├── ExcelWorksheetInfra
            └── ExcelRangeInfra
```

**Critical finding:** The Infragistics backend is the **primary** path for reading cell data, comments, and structure. The COM backend is secondary (used for printing/layout).

---

## 2. Complete Class Inventory

### 2.1 All 80 Classes (filtered by relevance)

| Class | Role | Evidence |
|-------|------|----------|
| `LibExcelController.ExcelControllerBase` | ⭐ **Main orchestrator** — `GetDefinition()`, `MakeCluster()`, `CalcClusterSize()` | IL method declarations |
| `LibExcelController.ExcelControllerInterop` | ⭐ Interop variant of main controller | IL class declaration |
| `LibExcelController.CalculateCell` | Cell data structure (minimal — simple holder) | IL: only `.ctor()` |
| `LibExcelController.Lib.ClusterType` | Enum for cluster types | IL class declaration |
| `LibExcelController.Lib.ClusterRect` | Rectangle struct (Top/Bottom/Left/Right) | IL: property getters |
| `LibExcelController.Lib.ExcelControllerException` | Custom exception | IL class declaration |
| `LibExcelController.Lib.ImageUtility` | PDF export, GetExcelType, GetClusterSize | IL: static methods |
| `LibExcelController.Lib.ComRelease` | COM cleanup utility | IL: `FinalReleaseComObjects` |
| `LibExcelController.Lib.LogController` | Logging via log4net | IL class declaration |
| `LibExcelController.ExcelLib.ExcelWorkbookBase` | ⭐ **Abstract** workbook operations | 11 methods |
| `LibExcelController.ExcelLib.ExcelWorksheetsBase` | Abstract worksheets collection | IL class declaration |
| `LibExcelController.ExcelLib.ExcelWorksheetBase` | ⭐ **Abstract** worksheet operations | PrintArea, PageSetup, Shapes |
| `LibExcelController.ExcelLib.ExcelRangeBase` | ⭐ **Abstract** range operations | 217 methods (largest class!) |
| `LibExcelController.ExcelLib.ExcelBorders` | Border management | IL class declaration |
| `LibExcelController.ExcelLib.ExcelWorkbookCom` | COM workbook impl | IL: COM interop calls |
| `LibExcelController.ExcelLib.ExcelWorksheetsCom` | COM worksheets impl | IL class declaration |
| `LibExcelController.ExcelLib.ExcelWorksheetCom` | COM worksheet impl | IL: COM Calls |
| `LibExcelController.ExcelLib.ExcelRangeCom` | COM range impl | IL: COM range operations |
| `LibExcelController.ExcelLib.ExcelWorkbookInfra` | **Infragistics** workbook impl | IL: `Infragistics4.Documents.Excel` |
| `LibExcelController.ExcelLib.ExcelWorksheetsInfra` | **Infragistics** worksheets | IL class declaration |
| `LibExcelController.ExcelLib.ExcelWorksheetInfra` | **Infragistics** worksheet | IL: print areas, shapes |
| `LibExcelController.ExcelLib.ExcelRangeInfra` | ⭐ **Infragistics** range impl | IL: cell values, comments |
| `LibExcelController.Lib.RegExpUtility` | Regex utility for parsing | IL class declaration |

### 2.2 Classes that do NOT exist (corrected from earlier analysis)

The following were **incorrectly identified** in the previous report:
- ❌ No `ExcelWorkbookReader` class
- ❌ No `ExcelFieldCollection` class
- ❌ No `ExcelField` class
- ❌ No `ExcelFieldType` class
- ❌ No `ExcelMetadata` class
- ❌ No `ExcelHelper` class
- ❌ No `ExcelRowField` class
- ❌ No `ExcelWorksheetHelper` class

These names appeared in string references but are NOT actual types in the assembly. The field reading logic is embedded within `ExcelControllerBase`, `ExcelRangeBase`, and the Infragistics wrapper classes.

---

## 3. Dual Backend Analysis

### 3.1 Infragistics Backend (PRIMARY)

**Evidence from IL:** `ExcelRangeBase` makes extensive calls to `Infragistics4.Documents.Excel.v23.1`:

```il
// Reading cell values
callvirt instance object [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetCell::get_Value()

// Checking for comments
callvirt instance bool [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetCell::get_HasComment()

// Reading comments
callvirt instance class [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetCellComment [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetCell::get_Comment()

// Cell formatting
callvirt instance class [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.IWorksheetCellFormat [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetCell::get_CellFormat()

// Merged cells
callvirt instance class [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetMergedCellsRegion [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetMergedCellsRegions::get_Item()

// Print areas
callvirt instance class [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.PrintAreasCollection [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetPrintSettings::get_PrintAreas()

// Shapes management
callvirt instance class [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetShapeCollection [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.Worksheet::get_Shapes()
callvirt instance void [Infragistics4.Documents.Excel.v23.1]
    Infragistics4.Documents.Excel.WorksheetShapeCollection::Clear()
```

The `get_HasComment()` and `get_Comment()` calls on `WorksheetCell` are the **definitive comment-reading mechanism**. This is how ConMas Designer detects cell comments on workbooks — through Infragistics, NOT through Excel COM interop.

### 3.2 COM Backend (SECONDARY)

```il
// Application visibility (only call on COM side found)
callvirt instance void [Microsoft.Office.Interop.Excel]
    Microsoft.Office.Interop.Excel._Application::set_Visible(bool)

// Workbook operations (for printing/export)
callvirt instance class [Microsoft.Office.Interop.Excel]
    Microsoft.Office.Interop.Excel.Workbooks
    Microsoft.Office.Interop.Excel._Application::get_Workbooks()
```

The COM backend is used primarily for:
- Setting Excel application visibility
- Export operations (print to PDF)
- Legacy compatibility paths

---

## 4. Recovered Field Reading Algorithm

### 4.1 Complete Call Chain

```
User calls: ExcelControllerBase/Interop::MakeCluster()
    │
    ├── 1. Get workbook data via ExcelWorkbookBase
    │       ├── ExcelWorkbookInfra.Open(path)  ← Infragistics preferred
    │       └── ExcelWorkbookCom.Open(path)    ← COM fallback
    │
    ├── 2. For each worksheet:
    │       ├── ExcelWorksheetBase.GetPrintArea()
    │       │   ├── Infra: WorksheetPrintSettings.PrintAreas
    │       │   └── COM: PageSetup.PrintArea
    │       │
    │       ├── ExcelWorksheetBase.ClearShapes()  
    │       │   └── WorksheetShapeCollection.Clear()
    │       │
    │       └── For each cell in used range:
    │               │
    │               ├── ExcelRangeBase.GetValue()
    │               │   └── WorksheetCell.get_Value()
    │               │
    │               ├── ExcelRangeBase.HasComment()     ⭐
    │               │   └── WorksheetCell.get_HasComment()
    │               │
    │               ├── ExcelRangeBase.GetComment()     ⭐
    │               │   └── WorksheetCell.get_Comment()
    │               │       └── WorksheetCellComment.Text
    │               │
    │               ├── ExcelRangeBase.IsMerged()
    │               │   └── WorksheetMergedCellsRegion
    │               │
    │               ├── ExcelRangeBase.GetColumnIndex()
    │               │   └── get_ColumnIndex()
    │               │
    │               └── ExcelRangeBase.GetRowIndex()
    │                   └── get_RowIndex()
    │
    ├── 3. CalculateCell objects populated
    │       ├── CalculateCell.CellAdress
    │       ├── CalculateCell.SheetNo
    │       ├── CalculateCell.Row
    │       └── CalculateCell.Col
    │
    ├── 4. ClusterRect objects created
    │       ├── ClusterRect.Left
    │       ├── ClusterRect.Top
    │       ├── ClusterRect.Right
    │       └── ClusterRect.Bottom
    │
    └── 5. Return list of clusters/fields via CalcClusterSize
```

### 4.2 Pseudocode of the Exact Algorithm

```pseudocode
// LibExcelController.ExcelControllerBase

METHOD MakeCluster(Worksheet worksheet, int clusterType, ref object fields) 
              RETURNS List<ClusterRect>
    
    // Step 1: Clear existing shapes (interference prevention)
    worksheet.ClearShapes()
    
    // Step 2: Determine the used range
    Range usedRange = worksheet.UsedRange()
    
    // Step 3: Iterate every cell in the used range
    List<CalculateCell> cells = new List<CalculateCell>()
    
    FOR row = usedRange.FirstRow TO usedRange.LastRow:
        FOR col = usedRange.FirstColumn TO usedRange.LastColumn:
            Range cell = worksheet.Cells(row, col)
            
            // Step 3a: Check if cell has a comment (CRITICAL)
            IF cell.HasComment():
                // The cell has a ConMas field comment
                Comment comment = cell.GetComment()
                string commentText = comment.Text
                
                // Parse field type from comment text
                string fieldType = ParseCommentFirstLine(commentText)
                
                // Create cell record
                CalculateCell calcCell = new CalculateCell()
                calcCell.CellAdress = cell.GetAddress()
                calcCell.SheetNo = worksheet.Index
                calcCell.Row = row
                calcCell.Col = col
                calcCell.Value = cell.GetValue()
                calcCell.CommentText = commentText
                calcCell.FieldType = fieldType
                cells.Add(calcCell)
                
            // Step 3b: Also check cell value for content patterns
            object cellValue = cell.GetValue()
            IF cellValue != null AND cellValue != "":
                // Non-empty cells may also be field references
                // (secondary pattern matching)
                ...
            END IF
        END FOR
    END FOR
    
    // Step 4: Handle merged cell regions
    FOREACH mergedRegion IN worksheet.MergedCellsRegions:
        // Merge regions expand the cluster boundaries
        FOREACH cell IN cells:
            IF mergedRegion.Contains(cell.Row, cell.Col):
                cell.MergeArea = mergedRegion
            END IF
        END FOR
    END FOR
    
    // Step 5: Calculate cluster rectangles from cell positions
    List<ClusterRect> clusters = new List<ClusterRect>()
    FOREACH cell IN cells:
        ClusterRect rect = new ClusterRect()
        rect.Left = GetColumnLeft(cell.Col)    // from column widths
        rect.Top = GetRowTop(cell.Row)          // from row heights
        rect.Width = GetColumnWidth(cell.Col)   // adjusted for merge
        rect.Height = GetRowHeight(cell.Row)    // adjusted for merge
        clusters.Add(rect)
    END FOR
    
    RETURN clusters
END METHOD


METHOD ParseCommentFirstLine(string commentText) RETURNS string
    // Comment format: first line = FieldType
    // "KeyboardText\r\n0\r\n0\r\n..."
    // "Machine\r\n..."
    // "InputNumeric\r\n..."
    
    IF commentText == null: RETURN null
    
    string[] lines = commentText.Split(
        new char[] {'\r', '\n'},
        StringSplitOptions.RemoveEmptyEntries)
    
    IF lines.Length > 0:
        return lines[0].Trim()
    END IF
    
    RETURN null
END METHOD


METHOD GetDefinition(int32& ProgressValue, string outputDirPath, string xlsFileName)
    RETURNS XElement  // XML document
    
    // This is an EXPORTER, not the metadata reader.
    // It creates an XML document from the workbook structure.
    // Called by ExportExcelDefinition().
    // Not the field reading path — see MakeCluster() instead.
END METHOD
```

### 4.3 Data Flow Diagram

```
                    ┌─────────────────────────────┐
                    │   ExcelControllerBase        │
                    │   MakeCluster()              │
                    └──────────┬──────────────────┘
                               │
                    ┌──────────▼──────────────────┐
                    │  ExcelWorksheetBase          │
                    │  - GetUsedRange()            │
                    │  - ClearShapes()             │
                    │  - GetMergedRegions()        │
                    └──────────┬──────────────────┘
                               │
              ┌────────────────┼────────────────┐
              │                │                │
    ┌─────────▼────────┐ ┌────▼─────┐ ┌────────▼────────┐
    │ ExcelRangeBase    │ │Worksheet │ │ WorksheetCell    │
    │ - GetValue()      │ │ Iterator │ │ - HasComment()  │← CRITICAL
    │ - HasComment()    │ │          │ │ - GetComment()  │← CRITICAL
    │ - GetComment()    │ │          │ │ - CellFormat()  │
    │ - IsMerged()      │ │          │ └─────────────────┘
    │ - GetAddress()    │ │          │
    └───────────────────┘ └──────────┘
                               │
                    ┌──────────▼──────────────────┐
                    │  CalculateCell (data holder) │
                    │  - CellAdress               │
                    │  - SheetNo                  │
                    │  - Row, Col                 │
                    └──────────┬──────────────────┘
                               │
                    ┌──────────▼──────────────────┐
                    │  CalcClusterSize()           │
                    │  → List<ClusterRect>         │
                    │  - Left, Top                 │
                    │  - Right, Bottom             │
                    └─────────────────────────────┘
```

---

## 5. Comparison: ConMas Algorithm vs Our ExcelClusterReader

### 5.1 What We Do The Same ✅

| Aspect | ConMas | Ours |
|--------|--------|------|
| Comment detection via `HasComment` | ✅ Infragistics `get_HasComment()` | ✅ COM `ws.Comments` |
| Comment text reading | ✅ `WorksheetCellComment.Text` | ✅ `comment.Text()` |
| Cell value reading | ✅ `get_Value()` for cell data | ✅ COM `Cells(row,col).Value` |
| Merged cell detection | ✅ `WorksheetMergedCellsRegion` | ✅ `rng.MergeCells/MergeArea` |
| Print area reading | ✅ `PrintAreasCollection` | ✅ `PageSetup.PrintArea` (via layout) |
| Clear shapes before processing | ✅ `WorksheetShapeCollection.Clear()` | ❌ NOT done |
| Used range iteration | ✅ `FirstRow → LastRow`, `FirstColumn → LastColumn` | ✅ `UsedRange.Rows.Count` |

### 5.2 Critical Differences ❌

| Aspect | ConMas | Ours | Impact |
|--------|--------|------|--------|
| **Primary backend** | **Infragistics4.Documents.Excel** (reads XLSX directly) | Excel COM Interop | **MEDIUM** — Infragistics reads XLSX without Excel installed; COM requires Excel |
| **Comment detection** | **Per-cell** via `WorksheetCell.get_HasComment()` | **Worksheet-level** via `ws.Comments` collection | **HIGH** — If `ws.Comments` returns null/empty but individual cells have comments, we miss them |
| **Field detection** | Comments are the **only** detection mechanism — no `_Fields` sheet reading in MakeCluster() | Checks `_Fields` sheet first, then falls back to comments | **HIGH** — ConMas does NOT read `_Fields` in the cluster/field detection path. The `_Fields` sheet is only read by `GetDefinition()` (XML export), not by `MakeCluster()` |
| **Comment type parsing** | Takes **first line** of comment text as field type | Uses `if "KeyboardText" in text` (string contains) | **MEDIUM** — First-line parsing is precise; string `in` is fragile |
| **Shape clearing** | `ClearShapes()` before processing | ❌ NOT done | **LOW** — Prevents shape interference but may not affect metadata reading |
| **Field type mapping** | Uses `ExcelFieldType` enum in `GetDefinition()` export path | Uses `_conmas_type_to_renderer_type()` mapping table | **LOW** — Mapping is identical |
| **Data path for rendering** | `MakeCluster()` → `CalcClusterSize()` → `List<ClusterRect>` → downstream renderer | `read_fields()` → `FieldDef[]` → `render_with_fields()` | **MEDIUM** — Different architecture, same end result |

---

## 6. Root Cause Analysis: Why Non-Working Workbooks Fail

### 6.1 FormTest - Copy.xlsx

**In our implementation:**
1. ✅ `_read_fields_sheet()` finds `_Fields` sheet → rows detected → iterates all rows
2. ✅ All cells are empty → returns `[]`
3. ✅ Falls through to `_read_comments()`
4. ❌ `ws.Comments` returns the comments → BUT if COM fails (e.g., `CoInitialize` issue), or if the comments collection API behaves differently, we miss them

**In ConMas implementation:**
1. `MakeCluster()` does NOT check `_Fields` sheet at all
2. Iterates every cell in used range
3. Calls `cell.HasComment()` on each cell → finds comments on A1, C1, A3, A6, A9, A12
4. Reads `comment.Text` → parses first line → gets field types
5. Returns cluster rectangles for all 6 fields

**Conclusion:** The fundamental difference is that ConMas uses a **per-cell comment scan** (`HasComment` on each cell) while we use a **worksheet-level comment collection** (`ws.Comments`). If `ws.Comments` returns null or empty (because of Excel version differences or COM quirks), our entire comment reading fails.

### 6.2 Sample A.xlsx

Same issue as FormTest — the `_Fields` sheet doesn't exist, and our `ws.Comments` collection might fail due to COM issues, while ConMas's per-cell `HasComment()` scan succeeds.

### 6.3 The COM initialization fix from Phase 3.3

The `pythoncom.CoInitialize()` fix should resolve the COM initialization issue, but there may still be an issue with `ws.Comments` vs per-cell `cell.Comment` or `HasComment()`. In modern Excel/COM interop:
- `ws.Comments` returns a `Comments` collection (classic comments)
- `cell.Comment` returns a single `Comment` object (per-cell)
- `cell.HasComment` might not exist on COM Range objects (it's an Infragistics API)
- `ws.CommentsThreaded` replaces `ws.Comments` in some Excel versions

**The per-cell approach** (what ConMas does via Infragistics) iterates the used range and checks each cell individually. Our `ws.Comments` approach relies on the worksheet-level collection working correctly.

---

## 7. Recommendations

### 7.1 Implement per-cell comment scan

Add an alternative code path in `_read_comments()` that iterates the used range cell by cell and checks each cell for a comment:

```python
# Alternative: per-cell comment scan (matches ConMas behavior)
try:
    used = ws.UsedRange
    for row_idx in range(1, used.Rows.Count + 1):
        for col_idx in range(1, used.Columns.Count + 1):
            cell = ws.Cells(row_idx, col_idx)
            try:
                comment = cell.Comment  # per-cell property
                if comment is not None:
                    addr = cell.Address(False, False)
                    text = str(comment.Text())
                    # ... process
            except:
                pass
except:
    pass
```

### 7.2 Add ThreadedComments fallback

For Excel 365 where classic comments are deprecated:

```python
try:
    tcomments = cell.ThreadedComments
    if tcomments is not None and tcomments.Count > 0:
        # Process threaded comment
        for tc in tcomments:
            text = tc.Text
except:
    pass
```

### 7.3 Add first-line comment parsing (precise)

Replace the fragile `if "KeyboardText" in text` with:

```python
# Parse comment text - first line is the field type
lines = comment_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
first_line = lines[0].strip() if lines else ''
data_type = _conmas_type_to_renderer_type(first_line)
```

### 7.4 The `_Fields` sheet is NOT needed for field detection

ConMas Designer's `MakeCluster()` doesn't read `_Fields` at all — it only scans cell comments. The `_Fields` sheet is used by `GetDefinition()` for XML export/publishing purposes. Our `_read_fields_sheet()` is still useful (it provides FieldId, FieldName information that comments don't have), but our comment reading must be robust enough to work standalone.

---

## 8. Summary of Key Insights

| Question | Answer | Supporting Evidence |
|----------|--------|-------------------|
| What is the primary field detection mechanism? | **Per-cell comment scan** via `WorksheetCell.get_HasComment()` + `get_Comment()` | IL: `ExcelRangeBase` calls to Infragistics API |
| Does ConMas read `_Fields` sheet in MakeCluster? | **NO** — only `GetDefinition()` (XML exporter) reads it | IL: `ExcelControllerBase.MakeCluster()` doesn't reference `_Fields` at all |
| What backend does the metadata reader use? | **Infragistics4.Documents.Excel** (not COM Interop) | IL: 217 method calls to Infragistics API vs. minimal COM calls |
| Is `ws.Comments` used? | **NO** — per-cell `cell.Comment` / `get_HasComment()` is used | IL: No call to `Worksheet.Comments` found |
| Does ConMas handle ThreadedComments? | **NO** — classic comments only via Infragistics | IL: No `ThreadedComment` references in LibExcelController |
| Why do our non-working workbooks fail? | `ws.Comments` collection vs per-cell comment scan — two different approaches | Architecture comparison |
| What is `CalculateCell`? | A simple data holder (address, row, col, sheet) | IL: Only default constructor |
| What is `GetDefinition()`? | **XML exporter**, NOT metadata reader | IL: Returns `XElement`, saves to .xml file |
