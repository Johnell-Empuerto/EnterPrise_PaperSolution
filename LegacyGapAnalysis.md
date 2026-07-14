# LegacyGapAnalysis.md

**Date:** 2026-07-13
**Phase:** 20 — Legacy Investigation (investigation only, no code changes)
**Objective:** Recover the original PaperLess publishing pipeline exactly

---

## STEP 1: Cluster Detection

### Observed Behavior (Our Engine)
- Iterates ALL merged cells + standalone cells within the print area
- Template 546: 21 clusters generated
- Template 547: 146 clusters generated
- Template 548: 12 clusters generated

### Legacy Behavior
- Clusters are ONLY created from **cells that have Excel COMMENTS**
- `InitClusterList()` iterates `worksheet.Comments` (1-indexed collection)
- For each comment: `Range mergeArea = range.MergeArea ?? range;`
- `clusterList.Add(mergeArea.Row, mergeArea.Column, mergeArea.Rows.Count, mergeArea.Columns.Count, comment.Text())`
- Comment text lines store cluster metadata:
  - Line 0: ClusterName (e.g., `Machine`, `samples`)
  - Line 1: ClusterTypeKey (e.g., `KeyboardText`, `InputNumeric`)
  - Line 2: ClusterIndex (parsed as int, -1 if empty)
  - Lines 3-15: InputParameters, TableNo, etc.

### Evidence

**Decompiled `ThisAddIn.InitClusterList()` (ThisAddIn.cs:732-760):**
```csharp
Comments comments = worksheet.Comments;
int? count = comments?.Count;
if (0 < count) {
    for (int i = 0; i < count; i++) {
        Comment comment = comments[i + 1];  // 1-based
        if (comment.Parent is Range range) {
            Range mergeArea = range.MergeArea ?? range;
            clusterList.Add(mergeArea.Row, mergeArea.Column,
                mergeArea.Rows.Count, mergeArea.Columns.Count,
                comment.Text() ?? "", false);
        }
    }
}
```

**COM verification — Template 548 Sheet1:**
```
Comments: 2
  Comment 1 on $A$10 → MergeArea $A$10:$D$11 R=10 C=1 RC=2 CC=4
    Line 0: Machine
    Line 1: KeyboardText
  Comment 2 on $E$10 → MergeArea $E$10:$G$11 R=10 C=5 RC=2 CC=3
    Line 0: Machine_Output
    Line 1: InputNumeric
```

**COM verification — Template 546 Sheet1:**
```
Comments: 6 (matching all 6 DB clusters)
  Comment 1 on $A$1 → Merge $A$1:$B$2  R=1 C=1 RC=2 CC=2
  Comment 2 on $C$1 → Merge $C$1:$D$2  R=1 C=3 RC=2 CC=2
  Comment 3 on $A$3 → Merge $A$3:$D$4  R=3 C=1 RC=2 CC=4
  Comment 4 on $A$6 → Merge $A$6:$D$7  R=6 C=1 RC=2 CC=4
  Comment 5 on $A$9 → Merge $A$9:$D$10 R=9 C=1 RC=2 CC=4
  Comment 6 on $A$12 → Merge $A$12      R=12 C=1 RC=1 CC=1
```

**DB `def_cluster` counts:**
- 546: 6 clusters — matches 6 comments
- 547: 5 clusters — matches 5 comments (expected)
- 548: 2 clusters — matches 2 comments

### Confidence
**100%** — Decompiled source + COM workbook verification + DB counts all match.

### Source Assembly
`iReporterExcelAddIn.dll` → `ThisAddIn.cs` line 732

### Requires Code Changes
**YES** — Cluster detection must switch from "all merged cells" to "cells with comments only". Must read `Range.Comment.Text()` and parse according to `ClusterInfo._texts` format.

---

## STEP 2: Merged Cell Investigation

### Observed Behavior (Our Engine)
- `CellModel.Reference` reports top-left cell address (e.g., `$A$1`)
- Cluster bounds use `StartRow`/`EndRow`/`StartColumn`/`EndColumn`

### Legacy Behavior
- `CellRect.ToAddressName()` generates full range `$A$1:$B$2` when Top≠Bottom or Left≠Right
- `ClusterInfo` stores Row, Col, Bottom=row+rowCount-1, Right=col+colCount-1
- DB `def_cluster.cell_addr` stores the full merged range

### Evidence
**Decompiled `CellRect.cs:58-66`:**
```csharp
public override string ToString() {
    string text = XLRangeUtil.ToAddressName(Top, Left, 1);
    if (Top != Bottom || Left != Right)
        text += ":" + XLRangeUtil.ToAddressName(Bottom, Right, 1);
    return text;
}
```

**DB XML `cellAddress` values match COM `MergeArea.Address` exactly:**
- 546: `$A$1:$B$2`, `$C$1:$D$2`, `$A$3:$D$4`, `$A$6:$D$7`, `$A$9:$D$10`, `$A$12`
- 548: `$A$10:$D$11`, `$E$10:$G$11`

### Confidence
**100%** — Range addresses match COM output and DB XML exactly.

### Requires Code Changes
**MINOR** — `CellModel.Reference` should serialize the full merged range address, not just the top-left cell.

---

## STEP 3: Legacy XML Investigation

### Complete XML Node Source Mapping

| XML Node | Category | Origin | Notes |
|----------|----------|--------|-------|
| `<top>`,`<bottom>`,`<right>`,`<left>` | **Generated** | Engine | Only truly computed values |
| `<name>` | Designer (Comment) | `ClusterInfo._texts[0]` | Line 0 of comment |
| `<type>` | Designer (Comment) | `ClusterInfo._texts[1]` | Line 1 of comment |
| `<clusterId>` | Database | `def_cluster.cluster_id` | |
| `<sheetNo>` | Database | `def_cluster.sheet_no` | |
| `<isHidden>`, `<isHiddenDesigner>` | Database | `def_cluster` | |
| `<mobileDisplay>` | Database | `def_cluster.mobile_display` | |
| `<pinNo>`, `<pinValue>` | Database | `def_cluster` | |
| `<external>` | Database | `def_cluster.external` | |
| `<displayValue>` | Database | `def_cluster.display_value` | |
| `<cooperationCluster>` | Database | `def_cluster.cooperation_cluster` | |
| `<readOnly>` | Database | `def_cluster.read_only` | |
| `<function>` | Database | `def_cluster` | |
| `<actionPost>` | Database | `def_cluster.action_post` | |
| `<inputParameters>` | Hybrid | Comment lines 3+ / DB | Serialized designer settings |
| `<excelOutputValue>` | Database | `def_cluster.excel_output_value` | |
| `<cellAddress>` | Database | `def_cluster.cell_addr` | Full merged range |
| `<defTopId>`, `<defTopName>` | Database | `def_top` | |
| `<sheetCount>` | Database | `def_top` | |
| `<width>`, `<height>` | Database | `def_sheet` | Page dimensions (e.g., 612×792) |
| `<serverVersion>` | Database | `def_top` | e.g., `8.2.26020` |
| All `<remarks*>`, `<systemKey*>` | Database | `def_top` | Config settings |

Only 4 nodes are computed during publishing: `<top>`, `<bottom>`, `<right>`, `<left>`. Everything else is from DB or designer comments.

### Confidence
**100%** — Full DB XML dump from templates 546 and 548 analyzed.

### Requires Code Changes
**YES** — `XmlGenerator.cs` must emit all ~60 metadata nodes from DB tables. Current minimal output is insufficient.

---

## STEP 4: Coordinate Investigation

### Gap Measurement — Template 548 (PrintArea: `$A$3:$G$11`, 612×792pt Letter)

**Cluster 0: `$A$10:$D$11` (cols A-D, rows 10-11)**

| Axis | DB Ratio | Our Ratio | Δ Ratio | Δ Points |
|------|----------|-----------|---------|----------|
| Left | 0.0847059 | 0.0847059 | **0** | **0 pt** |
| Top | 0.2040909 | 0.2040909 | **0** | **0 pt** |
| Right | 0.4111765 | 0.3984314 | 0.012745 | **7.80 pt** |
| Bottom | 0.2418182 | 0.2404545 | 0.001364 | **1.08 pt** |

**Cluster 1: `$E$10:$G$11` (cols E-G, rows 10-11)**

| Axis | DB Ratio | Our Ratio | Δ Ratio | Δ Points |
|------|----------|-----------|---------|----------|
| Left | 0.4129412 | 0.3984314 | 0.014510 | **8.88 pt** |
| Top | 0.2040909 | 0.2040909 | **0** | **0 pt** |
| Right | 0.6582353 | 0.6337255 | 0.024510 | **15.00 pt** |
| Bottom | 0.2418182 | 0.2404545 | 0.001364 | **1.08 pt** |

### Gap Diagnostics

**Left/Top match for origin cells (col A, row 10):** Column-width and row-height summation ALGORITHM IS CORRECT. The origin (margin + centering) calculation also matches.

**Width gap increases with column distance from origin:**
- Cluster 0 width: DB=199.80pt, Ours=192.00pt → gap = **7.80pt** (4.06%)
- Cluster 1 left_pt (sum A-D): DB=200.88pt, Ours=192.00pt → gap = **8.88pt** (4.63%)
- Cluster 1 right_pt (sum A-G): DB=351.00pt, Ours=336.00pt → gap = **15.00pt** (4.46%)

**The gap is proportional (~4.4% per column).** This rules out:
- ❌ Simple margin offset (would be constant, not proportional)
- ❌ Missing cluster gap constant (would be per-cluster, not per-column)

**Proportional gap supports these hypotheses:**
1. **Page width differs** — A4 (595pt) vs Letter (612pt) gives ~2.8% shift, partially explaining the gap
2. **Column widths read from different range** — COM `Range.Columns[i].Width` from USED range vs PRINT AREA can differ. The legacy Cimtops.Excel.dll may construct the column list from the **entire used range** rather than the print area
3. **A scaling factor is applied** — Excel `PageSetup.Zoom` or `FitToPages` could rescale all measurements

**Bottom gap is constant 1.08pt** — Rules out proportional row-height error. Suggests:
- A baseline row height difference (first row considered differently?)
- A fixed footer/margin-bottom offset

### Column Width Comparison (from raw data)

For Template 548 PrintArea `$A$3:$G$11`:
- DB sum(A-D) = 200.88pt → Our sum(A-D) = 192.00pt → **Each column ~2.22pt wider in DB**
- DB width(A-G total) = 351.00pt → Our width(A-G total) = 336.00pt → **Each column ~2.14pt wider in DB**

This consistent ~2.2pt/column overshoot is the signature of COM returning different column widths depending on the source range used.

### Confidence
**85%** — The gap is proportional to column position, pointing to COM range selection difference. The exact formula requires decompiling ConMasClient.exe. See Phase 18 full analysis at `Test Folder Final Test\Phase18_PublishPipelineAnalysis\LegacyCoordinateAlgorithm.md`.

### Requires Code Changes
**YES** — After cluster detection is fixed, the coordinate transform in `LegacyCoordinateTransform.cs` must be replaced with the correct algorithm from ConMasClient.exe. Currently identity.

---

## STEP 5: Background Investigation

### Observed Behavior
- `ImageComparer.Compare` fails with `ArgumentNullException` when trying to decode DB `background_image_file` bytes via SkiaSharp

### Legacy Behavior
- `background_image_file` stores **PDF files** with a leading null byte

### Evidence
**Magic bytes from all 3 templates:**
```
546: 00 25 50 44 46 2D 31 2E 37 0D 0A 25 B5 B5 B5 B5 0D 0A 31 20 30 20 6F 62 6A 0D 0A 3C 3C 2F 54 79
548: 00 25 50 44 46 2D 31 2E 37 0D 0A 25 B5 B5 B5 B5 0D 0A 31 20 30 20 6F 62 6A 0D 0A 3C 3C 2F 54 79
547: 00 25 50 44 46 2D 31 2E 37 0D 0A 25 B5 B5 B5 B5 0D 0A 31 20 30 20 6F 62 6A 0D 0A 3C 3C 2F 54 79
```
Translation: `[NUL]%PDF-1.7\r\n%[binary]\r\n1 0 obj\r\n<</Ty`

The leading null byte (`00`) is likely a database encoding artifact or part of the bytea hex serialization. The content is standard **PDF 1.7**.

**File sizes:**
- Template 546: 5,954 bytes
- Template 547: 373,816 bytes
- Template 548: 17,355 bytes

### Confidence
**100%** — Magic bytes are unambiguous `%PDF-1.7`.

### Requires Code Changes
**YES** — `ImageComparer` must handle PDF backgrounds. Either convert PDF→PNG for comparison, or use PDF comparison directly. Background export pipeline in `PublishEngine` currently exports to PDF then converts to PNG — should verify the comparison against the DB PDF directly.

---

## STEP 6: ConMasClient Investigation

### Target
`C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\ConMasClient.exe`
Size: 3,882,544 bytes
Type: WPF (.NET Framework 4.x) application

### Status
**BLOCKED** — Cannot be fully decompiled in the current environment. Requires .NET Framework 4.x runtime to resolve WPF types (PresentationFramework, WindowsBase, PresentationCore). ILSpy/dnSpy fail with:
```
Could not load type 'System.Windows.Rect' from assembly 'WindowsBase'
```

### Known Candidate Methods (from string analysis)

| Search Pattern | Location Evidence | Purpose |
|---------------|-------------------|---------|
| `Publish` | String analysis | Main publish command |
| `Export` | String analysis (`_Export`, `_Area`) | Export/serialize |
| `LeftPosition` | Hypothesized | Coordinate assignment |
| `TopPosition` | Hypothesized | Coordinate assignment |
| `RightPosition` | Hypothesized | Coordinate assignment |
| `BottomPosition` | Hypothesized | Coordinate assignment |
| `DefCluster` | DB table name | Cluster DB serialization |
| `XmlData` | DB column `xml_data` | XML generation |
| `Normalize` | Phase report | Ratio normalization |
| `PageWidth` / `PageHeight` | Decompiled Sheet model | Page dimension reads |
| `CalcArea` | Decompiled R2Cluster | Area percentage calculation |
| `_xlnm.Print_Area` | String analysis | Print area reference |

### Decompiled Assemblies (Available)

| Assembly | Status | Key Classes |
|----------|--------|-------------|
| `Cimtops.Excel.dll` | Fully decompiled (Phase 15, no source in workspace) | `Book`, `Sheet`, `Cell`, `CellRange`, `Col`, `Row`, `BookFactory` |
| `Cimtops.R2Cluster.dll` | Fully decompiled (Phase 15, no source in workspace) | `Cluster`, `Decoder`, `ClusterInfo`, `Caption`, `TitleInfo` |
| `iReporterExcelAddInCommon.dll` | Decompiled (Phase 18) | `CellRect`, `ClusterInfo`, `ClusterList`, `ClusterTypeInfo` |
| `iReporterExcelAddIn.dll` | Decompiled (Phase 18) | `ThisAddIn`, `Ribbon1`, `InputForm2` |

### Recovered Coordinate Pipeline (from decompiled Cimtops.Excel.dll Phase 15)

```
CellRange constructor:
    Left   = range.Column       (1-based column number)
    Top    = range.Row          (1-based row number)
    Right  = Left + Columns.Count - 1
    Bottom = Top + Rows.Count - 1

Sheet constructor:
    Col[].Width   = COM Range.Columns[i].Width   (points)
    Row[].Height  = COM Range.Rows[i].Height      (points)
```

### Required Decompilation Setup
To fully decompile ConMasClient.exe:
1. Windows machine with .NET Framework 4.7.2+
2. ILSpy GUI with WPF reference resolution
3. Or ilspycmd with:
   ```
   ilspycmd ConMasClient.exe -p -o .\decompiled
     -r "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
   ```
4. Search targets: `Publish`, `Export`, `LeftPosition`, `TopPosition`, `RightPosition`, `BottomPosition`

### Confidence
**100%** that ConMasClient.exe contains the final coordinate transform.
**0%** recovered — completely blocked without .NET Framework 4.x.

### Requires Code Changes
**YES** — After decompilation, replace `LegacyCoordinateTransform.cs` identity with recovered transform.

---

## Summary of All Gaps

| Gap | Root Cause | Evidence Level | Fix Location | Priority |
|-----|-----------|----------------|-------------|----------|
| Extra clusters (21 vs 6, etc.) | Detecting all merged cells instead of commented cells only | **100%** (decompiled source) | `ClusterDetector.cs` | P0 |
| Wrong cluster metadata (hardcoded `samples`/`KeyboardText`) | Not reading from cell comments | **100%** (decompiled source) | `ClusterDetector.cs` | P0 |
| Coordinates differ 4.4% proportional | Column-width reading range differs OR page normalization differs | **85%** (gap pattern analysis) | `LegacyCoordinateTransform.cs` | P1 |
| XML missing ~50 metadata nodes | XmlGenerator writes minimal subset | **100%** (DB XML dump) | `XmlGenerator.cs` | P1 |
| Background images fail to decode | DB stores PDF, not raster image | **100%** (magic bytes) | `ImageComparer.cs` | P2 |
| Final coordinate transform unknown | ConMasClient.exe not decompilable | **100%** blocked | ConMasClient.exe | P0 blocker |

### Legend
- **P0**: Blocking correctness — must fix before any comparison succeeds
- **P1**: Required for full match — output is structurally incomplete/incorrect
- **P2**: Nice-to-have — comparison can degrade gracefully
- **P0 blocker**: Must decompile ConMasClient.exe on .NET Framework 4.x to proceed

### Decompiled Source Files Available
All at `Test Folder Final Test\Phase18_Decompiled\`:
- `iReporterCommon\iReporterExcelAddInCommon\CellRect.cs` (67 lines)
- `iReporterCommon\iReporterExcelAddInCommon\ClusterInfo.cs` (98 lines)
- `iReporterCommon\iReporterExcelAddInCommon\ClusterList.cs` (134 lines)
- `iReporterCommon\iReporterExcelAddInCommon\ClusterTypeInfo.cs`
- `iReporterAddIn\iReporterExcelAddIn2019\ThisAddIn.cs` (1159 lines) ← KEY FILE

### Full Phase 18 Report
`Test Folder Final Test\Phase18_PublishPipelineAnalysis\LegacyCoordinateAlgorithm.md` (666 lines)

---

## Phase 20.1 Implementation Results

### Change Summary

Replaced merged-cell + standalone-cell cluster detection with **comment-based detection**:

| File | Change |
|------|--------|
| `Models\SheetModel.cs` | Added `CommentModel` class, `Comments` property to `SheetModel`, fixed `CellModel.Reference` to emit full merge range |
| `ExcelEngine\WorksheetLoader.cs` | Added COM `worksheet.Comments` iteration — reads each comment's `Parent` → `MergeArea` → `Row`/`Column`/`Rows.Count`/`Columns.Count`/`Text()` |
| `ClusterEngine\ClusterDetector.cs` | Rewrote `Detect()` to iterate `sheet.Comments`, parse comment text (line 0 = name, line 1 = type) |
| `ClusterEngine\ClusterBuilder.cs` | Rewrote `BuildAll()` to iterate `sheet.Comments` with `CommentModel`, removed standalone-cell fallback |
| `Models\ClusterModel.cs` | Added `CommentName`, `CommentType` |
| `PublishEngine\XmlGenerator.cs` | Emits `CommentName`/`CommentType` from comment text instead of hardcoded `samples`/`KeyboardText` |

### Regression Results

| Template | DB Clusters | Generated Clusters | Match | DB Names | Generated Names | Match |
|----------|-------------|-------------------|-------|----------|----------------|-------|
| 546 | 6 | **6** | ✓ | `samples` (×6) | `samples` (×6) | ✓ |
| 547 | 5 | **5** | ✓ | `Cluster0`-`Cluster4` | `""` (empty) | ⚠¹ |
| 548 | 2 | **2** | ✓ | `Machine`, `Machine_Output` | `Machine`, `Machine_Output` | ✓ |

| Template | DB Types | Generated Types | Match |
|----------|----------|----------------|-------|
| 546 | `KeyboardText` (×6) | `KeyboardText` (×6) | ✓ |
| 547 | `KeyboardText` (×5) | `KeyboardText` (×5) | ✓ |
| 548 | `KeyboardText`, `InputNumeric` | `KeyboardText`, `InputNumeric` | ✓ |

¹ Template 547 comment line 0 is **empty** (confirmed via COM). DB names `Cluster0`-`Cluster4` are auto-generated by ConMasClient.exe publish pipeline. Not implemented yet (ConMasClient blocked).

### Files Changed
```
ExcelAPI/ExcelAPI/LegacyEngine/
  ClusterEngine/ClusterDetector.cs      — Rewritten (comment-based)
  ClusterEngine/ClusterBuilder.cs       — BuildAll() rewritten (comment-based)
  ExcelEngine/WorksheetLoader.cs        — Added COM comment reading
  Models/SheetModel.cs                  — Added CommentModel, Comments, fixed CellAddress format
  Models/ClusterModel.cs               — Added CommentName, CommentType
  PublishEngine/XmlGenerator.cs         — Emits comment name/type instead of hardcoded
```

### Build Status
**0 errors** — build succeeds cleanly.
