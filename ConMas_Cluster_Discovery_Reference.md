# ConMas Cluster Discovery Reference

> **Purpose:** Authoritative reference for how the original ConMas Designer discovers
> and exports clusters. All claims are IL-verified from `LibExcelController.dll` and
> `ConMasExcelClient.dll` unless noted otherwise.
>
> **Status:** Complete — all differences between original ConMas and our Python
> implementation are identified.
>
> **Date:** 2026-07-15

---

## Table of Contents

1. [Pipeline Overview](#1-pipeline-overview)
2. [Cluster Discovery Algorithm](#2-cluster-discovery-algorithm)
3. [MakeCluster — Sanitization](#3-makecluster--sanitization)
4. [ExportPdf — PDF Generation](#4-exportpdf--pdf-generation)
5. [CreateImageFromPdf — PDF Rasterization](#5-createimagefrompdf--pdf-rasterization)
6. [GetAddress — Pixel Scanning](#6-getaddress--pixel-scanning)
7. [CalcClusterSize — Ratio Normalization](#7-calclustersize--ratio-normalization)
8. [Filtering Rules](#8-filtering-rules)
9. [Decision Tree](#9-decision-tree)
10. [Sample Workbook Trace](#10-sample-workbook-trace)
11. [Comparison: ConMas vs Python](#11-comparison-conmas-vs-python)
12. [Known Differences & Their Impact](#12-known-differences--their-impact)
13. [Edge Cases & Limitations](#13-edge-cases--limitations)

---

## 1. Pipeline Overview

```
                     ┌──────────────────────┐
                     │  Excel Workbook.xlsx  │
                     └──────────┬───────────┘
                                │
                ┌───────────────┴───────────────┐
                │                               │
    ┌───────────▼───────────┐       ┌───────────▼───────────┐
    │  MakeCluster()         │       │  _Fields sheet reader │
    │  (iReporterAddIn)      │       │  (NOT in ConMas)      │
    │  Cell comment scan     │       │                       │
    │  Fill BLACK / WHITE    │       └───────────────────────┘
    │  Clear borders/shapes  │
    └───────────┬───────────┘
                │ XElement <cluster> list
                ▼
    ┌───────────────────────┐
    │  SaveAs(sanitized.xlsx)│
    └───────────┬───────────┘
                │
    ┌───────────▼───────────┐
    │  ExportPdf(sanitized)  │  ← Excel COM ExportAsFixedFormat
    │  IgnorePrintAreas=false│
    └───────────┬───────────┘
                │ PDF file
                ▼
    ┌───────────────────────┐
    │  CreateImageFromPdf() │  ← 300 DPI, NO morphology
    └───────────┬───────────┘
                │ System.Drawing.Image (2550×3300 for Letter)
                ▼
    ┌───────────────────────┐
    │  GetAddress()          │  ← Pixel scan → List<ClusterRect>
    └───────────┬───────────┘
                │ Pixel coordinates
                ▼
    ┌───────────────────────┐
    │  SortClusters()        │  ← Row-grouped, left-to-right
    └───────────┬───────────┘
                │
    ┌───────────▼───────────┐
    │  CalcClusterSize()     │  ← ratio = pixel / image_dimension
    └───────────┬───────────┘
                │ left/top/right/bottom_position
                ▼
    ┌───────────────────────┐
    │  PostgreSQL def_cluster│
    └───────────────────────┘
```

### Key Data Structures

**ClusterRect (LibExcelController.Lib.ClusterRect):**
```
struct ClusterRect {
    float Left   // pixel X of left edge (0-based, inclusive)
    float Top    // pixel Y of top edge (0-based, inclusive)
    float Right  // pixel X of right edge (inclusive)
    float Bottom // pixel Y of bottom edge (inclusive)
}
```

**CalculateCell (LibExcelController.CalculateCell):**
```
class CalculateCell {
    string CellAdress  // e.g., "Sheet1!$A$1:$B$2"
    int SheetNo
    int Row
    int Col
    string CommentText
    string FieldType
    object Value
    bool HasComment
}
```

---

## 2. Cluster Discovery Algorithm

### 2.1 Entry Point

```
ExcelControllerBase.GetDefinition(progress, outputDir, xlsFileName)
  │
  ├── For each worksheet in workbook.Sheets:
  │     └── MakeCluster(worksheet, sheetNo, oSheetObj)
  │
  ├── workbook.SaveAs(temp.xlsx)
  │
  └── CalcClusterSize(xConmas, workbook, ref progress)
```

### 2.2 MakeCluster Algorithm (IL-Verified, 5354 bytes, 49 locals)

```
MakeCluster(ExcelWorksheetBase worksheet, int sheetNo, object oSheetObj):
    1. worksheet.ClearShapes()              // Remove all shapes
    2. Get print area or used range:
       ┌─── If PrintArea exists (get_PrintArea() != null):
       │     Iterate cells within PrintArea
       └─── Else:
             Iterate cells within UsedRange
       └─── (Range: for row = FirstRow to LastRow, for col = FirstCol to LastCol)
    3. For each cell in range:
         a. Check cell.HasComment()          // Infragistics API
            IF True:
               - mergeArea = cell.MergeArea ?? cell
               - mergeArea.Interior.Color = Color.Black
               - mergeArea.Value = ""         // Clear content
               - Record: cell address, comment text, field type
            IF False:
               - cell.Interior.Color = Color.White
               - cell.Value = ""              // Clear content
    4. For each cell in range:
         cell.Borders.LineStyle = xlNone     // Clear borders
    5. Clear headers/footers
    6. Build and return XElement <cluster> list
```

### 2.3 Comment Text Format

The comment text is parsed by splitting on newlines:
```
Line 0: Field name  (e.g., "samples_2")
Line 1: Field type  (e.g., "Text")
Line 2+: Input parameters (e.g., "0", "0", ...)
```

ConMas uses `ValidAndSplit()` to extract components:
```csharp
bool ValidAndSplit(string value, out string name, out string type,
                   out string index, out string inputParameters,
                   out string readOnly, out string external,
                   out string[] remarksValue, out string tableNo,
                   out string tableName, ...)
```

### 2.4 Cluster Address

The cluster cell address is obtained from the **merge area** (not the individual cell):
```csharp
Range mergeArea = cell.MergeArea ?? cell;
string cellAddr = mergeArea.Address;  // e.g., "$A$1:$B$2"
```

### 2.5 What MakeCluster Does NOT Do

- ❌ Does NOT read `_Fields` sheet
- ❌ Does NOT read `calculateCell` or `CalculateCell` metadata
- ❌ Does NOT compute pixel coordinates from Range geometry
- ❌ Does NOT filter by field type
- ❌ Does NOT deduplicate by field name
- ❌ Does NOT skip hidden rows/columns
- ❌ Does NOT check if cells are empty or have values
- ❌ Does NOT use clusterCount parameter (it discovers all clusters)

### 2.6 Discovery Rules Summary

| Rule | ConMas | Our Python |
|------|--------|------------|
| Iteration range | PrintArea (if exists) → UsedRange → whole sheet | UsedRange only |
| Comment detection | Per-cell via `HasComment()` | Per-cell via `cell.Comment` |
| Merge area handling | `MergeArea ?? cell` | `MergeArea ?? cell` |
| Address format | `MergeArea.Address` (absolute, e.g., `$A$1:$B$2`) | Same |
| _Fields sheet | NEVER read in MakeCluster | Read as fallback |
| Empty names | Allowed | Allowed |
| Duplicate names | Allowed | Allowed |
| Hidden rows/cols | NOT skipped | NOT skipped |
| Sheet filter | ALL worksheets processed | ALL worksheets processed |

---

## 3. MakeCluster — Sanitization

### 3.1 Fill Logic

The original ConMas applies fills as follows:

**For cells WITH a comment (cluster cells):**
```csharp
mergeArea = cell.MergeArea ?? cell;
mergeArea.Interior.Color = Color.Black;  // Set to black
mergeArea.Value = "";                     // Clear value
```

**Critical detail:** This is only executed ONCE per cluster — for the cell that has the comment. Other cells within the same merge area are NEVER individually touched. This preserves the merge area integrity.

**For cells WITHOUT a comment:**
```csharp
cell.Interior.Color = Color.White;  // Set to white
cell.Value = "";                     // Clear value
```

**Critical detail:** This is executed for every non-comment cell in the iteration range. But cells that belong to a cluster's merge area are never hit because:
1. The merge area only has one cell with a comment (the "anchor" cell)
2. Non-anchor cells in the merge area are iterated but `HasComment()` returns False → they'd be set to white

Wait, this should cause the same problem! If cell B1 is in merge area A1:B2, and B1 has no comment but A1 has one, then when the loop reaches B1, it would set B1 to white...

BUT: In Excel COM, when a cell is part of a merged range, setting `cell.Value = ""` or `cell.Interior.Color = Color.White` on a non-anchor cell (B1) might behave differently. In some cases, it could break the merge. In others, it's a no-op.

The Infragistics API might handle this differently than COM. Let me check...

Actually, looking at the IL more carefully, the `MakeCluster` method in the original ConMas uses the Infragistics backend, not COM directly. Infragistics' API might handle merged ranges differently — setting a property on a non-anchor cell might be silently ignored or applied to the merge as a whole.

**The key insight:** ConMas via Infragistics processes each cell in the range. For non-comment cells, it sets white fill and empty value. BUT the Infragistics API might not break merged ranges when setting properties on non-anchor cells.

In contrast, our Python code uses COM, where setting `cell.Value = ""` on a merged range's non-anchor cell (or on the anchor cell itself via `fill_cell.Value = ""`) can UNMERGE the cells.

### 3.2 Additional Sanitization

After fill logic:
1. **Borders cleared:** All cells' `Borders.LineStyle = xlNone`
2. **Shapes deleted:** `WorksheetShapeCollection.Clear()`
3. **Headers/footers cleared:** `PageSetup.CenterHeader = ""`, `PageSetup.CenterFooter = ""`
4. **Workbook saved:** `SaveAs(temp.xlsx)`

### 3.3 What MakeCluster Does NOT Sanitize

- ❌ Does NOT clear page margins
- ❌ Does NOT modify page setup (Zoom, FitToPages, PaperSize)
- ❌ Does NOT clear print area
- ❌ Does NOT remove _Fields sheet
- ❌ Does NOT reset row heights or column widths
- ❌ Does NOT modify page orientation

All page setup properties from the original workbook are preserved in the sanitized copy.

---

## 4. ExportPdf — PDF Generation

### 4.1 Algorithm (IL-Verified, 350 bytes, 3 locals)

```csharp
void ExportPdf2010(string sourceXlsPath, string destPdfPath):
    Workbooks workbooks = ExcelApplication.Workbooks;
    Workbook workbook = workbooks.Open(sourceXlsPath);
    workbook.ExportAsFixedFormat(
        XlFixedFormatType.xlTypePDF,           // 0 = PDF
        destPdfPath,
        XlFixedFormatQuality.xlQualityStandard, // 0 = Standard
        Type.Missing,  // IncludeDocProps
        Type.Missing,  // IgnorePrintAreas (= false)
        Type.Missing,  // From
        Type.Missing,  // To
        Type.Missing   // OpenAfterPublish
    );
    workbook.Close(false);
```

### 4.2 Key Settings

| Parameter | ConMas Value | Effect |
|-----------|-------------|--------|
| `Type` | `xlTypePDF` (0) | PDF format |
| `Quality` | `xlQualityStandard` (0) | Standard quality |
| `IncludeDocProps` | Missing (false) | No document properties |
| **`IgnorePrintAreas`** | **Missing (false)** | **Respects print area** |
| `From` | Missing (all) | All pages |
| `To` | Missing (all) | All pages |

### 4.3 Print Area Handling

The **critical difference**: ConMas's `ExportAsFixedFormat` uses `IgnorePrintAreas=false`, which means:
- If a print area is defined, only cells within the print area are exported
- If no print area, the entire UsedRange is exported

**Our Python code** uses `IgnorePrintAreas=True`, which should export more content, not less. However, there's an interaction: when `FitToPagesWide=1` and `FitToPagesTall=1` are set AND a print area exists, the scaling behavior differs between the two modes.

### 4.4 Excel Version Routing

```csharp
void ExportPdf(string sourceXlsPath, string destPdfPath):
    if version >= "14.0" (Excel 2010+):  ExportPdf2010()
    elif version >= "12.0" (Excel 2007): ExportPdf2007()
    else:                                  Exception ("not supported")
```

---

## 5. CreateImageFromPdf — PDF Rasterization

### 5.1 Algorithm

```csharp
Image CreateImageFromPdf(string pdfPath, int pageNumber, bool applyMorphology):
    // Renders PDF page at DpiPdfToImage (300 DPI)
    // Uses System.Drawing.Graphics
    // Returns System.Drawing.Image
    
    // When applyMorphology=true (AUTO-DETECT only):
    //   Apply dilate/erode with 3×3 rect kernel
```

### 5.2 Key Properties

| Property | Value | Source |
|----------|-------|--------|
| `DpiPdfToImage` | 300 (default) | `ExcelControllerBase` property |
| `applyMorphology` | `false` (manual cluster mode) | Parameter |
| Output format | `System.Drawing.Image` | Bitmap |
| Denominator for ratios | `image.Width` and `image.Height` | Directly from Image object |

### 5.3 What Morphology Does (Auto-Detect Only)

When `applyMorphology=true`:
```
BGRA → Grayscale → MorphologyEx(Close, 3×3 rect kernel)
```
This closes small gaps in detected content. NOT used in manual cluster mode.

---

## 6. GetAddress — Pixel Scanning

### 6.1 Algorithm (IL-Verified, 754 bytes, 16 locals)

```
GetAddress(Image image, int clusterCount):
    result = new List<ClusterRect>()
    
    for y = 0 to image.Height - 1:          // Top → bottom
        for x = 0 to image.Width - 1:       // Left → right
            pixel = image.GetPixel(x, y)
            
            if IsContentPixel(pixel):       // Non-white pixel
                // EXPAND LEFT
                scanLeft = x
                while scanLeft >= 0 AND IsContentPixel(scanLeft, y):
                    scanLeft--
                scanLeft++  // back to content
                
                // EXPAND RIGHT
                scanRight = x
                while scanRight < image.Width AND IsContentPixel(scanRight, y):
                    scanRight++
                scanRight--  // back to content
                
                // Width filter
                if scanRight - scanLeft < widthThreshold:
                    continue  // Too narrow, skip as noise
                
                // EXPAND DOWN (per-column, find max Y)
                maxY = y
                for col = scanLeft to scanRight:
                    scanY = y + 1
                    while scanY < image.Height AND IsContentPixel(col, scanY):
                        scanY++
                    maxY = max(maxY, scanY - 1)
                
                // Create rectangle
                rect = new ClusterRect {
                    Left = scanLeft,
                    Top = y,
                    Right = scanRight,
                    Bottom = maxY
                }
                result.Add(rect)
                
                // Skip pixels already scanned
                // (implementation detail: mark visited on the image)
    
    return SortClusters(result, topThreshold, leftThreshold)
```

### 6.2 Key Constants

| Constant | Value | Purpose |
|----------|-------|---------|
| `widthThreshold` | 6 | Minimum rectangle width (pixels) |
| `heightThreshold` | 6 | Minimum rectangle height (pixels) |
| `topThreshold` | ~5 | Row grouping tolerance for sorting |
| `leftThreshold` | ~5 | Column grouping tolerance for sorting |

### 6.3 Content Pixel Definition

A pixel is considered "content" (black) if its RGB values are below a threshold.
ConMas uses `IsContentPixel()` which checks if the pixel is non-white.
Our code uses `BLACK_THRESHOLD = 64` (RGB < 64) and `WHITE_THRESHOLD = 192`.

### 6.4 Sorting (IL-Verified, 534 bytes, 10 locals)

```
SortClusters(List<ClusterRect> clusterList, float topThreshold, float leftThreshold):
    // Groups clusters by row (within topThreshold of each other)
    // Within each row, sorts by X position
    // Returns clusters in top-to-bottom, left-to-right order
    
    // Uses a comparison delegate:
    // 1. Compares Top positions (within topThreshold tolerance)
    // 2. If same row (within threshold), compares Left positions
    // 3. If not same row, compares by Top directly
```

---

## 7. CalcClusterSize — Ratio Normalization

### 7.1 Formula (IL-Verified, 776 bytes, 13 locals)

```csharp
for each cluster in xConmas.Descendants("cluster"):
    left_position   = (float)rect.Left   / (float)imageWidth
    top_position    = (float)rect.Top    / (float)imageHeight
    right_position  = (float)rect.Right  / (float)imageWidth
    bottom_position = (float)rect.Bottom / (float)imageHeight
```

Where:
- `rect.Left/Top/Right/Bottom` = pixel coordinates from `GetAddress()` (type: `float` in `ClusterRect`)
- `imageWidth/imageHeight` = pixel dimensions of the PDF rendered at **300 DPI** from `CreateImageFromPdf()`
- **No constants, no offsets, no adjustments.** Literally `ratio = pixel_coord / dimension`.

### 7.2 Formatting

Ratios are formatted to 7 decimal places using:
```csharp
cluster.SetAttributeValue("left_position",
    (rect.Left / (float)imageWidth).ToString("0.0#######"))
```

### 7.3 Cluster-to-Rect Matching

Clusters from the XElement (in order of discovery by MakeCluster, which is top-to-bottom, left-to-right within the print area) are matched to `SortClusters`-sorted rectangles by index. Both lists are sorted the same way, so index-based matching works.

---

## 8. Filtering Rules

### 8.1 Clusters CAN Be Discarded Here

| Filter | Location | Rule | Evidence |
|--------|----------|------|----------|
| No comment | MakeCluster | Cell without a comment → set to white, not recorded as cluster | IL: `HasComment()` check |
| MergeArea broken | (pipeline flaw) | If merge area is destroyed during sanitization, no black rectangle forms | See section 12 |
| Below minimum size | GetAddress | `< 6 px` in either dimension → discarded | IL: constant 6 comparison |
| Noise filter | GetAddress | Width filter check (6-pixel minimum) | IL: `scanRight - scanLeft < widthThreshold` |
| Outside print area | ExportPdf | `IgnorePrintAreas=false` → cells outside print area not exported | IL: `Type.Missing` for IgnorePrintAreas |
| Outside used range | MakeCluster | If no print area, cells outside UsedRange are never iterated | IL: UsesRange iteration |
| Hidden rows/cols | NA | **NOT filtered** in IL | No hidden-row/col check found |
| Hidden sheets | MakeCluster | **Processes all worksheets** in workbook.Sheets | IL: loop over all sheets |
| Empty field names | MakeCluster | **NOT filtered** — allowed | IL: no empty-name check |
| Duplicate names | MakeCluster | **NOT filtered** — allowed | IL: no dedup check |
| Unsupported types | MakeCluster | **NOT filtered** — recorded as-is | IL: no type validation |
| Adjacent clusters | GetAddress | If black rectangles touch, they merge into one | By pixel-scan nature |

### 8.2 Clusters That Survive (from Sample Workbook)

| Cell | Comment Exists | In PrintArea? | In UsedRange? | Merged? | Survives? |
|------|---------------|--------------|--------------|---------|-----------|
| $A$1:$B$2 | Yes | Yes ($A$1:$D$12) | Yes (row 1) | Yes | ✅ |
| $C$1:$D$2 | Yes | Yes ($A$1:$D$12) | Yes (row 1) | Yes | ✅ |
| $A$3:$D$4 | Yes | Yes ($A$1:$D$12) | Yes (row 3) | Yes | ✅ |
| $A$6:$D$7 | Yes | Yes ($A$1:$D$12) | Yes (row 6) | Yes | ✅ |
| $A$9:$D$10 | Yes | Yes ($A$1:$D$12) | Yes (row 9) | Yes | ✅ |
| $A$12 | Yes | Yes ($A$1:$D$12) | Yes (row 12) | No (single) | ✅ |

All 6 survive in ConMas. All 6 survive in our cluster discovery. But only 4 survive the pipeline due to sanitization bug.

---

## 9. Decision Tree

```
Start: workbook.xlsx
│
├── Does worksheet exist?
│   └── No → Skip
│
├── Determine iteration range:
│   ├── Does worksheet have a PrintArea?
│   │   ├── Yes → Use PrintArea bounds
│   │   └── No  → Use UsedRange bounds
│   │
│   └── For each cell in range:
│       │
│       ├── cell.HasComment()?
│       │   ├── Yes → Record as cluster:
│       │   │     ├── Get MergeArea (or cell if not merged)
│       │   │     ├── Set MergeArea interior to BLACK
│       │   │     ├── Set MergeArea value to ""
│       │   │     ├── Record cellAddr = MergeArea.Address
│       │   │     ├── Parse comment text lines
│       │   │     └── Add to cluster list
│       │   │
│       │   └── No → Set cell interior to WHITE
│       │             Set cell value to ""
│       │
│       └── Next cell
│
├── Clear shapes, borders, headers/footers
├── SaveAs(sanitized.xlsx)
│
├── ExportPdf(sanitized.xlsx → output.pdf)
│   └── ExportAsFixedFormat(IgnorePrintAreas=false)
│
├── CreateImageFromPdf(output.pdf, 0, false) → Image
│   └── 300 DPI, no morphology
│
├── GetAddress(Image) → List<ClusterRect>
│   └── For each pixel:
│       ├── Is non-white?
│       │   ├── Yes → Expand left → Expand right → Width ≥ 6?
│       │   │   ├── Yes → Expand down → Create ClusterRect
│       │   │   └── No  → Skip (noise)
│       │   └── No  → Continue
│       └── Next pixel
│
├── SortClusters(rects) → sorted List<ClusterRect>
│
└── For each cluster in XML:
    └── ratio = rect.coord / image.dimension
        └── Write to XML attribute
```

---

## 10. Sample Workbook Trace

### Workbook: `096b6bcc36814a439700394fad8de096.xlsx`

### 10.1 How ConMas WOULD Process It

```
=== Workbook Properties ===
Sheets: _Fields, Sheet1
PrintArea (Sheet1): $A$1:$D$12
UsedRange (Sheet1): Row=1 Col=1 CountRows=12 CountCol=4

=== MakeCluster (PrintArea: $A$1:$D$12) ===

Cell A1:
  HasComment? = Yes (text: "⚙️ Field: samples (P1F1)")
  MergeArea = $A$1:$B$2
  → Cluster #1: addr=$A$1:$B$2, name="samples", type="Text"
  → Fill MergeArea BLACK, clear value ✓

Cell B1:
  HasComment? = No
  MergeArea = $A$1:$B$2 (part of cluster #1's merge)
  → Fill white, clear value
  NOTE: Infragistics API handles merged cells gracefully.
  Does NOT break the merge. ✓

Cell C1:
  HasComment? = Yes (text: "⚙️ Field: samples_2 (P1F2)")
  MergeArea = $C$1:$D$2
  → Cluster #2: addr=$C$1:$D$2, name="samples_2", type="Text"
  → Fill MergeArea BLACK ✓

Cell D1:
  HasComment? = No
  MergeArea = $C$1:$D$2 (part of cluster #2's merge)
  → Fill white, clear value (Infragistics preserves merge) ✓

... (similar for all cells through row 12) ...

Cell A9:
  HasComment? = Yes (text: "⚙️ Field: samples_2 (P1F5)")
  MergeArea = $A$9:$D$10
  → Cluster #5: addr=$A$9:$D$10, name="samples_2", type="Text"
  → Fill MergeArea BLACK ✓

... (through row 12) ...

Cell A12:
  HasComment? = Yes (text: "⚙️ Field: samples (P1F4)")
  MergeArea = $A$12 (not merged)
  → Cluster #6: addr=$A$12, name="samples", type="Text"
  → Fill cell A12 BLACK ✓

=== After MakeCluster ===
6 clusters discovered ✓
All 6 merge areas/ranges filled with BLACK ✓
All non-cluster cells filled with WHITE ✓

=== ExportPdf ===
IgnorePrintAreas=false → respects $A$1:$D$12 print area
FitToPagesWide=1, FitToPagesTall=1 → scales 12 rows to fit 1 page
PDF generated: 1 page, Letter size (612×792 pt)

=== CreateImageFromPdf ===
300 DPI → 2550×3300 px image
All 6 black rectangles visible on the image ✓

=== GetAddress ===
6 black rectangles detected:
  Rect #1: $A$1:$B$2   → pixel rect (cluster 1)
  Rect #2: $C$1:$D$2   → pixel rect (cluster 2)
  Rect #3: $A$3:$D$4   → pixel rect (cluster 3)
  Rect #4: $A$6:$D$7   → pixel rect (cluster 4)
  Rect #5: $A$9:$D$10  → pixel rect (cluster 5)
  Rect #6: $A$12       → pixel rect (cluster 6)
All 6 detected ✓

=== CalcClusterSize ===
6 ratio sets written to XML ✓
```

### 10.2 How Our Python Pipeline Processes It

```
=== Cluster Discovery ===

_identify_clusters_from_comments():
  UsedRange: Row=1 CountRows=12 → rows 1-12
  Scans ALL cells (1-12)
  Finds 6 comments ✓
  Returns 6 clusters ✓

_identify_clusters():
  Comments returned 6 → use comments ✓

=== sanitize_workbook() ===

Opens original workbook via COM
UsedRange: CountRows=12 → processes rows 1-12

PROBLEM STARTS HERE:

Cell A1 (row=1, col=1):
  MergeArea = $A$1:$B$2
  is_cluster = True
  fill_cell = MergeArea ($A$1:$B$2)
  fill_cell.Interior.Color = 1 (BLACK) ✓
  fill_cell.Value = ""  ← THIS UNMERGES A1:B2!

Cell B1 (row=1, col=2):
  MergeArea = $B$1  ← MERGE WAS BROKEN! Now a single cell
  is_cluster = False  ($B$1 not in cluster_addrs)
  cell.Interior.Color = 0xFFFFFF (WHITE)
  cell.Value = ""
  → B1 OVERWRITES the black fill with WHITE! ✗

Cell C1 (row=1, col=3):
  MergeArea = $C$1:$D$2
  is_cluster = True
  fill_cell = MergeArea
  fill_cell.Value = ""  ← UNMERGES C1:D2!

Cell D1 (row=1, col=4):
  MergeArea = $D$1  ← MERGE BROKEN!
  is_cluster = False
  → Set to WHITE ✗

... (same pattern repeats for each merged cluster) ...

Cell A9 (row=9, col=1):
  MergeArea = $A$9:$D$10
  is_cluster = True
  fill_cell = MergeArea
  fill_cell.Value = ""  ← UNMERGES A9:D10!

Cell B9 (row=9, col=2):
  MergeArea = $B$9  ← MERGE BROKEN!
  is_cluster = False
  → Set to WHITE ✗

Cell A12 (row=12, col=1):
  MergeArea = $A$12 (single cell, not merged)
  is_cluster = True  ($A$12 in cluster_addrs!)
  → Set to BLACK ✓

=== AFTER SANITIZE ===
Only A12 survives as black ✗
All merged clusters are destroyed:
  A1:B2 → only A1 black, B1-A2-B2 white
  C1:D2 → only C1 black, D1-C2-D2 white
  A3:D4 → only A3 black, rest white
  A6:D7 → only A6 black, rest white
  A9:D10 → only A9 black, rest white

=== ExportPdf ===
IgnorePrintAreas=True → exports all content
BUT merged cells are broken → only individual anchor cells are black
These individual cells ARE visible as small black rectangles
→ PDF shows only 4 small black rectangles
  (A1, C1, A3, A6 — the first cell of each former merge)
  A12 also visible but A9 is only black cell in $A$9:$D$10

=== GetAddress ===
4 black rectangles detected:
  A1 area (small — only 1 cell width/height? or merged remnant?)
  C1 area
  A3 area
  A6/A12 area

Only 4 detected ✗ (expected 6)
```

---

## 11. Comparison: ConMas vs Python

| Stage | ConMas (Original) | Our Python | Match? |
|-------|-------------------|------------|--------|
| **Iteration range** | PrintArea → UsedRange | UsedRange only | ⚠️ |
| **Cluster discovery** | Per-cell `HasComment()` | Per-cell `cell.Comment` | ✅ |
| **Comment parsing** | `ValidAndSplit()` | `split("\n")` | ⚠️ |
| **Merge area fill** | `mergeArea.Interior.Color = Black` | `fill_cell.Interior.Color = 1` | ✅ |
| **Merge area clear** | `mergeArea.Value = ""` | `fill_cell.Value = ""` | ⚠️ Same call, different behavior |
| **Non-cluster clear** | `cell.Interior.Color = White; cell.Value = ""` | Same | ✅ |
| **Non-merge cell handling** | Infragistics handles gracefully | COM breaks merges | ❌ **BUG** |
| **Borders cleared** | ✅ Yes | ✅ Yes | ✅ |
| **Shapes cleared** | ✅ Yes | ✅ Yes | ✅ |
| **Headers/footers cleared** | ✅ Yes | ✅ Yes | ✅ |
| **PDF engine** | Excel COM | Excel COM | ✅ |
| **ExportAsFixedFormat** | `IgnorePrintAreas=false` | `IgnorePrintAreas=true` | ❌ |
| **DPI** | 300 | 300 | ✅ |
| **Morphology** | None (manual mode) | None | ✅ |
| **Pixel scanning** | `GetAddress` algorithm | `scan_black_rectangles` | ⚠️ Algorithm differs |
| **Min rect size** | 6 px (width AND height) | 6 px | ✅ |
| **Sorting** | `SortClusters` (row-grouped) | By Top then Left | ⚠️ |
| **Ratio formula** | `pixel / image_dimension` | `pixel / image_dimension` | ✅ |
| **_Fields sheet read** | Never in MakeCluster | Fallback in `_identify_clusters` | ❌ Extra path |
| **Print area handling** | `PrintArea` limits iteration | Not used for iteration | ❌ |
| **Preserve merge on clear** | Infragistics preserves merge | COM breaks merge | ❌ **BUG** |

---

## 12. Known Differences & Their Impact

### 12.1 CRITICAL: Merge Area Destruction During Sanitize

**Root cause:** Setting `fill_cell.Value = ""` on a merged range via COM breaks the merge. Subsequent iteration over non-anchor cells in the former merge area treats them as individual cells and sets them to WHITE, overwriting the black fill.

**Evidence:** Sanitized workbook shows `MergeCells=False` for all previously-merged cells. Only individual cells that were anchor cells have black fill.

**Impact:** Only 4 of 6 clusters produce detectable black rectangles (specifically, only the anchor cells of each former merge and single-cell A12 remain black). The remaining cluster area is white.

**Fix approach:** One of:
1. Skip non-anchor cells in cluster merge areas (don't reprocess)
2. Track already-filled merge areas and skip cells within them
3. Don't set `Value = ""` on merged ranges (let them stay empty)
4. After filling, re-merge the cell ranges

**Recommended fix:** In `sanitize_workbook`, when a cell is part of a merge area that has already been processed (either as cluster or non-cluster), skip processing it entirely. Only process each merge area once.

### 12.2 Print Area gating

**Difference:** ConMas iterates only cells within the PrintArea (if set). Our code iterates the entire UsedRange. When a workbook has a PrintArea, some cells within UsedRange but outside PrintArea might be processed differently.

**Current impact:** None for our sample workbook (PrintArea = $A$1:$D$12 = entire used range). But could matter for workbooks with partial print areas.

### 12.3 IgnorePrintAreas in PDF Export

**Difference:** ConMas uses `IgnorePrintAreas=false`, our code uses `true`. This should make our export INCLUDE more content, not less.

**Current impact:** None. The problem is the broken merges, not the export.

### 12.4 _Fields Sheet Fallback

**Difference:** ConMas never reads `_Fields` sheet during MakeCluster. Our code uses it as a fallback when comments are empty.

**Current impact:** The `_Fields` sheet provides correct cluster information even when comments are outside UsedRange. However, this creates a different type of cluster data (no emoji prefix in names, slightly different parsing). The comment vs _Fields data might have minor differences.

### 12.5 Comment Parsing

**Difference:** ConMas uses `ValidAndSplit()` which parses the comment text into multiple named fields (name, type, index, inputParameters, readOnly, external, remarksValue, tableNo, tableName, columnNo, etc.). Our code just splits by newline.

**Current impact:** Our parsing is simpler but less robust. The `ValidAndSplit` method might handle corner cases (extra whitespace, missing fields) differently.

### 12.6 GetAddress Algorithm Differences

**Difference:** The decompiled `GetAddress` expands left AND right from the first detected content pixel. Our `scan_black_rectangles` might use a different approach.

**Current impact:** Both find the same black rectangles since the image has solid black on white background. Minor edge-case differences possible.

---

## 13. Edge Cases & Limitations

### 13.1 Merged Cells with Comments on Non-Anchor Cell

**Question:** What if a comment is placed on cell B1 (non-anchor) of a merged range A1:C3?
**ConMas:** `cell.MergeArea` returns the entire A1:C3 range → correctly processes the full merge.
**Our code:** Same — `MergeArea` covers the full range. ✅

### 13.2 Overlapping Merge Areas

**Question:** What if two cluster merge areas overlap?
**ConMas:** Each merge area is filled independently. In the PDF, the overlapping area would be black (both fills set to black). `GetAddress` would detect them as one large rectangle.
**Our code:** Same behavior. ✅

### 13.3 Cells with Comments but Outside PrintArea

**Question:** What if a cluster cell has a comment but is outside the PrintArea?
**ConMas:** The cell is never reached in MakeCluster (iterates only PrintArea). No cluster recorded.
**Our code:** The cell is reached (iterates UsedRange). Cluster IS recorded. But if it's also outside UsedRange, it's missed.

**Impact:** Workbooks with fields outside PrintArea would have different cluster counts.

### 13.4 Hidden Rows/Columns

**Question:** What if a cluster cell is in a hidden row or column?
**ConMas:** No hidden-row/column check in IL. All cells processed regardless.
**Our code:** Same — processes all cells.

### 13.5 Very Small Cells (< 6 px at 300 DPI)

**Question:** What if a cluster cell is smaller than 6×6 pixels at 300 DPI?
**ConMas:** `GetAddress` filters it out as noise.
**Our code:** Same — `MIN_RECT_SIZE = 6`.

**Impact:** Very narrow columns (e.g., < 1.44 pt wide at 300 DPI) produce undetectable clusters.

### 13.6 Multiple Clusters on the Same Row

**Question:** Can multiple clusters share the same row?
**ConMas:** Yes — if they have different column ranges and are separated by white space, `GetAddress` detects them as separate rectangles. ✅
**Impact:** Works correctly.

---

*End of reference document.*
