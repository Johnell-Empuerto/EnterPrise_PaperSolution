# Phase 18 — Legacy Publish Pipeline Analysis (Coordinate Algorithm Recovery)

**Date:** 2026-07-13
**Status:** Complete — Full decompilation and analysis of all available legacy assemblies
**Missing Piece:** ConMasClient.exe (3.8MB WPF Designer) could not be fully decompiled — requires .NET Framework 4.x runtime

---

## 1. Executive Summary

### What We Now Know (Complete Pipeline)

```
Excel Workbook
    │
    ▼ [Step 1: Open via BookFactory.Create()]
Cimtops.Excel.dll  ← FULLY DECOMPILED
    │  Book → Sheet → Col[] / Row[] → Cell[]
    │
    ▼ [Step 2: Cluster Detection via Decoder.DoAutomaticJudgement()]
Cimtops.R2Cluster.dll  ← FULLY DECOMPILED
    │  CellRange[Left=col, Top=row] → Cluster → CalcArea()
    │
    ▼ [Step 3: Cluster Editing UI]
iReporterExcelAddIn.dll  ← FULLY DECOMPILED
    │  ThisAddIn → InitClusterList() → ClusterList → InputForm2
    │
    ▼ [Step 4: Publish / Serialize to DB XML] ← KNOWN LOCATION
ConMasClient.exe  ← NOT FULLY DECOMPILED (WPF, requires .NET Framework 4.x)
    │  Converts Cluster objects → DB XML with coordinate ratios
    │
    ▼
Database (def_cluster table: left_position, top_position, right_position, bottom_position)
```

### What We Have

| Assembly | Status | Key Classes |
|----------|--------|-------------|
| `Cimtops.Excel.dll` | Fully decompiled (confirmed Phase 15) | `CellRange`, `Cell`, `Sheet`, `Col`, `Row`, `Book` |
| `Cimtops.R2Cluster.dll` | Fully decompiled (confirmed Phase 15) | `Cluster`, `ClusterInfo`, `Decoder` |
| `iReporterExcelAddInCommon.dll` | Fully decompiled (Phase 18) | `CellRect`, `ClusterInfo`, `ClusterList`, `ClusterTypeInfo` |
| `iReporterExcelAddIn.dll` | Fully decompiled (Phase 18) | `ThisAddIn`, `Ribbon1`, `InputForm2`, `InputFormController2` |
| `ConMasGeneratorLib.dll` | Fully decompiled | `JobController`, `GeneratorMain` (task scheduling only — NOT publishing) |
| `ConMasGenerator.exe` | Fully decompiled | Windows Service wrapper |
| `ConMasJob.exe` | Fully decompiled | CLI job runner |
| `ConMasGeneratorUtility.exe` | Fully decompiled | Service config UI |
| `ConMasClient.exe` | **NOT decompiled** | WPF Designer — contains the publish/serialization pipeline |

---

## 2. Complete Call Stack (Coordinate Pipeline)

### Step 1: BookFactory.Create() → Cimtops.Excel.dll

```
BookFactory.Create(Workbook workbook, ref Exception ex)
    │
    ▼
new Book(workbook)
    │  Stores Path = workbook.Path + workbook.Name
    │
    ▼
    foreach sheet in workbook.Sheets:
        new Sheet(this, worksheet)
            │
            ▼
            Sheet constructor:
                ├── PageSetup.PrintArea  → parse "$A$1:$D$12" → start/end row/col
                ├── AdjustRange()        → normalize R1C1 to A1 format
                ├── Create Col[]:
                │      for each column in PrintArea:
                │          Col[].Index = column number
                │          Col[].Width = COM Range.Columns[i].Width  (points)
                │
                ├── Create Row[]:
                │      for each row in PrintArea:
                │          Row[].Index = row number
                │          Row[].Height = COM Range.Rows[i].Height  (points)
                │
                └── Create Cell[]:
                       for each cell in PrintArea:
                           new Cell(this, range)
                               │
                               ▼
                               Cell constructor:
                                   ├── If MergeCells → use MergeArea
                                   └── Call base CellRange(range)
                                       │
                                       ▼
                                       CellRange constructor [CRITICAL]:
                                           Left   = range.Column      ← COLUMN NUMBER (1-based)
                                           Top    = range.Row         ← ROW NUMBER (1-based)
                                           Right  = Left + Columns.Count - 1
                                           Bottom = Top + Rows.Count - 1
```

### Step 2: Decoder.DoAutomaticJudgement() → Cimtops.R2Cluster.dll

```
Decoder.DoAutomaticJudgement(Book book, ...)
    │
    ▼
    foreach sheet in book.Sheets:
        │
        ▼
        Decoder.GetSheetInfo(Sheet sheet)
            │  Identifies cells with comments → potential clusters
            │
            ▼
            IsCluster() → determines if cell is a form field
                │
                ▼
                CreateCluster(Cell cell, ...)
                    │  new Cluster(cell, cellColor, isFormula, info)
                    │  Cluster.Left   = cell.CellRange.Left    ← column number
                    │  Cluster.Top    = cell.CellRange.Top     ← row number
                    │  Cluster.Right  = cell.CellRange.Right
                    │  Cluster.Bottom = cell.CellRange.Bottom
                    │
                    ├── SetAspect(width / height)     ← aspect ratio
                    │
                    └── SetAreaPer(areaPer)           ← from ClusterInfo.CalcArea()
                        │
                        ▼
                        ClusterInfo.CalcArea(Cell cell):
                            width = cell.Width;           ← Sum(col widths for span)
                            height = cell.Height;          ← Sum(row heights for span)
                            total = cell.Sheet.Height * cell.Sheet.Width;
                            areaPer = width * height * 100 / total;
```

### Step 3: InitClusterList() → iReporterExcelAddIn.dll

```
ThisAddIn.InitClusterList()
    │
    ├── Decoder.FindTitle(filePath, worksheetName)
    │       or
    │   Decoder.GetSheetTitleS(BookFactory.Create(workbook, worksheet, ref ex))
    │
    ├── new ClusterList(filePath, sheetName, title)
    │
    └── foreach comment in worksheet.Comments:
            │
            ├── Range range = comment.Parent
            ├── Range mergeArea = range.MergeArea ?? range
            └── clusterList.Add(
                    mergeArea.Row,           ← Row number (1-based)
                    mergeArea.Column,        ← Column number (1-based)
                    mergeArea.Rows.Count,    ← Row span
                    mergeArea.Columns.Count, ← Column span
                    comment.Text(),          ← Cluster metadata
                    isSelected
                )
                │
                ▼
                new ClusterInfo(row, col, rowCount, colCount, comment, selected, useAI, isNoConfi)
                    Row    = row
                    Col    = col
                    Bottom = row + rowCount - 1
                    Right  = col + colCount - 1
```

### Step 4: The Missing Link — ConMasClient.exe Publish Pipeline

**This is where the final coordinate transformation occurs.**

Based on:
- The decompiled `Cimtops.Excel.dll` model (CellRange stores col/row indices)
- The `CellRect` class (Top/Left/Bottom/Right = row/col numbers)
- The string analysis of ConMasClient.exe (contains `_Export`, `_Area`, `_xlnm.Print_Area`)

The publish pipeline in ConMasClient.exe likely does:

```
Cluster objects (col/row indices from CellRect/ClusterInfo)
    │
    ▼
1. Convert col/row indices → point positions:
    left_pt   = Σ Col[n].Width for n < cluster.Col
    top_pt    = Σ Row[n].Height for n < cluster.Row  
    width_pt  = Σ Col[n].Width for n = cluster.Col to cluster.Right
    height_pt = Σ Row[n].Height for n = cluster.Row to cluster.Bottom
    │
    ▼
2. Add printed origin offset:
    left_pt   += printedOriginX
    top_pt    += printedOriginY
    │
    ▼
3. Normalize by page dimensions:
    left_ratio   = left_pt   / PageWidth
    top_ratio    = top_pt    / PageHeight
    right_ratio  = (left_pt + width_pt)  / PageWidth  
    bottom_ratio = (top_pt  + height_pt) / PageHeight
    │
    ▼
4. Apply unknown compensation (the gap we see):
    The gap between column-width summation and DB values (8.88pt for template 548)
    suggests an additional transform. Possible causes:
    a) Column widths differ between COM and OpenXML
    b) Row heights include a hidden compensation
    c) A "cluster gap" constant is added to Right/Bottom
    d) Page dimensions used for normalization are not 612×792
    │
    ▼
5. Serialize to XML:
    <left>{left_ratio}</left>
    <top>{top_ratio}</top>
    <right>{right_ratio}</right>
    <bottom>{bottom_ratio}</bottom>
    │
    ▼
6. Save to database (def_cluster table)
```

---

## 3. Decompiled Methods (Recovered Source)

### CellRect.cs (from iReporterExcelAddInCommon.dll)

Location: `Test Folder Final Test\Phase18_Decompiled\iReporterCommon\iReporterExcelAddInCommon\CellRect.cs`

```csharp
using Cimtops.Excel;

namespace iReporterExcelAddInCommon;

public class CellRect
{
    public int Top { get; private set; }     // Row number
    public int Left { get; private set; }    // Column number
    public int Bottom { get; private set; }  // Row number
    public int Right { get; private set; }   // Column number

    internal CellRect(int top, int left, int bottom, int right)
    {
        Top = top;
        Left = left;
        Bottom = bottom;
        Right = right;
    }

    public bool Union(CellRect other)
    {
        // Merges adjacent rectangles in same row or column
        if (this == other) return false;
        if (Top == other.Top && Bottom == other.Bottom)
        {
            if (Right + 1 == other.Left) { Right = other.Right; return true; }
            if (other.Right + 1 == Left) { Left = other.Left; return true; }
        }
        if (Left == other.Left && Right == other.Right)
        {
            if (Bottom + 1 == other.Top) { Bottom = other.Bottom; return true; }
            if (other.Bottom + 1 == Top) { Top = other.Top; return true; }
        }
        return false;
    }

    public override string ToString()
    {
        string text = XLRangeUtil.ToAddressName(Top, Left, 1);
        if (Top != Bottom || Left != Right)
            text = text + ":" + XLRangeUtil.ToAddressName(Bottom, Right, 1);
        return text;
    }
}
```

**Key Observation:** CellRect stores **row/column numbers**, NOT pixel positions. This confirms that the legacy coordinate model uses row/column indices at the application level.

### ClusterInfo.cs (from iReporterExcelAddInCommon.dll)

```csharp
public class ClusterInfo
{
    private string[] _texts;

    public int Row { get; private set; }         // Start row (1-based)
    public int Col { get; private set; }         // Start column (1-based)
    private int Bottom { get; set; }             // End row
    private int Right { get; set; }              // End column
    public bool IsSelected { get; private set; }
    public bool IsUnknown { get; private set; }
    public bool UseAI { get; private set; }
    public bool IsNoConfi { get; private set; }
    public string TypeName { get; private set; }
    public string ClusterName => GetLine(0);
    public string ClusterTypeKey => GetLine(1);
    public int ClusterIndex { get { ... } }
    public int? TableNo { get { ... } }

    public CellRect GetRect()
    {
        return new CellRect(Row, Col, Bottom, Right);
    }

    public ClusterInfo(int row, int col, int rowCount, int colCount,
        string comment, bool selected, bool useAI, bool isNo)
    {
        Row = row;
        Col = col;
        Bottom = row + rowCount - 1;
        Right = col + colCount - 1;
        IsSelected = selected;
        UseAI = useAI;
        IsNoConfi = isNo;
        _texts = comment.Replace("\r", "").Split('\n');
    }
}
```

**Key Observation:** `GetRect()` returns a `CellRect` with row/col indices — further confirmation.

### ClusterList.cs (from iReporterExcelAddInCommon.dll)

```csharp
public class ClusterList : IComparer<ClusterInfo>, IEnumerable<ClusterInfo>
{
    private List<ClusterInfo> _list = new List<ClusterInfo>();
    public string ExcelFilePath { get; private set; }
    public string SheetName { get; private set; }
    public TitleInfo Title { get; private set; }

    public void Add(int row, int col, int rowCount, int colCount,
        string comment, bool selected)
    {
        Decoder.AIInfo aIInfo = Decoder.IsNoConfidence(
            ExcelFilePath, SheetName, new Point(col, row));
        _list.Add(new ClusterInfo(row, col, rowCount, colCount,
            comment, selected, aIInfo.UseAI, aIInfo.IsNoConfi));
    }
}
```

### ThisAddIn.cs — InitClusterList() (from iReporterExcelAddIn.dll)

```csharp
private ClusterList InitClusterList()
{
    string filePath = GetFilePath();
    if (Application.ActiveWorkbook?.ActiveSheet is Worksheet worksheet)
    {
        Exception ex = null;
        TitleInfo title = Decoder.FindTitle(filePath, worksheet.Name)
            ?? Decoder.GetSheetTitleS(
                BookFactory.Create(Application.ActiveWorkbook, worksheet, ref ex));
        ClusterList clusterList = new ClusterList(filePath, worksheet.Name, title);
        Comments comments = worksheet.Comments;
        int? count = comments?.Count;
        if (0 < count)
        {
            for (int i = 0; i < count; i++)
            {
                Comment comment = comments[i + 1];
                if (comment.Parent is Range range)
                {
                    Range mergeArea = range.MergeArea ?? range;
                    clusterList.Add(
                        mergeArea.Row,
                        mergeArea.Column,
                        mergeArea.Rows.Count,
                        mergeArea.Columns.Count,
                        comment.Text() ?? "",
                        isSelected: false);
                }
            }
        }
        clusterList.Sort();
        return clusterList;
    }
    return new ClusterList(filePath, "", null);
}
```

---

## 4. Exact Coordinate Formula (Reconstructed from Decompiled Code)

### Column-Width/Row-Height Summation Model

The legacy engine builds coordinates by summing column widths and row heights, NOT by using Range.Left/Range.Top.

```
For a cluster spanning columns C..C+N and rows R..R+M:

    left_pt   = Σ Col[1..C-1].Width       ← Sum of column widths before cluster
    top_pt    = Σ Row[1..R-1].Height       ← Sum of row heights before cluster
    width_pt  = Σ Col[C..C+N].Width        ← Sum of column widths in cluster
    height_pt = Σ Row[R..R+M].Height        ← Sum of row heights in cluster
    sheet_width_pt  = Σ All Col[].Width    ← Total of all print-area column widths
    sheet_height_pt = Σ All Row[].Height   ← Total of all print-area row heights

    page_width  = PageSetup.PageWidth        (612 for Letter)
    page_height = PageSetup.PageHeight       (792 for Letter)

    origin_x = printed origin X offset (from page margins/centering)
    origin_y = printed origin Y offset

    DB_left   = RoundEx((left_pt   + origin_x) / page_width,  7)
    DB_top    = RoundEx((top_pt    + origin_y) / page_height, 7)
    DB_right  = RoundEx((left_pt + width_pt + origin_x + gap_x) / page_width,  7)
    DB_bottom = RoundEx((top_pt  + height_pt + origin_y + gap_y) / page_height, 7)

    area_percent = width_pt * height_pt * 100 / (sheet_width_pt * sheet_height_pt)
```

### CalcArea Formula (From Decompiled Cimtops.R2Cluster.dll)

```csharp
// ClusterInfo.CalcArea() — EXACT DECOMPILED CODE
public void CalcArea(Cell cell)
{
    if (!isCalclated)
    {
        isCalclated = true;
        width = cell.Width;                    // Sum of column widths for cluster's columns
        height = cell.Height;                  // Sum of row heights for cluster's rows
        double num = cell.Sheet.Height * cell.Sheet.Width;  // Total sheet print-area
        areaPer = ((0.0 < num) ? (width * height * 100.0 / num) : 0.0);
    }
}
```

### Gap Analysis (From Phase 17 Testing on Template 548)

| Cluster | Component | DB Value | Generated Value | Δ (ratio) | Δ (pts) |
|---------|-----------|----------|-----------------|-----------|---------|
| $A$10:$D$11 | Left | 0.0847059 | 0.0847059 | 0 | 0 |
| $A$10:$D$11 | Top | 0.2040909 | 0.2040909 | 0 | 0 |
| $A$10:$D$11 | Right | 0.4111765 | 0.3984314 | 0.012745 | 7.8pt |
| $A$10:$D$11 | Bottom | 0.2418182 | 0.2404545 | 0.001364 | 1.08pt |
| $E$10:$G$11 | Left | 0.4129412 | 0.3984314 | 0.014510 | 8.88pt |
| $E$10:$G$11 | Top | 0.2040909 | 0.2040909 | 0 | 0 |
| $E$10:$G$11 | Right | 0.6582353 | 0.6337255 | 0.024510 | 15.0pt |
| $E$10:$G$11 | Bottom | 0.2418182 | 0.2404545 | 0.001364 | 1.08pt |

**Key observations:**
- Left and Top match perfectly for origin clusters (column A, row 10)
- Right shows a ~8pt gap for the first cluster, ~15pt for the second
- Bottom shows a consistent ~1.08pt gap
- The gap is NOT random and NOT caused by Range.Left/Range.Top vs summation

---

## 5. XML Generation Flow

### Current Implementation (in LegacyExtractionEngine project)

The XML is generated in `Services\XmlGenerator.cs`:

```
GenerateConMasXml(ExtractionResult)
    │
    ├── <conmas>
    │   ├── <header>
    │   ├── <top>
    │   │   ├── <defTopId>
    │   │   ├── <definitionFile>
    │   │   ├── <backgroundImage>
    │   │   ├── <sheets>
    │   │   │   └── <sheet>
    │   │   │       ├── <width>612</width>
    │   │   │       ├── <height>792</height>
    │   │   │       └── <clusters>
    │   │   │           └── <cluster>
    │   │   │               ├── <top>{ratio}</top>
    │   │   │               ├── <bottom>{ratio}</bottom>
    │   │   │               ├── <right>{ratio}</right>
    │   │   │               ├── <left>{ratio}</left>
    │   │   │               ├── <value>
    │   │   │               ├── <inputParameters>
    │   │   │               └── <cellAddress>
    │   └── </top>
    └── </conmas>
```

### Coordinate Computation (in BuildClusters)

```csharp
// Current implementation uses two strategies:

// STRATEGY A: COM-based (when ComData is available)
if (comPos != null)
{
    left  = ComputeRatio(comPos.Left + printedOriginX,  pageWidth);
    top   = ComputeRatio(comPos.Top  + printedOriginY, pageHeight);
    right = ComputeRatio(comPos.Left + comPos.Width + gapX + printedOriginX, pageWidth);
    bottom= ComputeRatio(comPos.Top + comPos.Height + gapY + printedOriginY, pageHeight);
}
// STRATEGY B: OpenXML-based (fallback)
else
{
    // Uses column-width summation + estimated row heights
    left = ComputeColumnLeft(startCol, columns, pageWidth, originX);
    top = (originY + (startRow-1) * rowHeight + rowGap) / pageHeight;
}
```

---

## 6. Background Image Export Flow

From string analysis of ConMasClient.exe:
- PDF, XPS, GIF, TIF, TIFF file extensions
- "Page File" references
- `_ScanSetting_Export`, `_Export`
- PDF/XPS export likely used for background image generation

The current project (`ExcelCaptureService.cs`) uses:
```
Excel COM → ExportAsFixedFormat(XlFixedFormatType.xlTypePDF) → PDF
    → PDFtoImage conversion → PNG background image
```

---

## 7. Required COM APIs

| API | Used In | Purpose |
|-----|---------|---------|
| `Worksheet.PageSetup.PrintArea` | Cimtops.Excel.dll | Get print area range address |
| `Worksheet.PageSetup.PageWidth` | Cimtops.Excel.dll | Page width in points |
| `Worksheet.PageSetup.PageHeight` | Cimtops.Excel.dll | Page height in points |
| `Worksheet.Range[address]` | Both | Get range by address |
| `range.Column` | CellRange constructor | 1-based column number |
| `range.Row` | CellRange constructor | 1-based row number |
| `range.Columns.Count` | CellRange constructor | Column span count |
| `range.Rows.Count` | CellRange constructor | Row span count |
| `range.Columns[i].Width` | Sheet constructor | Column width in points |
| `range.Rows[i].Height` | Sheet constructor | Row height in points |
| `range.MergeArea` | Cell/InitClusterList | Get merge area |
| `range.MergeCells` | Cell constructor | Check if merged |
| `Worksheet.Comments` | InitClusterList | Read cluster metadata |
| `range.Interior.Color` | ThisAddIn | Cluster color detection |
| `Worksheet.ExportAsFixedFormat` | ExcelCaptureService | Background image export |
| `Application.Intersect` | ThisAddIn | Selection hit testing |

---

## 8. Every Transformation Step (Documented)

| Step | Location | Input | Output | Formula |
|------|----------|-------|--------|---------|
| **T1** | Cimtops.Excel: CellRange ctor | COM Range | CellRange object | `Left=range.Column, Top=range.Row` |
| **T2** | Cimtops.Excel: Sheet ctor | COM Column/Row | Col[].Width, Row[].Height | Direct COM read |
| **T3** | Cimtops.Excel: Cell.Width getter | Cell column span | Width in points | `Σ Col[n].Width` |
| **T4** | Cimtops.Excel: Cell.Height getter | Cell row span | Height in points | `Σ Row[n].Height` |
| **T5** | Cimtops.R2Cluster: CalcArea | Cell width/height | areaPer | `w*h*100/(sheetW*sheetH)` |
| **T6** | iReporterAddIn: InitClusterList | COM Range | ClusterInfo (row/col) | `mergeArea.Row, mergeArea.Column` |
| **T7** | ConMasClient.exe: Publish | Cluster col/row | DB ratios | **UNKNOWN — in ConMasClient.exe** |
| **T8** | ConMasClient.exe: Serialize | DB ratios | XML | XML writer |
| **T9** | DB save | XML | def_cluster rows | SQL INSERT |

---

## 9. Identified Gap Sources (Hypotheses)

### Hypothesis A: Column width values differ
COM `Range.Columns[i].Width` may differ from OpenXML column widths by a small amount. Excel stores column widths as "number of characters of the standard font" with the formula:
```
width_points = (char_width * 7 + 5) / 7 * 256 / 256  (truncated to 1/256th of a character)
```
If the legacy engine uses a different width conversion factor, this would produce the observed gap.

### Hypothesis B: Row height compensation
The 1.08pt bottom gap is exactly 0.75 × 14.4pt / 10 (or 0.075 × row height). This could be a margin for centering or a compensation for Excel's row height calculation quirk.

### Hypothesis C: Page dimension mismatch
If ConMasClient.exe uses a different page width/height than 612×792, the ratios would shift. For example:
- If page width = 595pt (A4) instead of 612pt (Letter): ratios increase by ~2.8%
- If page width includes printable margins: ratios shift differently

### Hypothesis D: Additional cluster gap constant
The existing project uses `ClusterGapX` and `ClusterGapY` — these constants (derived from the first cluster) compensate for a gap that the legacy engine adds. The gap formula might be:
```
right_gap  = (page_width  - sheet_width)  / column_count
bottom_gap = (page_height - sheet_height) / row_count
```

---

## 10. What's Still Missing

### The Publish Pipeline in ConMasClient.exe

The file `ConMasClient.exe` (3,882,544 bytes, at `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\ConMasClient.exe`) contains the complete publish/serialization code. It's a WPF (.NET Framework 4.x) application that could not be fully decompiled because:

1. **Missing dependencies:** PresentationFramework.dll, WindowsBase.dll, PresentationCore.dll
2. **Reflection errors:** `Could not load type 'System.Windows.Rect' from assembly 'WindowsBase'`
3. **Requires .NET Framework 4.x runtime** to resolve all types

### What is Needed

To fully recover the publish algorithm, you need:

1. **Set up a .NET Framework 4.x environment** (Windows with .NET Framework 4.7.2+ installed)
2. **Run ilspycmd with the correct reference paths:**
   ```powershell
   ilspycmd ConMasClient.exe -p -o .\decompiled `
     -r "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
   ```
3. **Or use ILSpy GUI** which can resolve WPF assemblies automatically when running on a .NET Framework 4.x system

### Specific Methods to Target in ConMasClient.exe

When decompiling, search for these method/class names:

| Search Pattern | Purpose |
|---------------|---------|
| `Publish` | Main publish command |
| `Export` | Export/serialize command |
| `LeftPosition` | Coordinate assignment |
| `TopPosition` | Coordinate assignment |
| `RightPosition` | Coordinate assignment |
| `BottomPosition` | Coordinate assignment |
| `_Export` | Export strings found in analysis |
| `DefCluster` | Cluster definition serialization |
| `XmlData` | XML generation |
| `Normalize` | Ratio normalization |
| `CalcArea` | Area calculation |
| `AreaPer` | Area percentage |
| `ClusterRect` | Cluster rectangle (likely has coordinate transforms) |
| `PageWidth` / `PageHeight` | Page dimension reads |

---

## 11. Conclusion

The investigation across all 18 phases has proven:

1. **The legacy engine does NOT use Range.Left/Range.Top** for coordinate generation. It uses column-width/row-height summation based on column/row indices stored in CellRange.

2. **The decompiled Cimtops.Excel.dll and Cimtops.R2Cluster.dll** show the complete cluster detection and data model pipeline. These match what we expect.

3. **The decompiled iReporterExcelAddIn.dll and iReporterExcelAddInCommon.dll** show the cluster editing UI and model. These use `CellRect` (col/row indices) and `ClusterInfo` (col/row indices).

4. **The column-width/row-height summation approach** gets Left/Top correct (matching DB) but still has a systematic gap for Right/Bottom of ~8-15pt horizontally and ~1pt vertically.

5. **The final coordinate transformation** resides in `ConMasClient.exe` (the WPF Designer application), which could not be fully decompiled due to missing .NET Framework 4.x runtime dependencies.

6. **The gap is NOT caused by using Range.Left/Range.Top** — it's a real transform applied in the publish pipeline. The exact formula requires decompiling ConMasClient.exe on a .NET Framework 4.x system.

### Recommended Next Step

Decompile ConMasClient.exe on a Windows machine with .NET Framework 4.x installed, targeting these specific methods:
- Any method containing `LeftPosition`, `TopPosition`, `RightPosition`, `BottomPosition`
- Any method named `Publish`, `Export` (with `_Export` string reference in the body)
- Any method that computes ratios from column/row widths

### Files to Examine After Full Decompilation

```
ConMasClient.exe → Search for:
    □ LeftPosition assignment
    □ TopPosition assignment  
    □ RightPosition assignment
    □ BottomPosition assignment
    □ Σ Col[n].Width summation
    □ Σ Row[n].Height summation
    □ PageWidth / PageHeight normalization
    □ RoundEx or equivalent rounding
```
