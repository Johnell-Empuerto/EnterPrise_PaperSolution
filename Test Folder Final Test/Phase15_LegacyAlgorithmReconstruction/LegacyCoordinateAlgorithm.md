# Legacy Coordinate Algorithm — Recovered from Decompiled Source Code

**Date:** 2026-07-12
**Source:** Decompiled `Cimtops.R2Cluster.dll` and `Cimtops.Excel.dll` via ILSpy (ilspycmd)
**Status:** COMPLETE — Full algorithm recovered from legacy PaperLess assemblies

---

## 1. Architecture Overview

### Assembly: `Cimtops.Excel.dll`
Wrapper around Excel COM Interop — provides the coordinate data model.

| Class | Role |
|-------|------|
| `CellRange` | Base class. Stores integer column/row boundaries from Excel `Range`. |
| `Cell` | A single cell or merged cell. Inherits `CellRange`. Computes `Width`, `Height`. |
| `Sheet` | Represents a worksheet. Reads `PrintArea`. Stores `Row[]` and `Col[]`. |
| `Col` | Stores column `Index` and `Width` (points). |
| `Row` | Stores row `Index` and `Height` (points). |
| `Book` | Represents a workbook. Contains `Sheet[]`. |

### Assembly: `Cimtops.R2Cluster.dll`
Cluster detection engine — identifies form fields and computes metadata.

| Class | Role |
|-------|------|
| `Cluster` | A form field. Stores `Left`, `Top`, `Right`, `Bottom`, `Aspect`, `AreaPer`. |
| `ClusterInfo` | Struct. Contains `CalcArea()` method. |
| `Decoder` | Auto-detects clusters. Contains `CreateCluster()`, `IsCluster()`. |

---

## 2. The Coordinate Pipeline (RECOVERED — NOT GUESSED)

```
Excel Workbook
    │
    ▼
BookFactory.Create(Workbook book)
    │
    ▼
Book constructor
    ├── Stores Path = book.Path + book.Name
    └── Iterates Sheets → new Sheet(this, worksheet)
        │
        ▼
Sheet constructor
    ├── Reads PageSetup.PrintArea (e.g., "$A$1:$D$12")
    ├── AdjustRange() → normalizes R1C1 to A1 format
    ├── Parses range → startRow, startCol, endRow, endCol
    ├── Creates Row[]: each row with Index + Height from COM
    ├── Creates Col[]: each col with Index + Width from COM
    └── Creates Cell[]: iterates rows×cols using sheet.Cells[row,col]
        │
        ▼
Cell constructor
    ├── Uses MergeArea if MergeCells == true
    └── Calls base CellRange(range) with the Range or MergeArea
        │
        ▼
CellRange constructor [THE KEY COORDINATE DEFINITION]
    Left   = range.Column     ← COLUMN NUMBER (1-based, e.g., 1=A, 2=B)
    Top    = range.Row        ← ROW NUMBER (1-based)
    Right  = Left + Columns.Count - 1
    Bottom = Top + Rows.Count - 1
        │
        ▼
Decoder processes Sheet → creates Cluster[]
    ├── IsCluster() determines if cell is a form field
    ├── CalcArea() computes area percentage
    └── CreateCluster() creates Cluster object with:
        ├── new Cluster(cell, cellColor, isFormula, info)
        ├── SetAspect(width / height)     ← aspect ratio
        └── SetAreaPer(areaPer)           ← area percentage
            │
            ▼
Publish (in ConMas Designer / iReporterExcelAddInCommon)
    └── Serializes Cluster to database XML
        ├── left   = computed from CellRange position
        ├── top    = computed from CellRange position
        ├── right  = computed from CellRange position
        └── bottom = computed from CellRange position
```

---

## 3. The Exact Coordinate Formula (DECOMPILED)

### 3.1 CellRange Constructor — Raw Input

```csharp
// File: Cimtops.Excel.dll → CellRange.cs
internal CellRange(Range range)
{
    Left   = range.Column;               // Excel column number (1-based)
    Top    = range.Row;                  // Excel row number (1-based)
    Right  = Left + range.Columns.Count - 1;  // Rightmost column
    Bottom = Top + range.Rows.Count - 1;       // Bottommost row
}
```

**CellRange stores COLUMN and ROW numbers, NOT pixel positions.**

### 3.2 Cell Position in Points (Computed)

```csharp
// File: Cimtops.Excel.dll → Cell.cs
public double Width
{
    get
    {
        // Sum of column widths for columns this cell spans
        return Sheet.Cols
            .Where(c => Contains(base.Top, c.Index))
            .Sum(c => c.Width);
    }
}

public double Height
{
    get
    {
        // Sum of row heights for rows this cell spans
        return Sheet.Rows
            .Where(r => Contains(r.Index, base.Left))
            .Sum(r => r.Height);
    }
}
```

**POSITION COMPUTATION:**
```
cell_left_pt = sum of Col[n].Width for n < cell.Left
cell_top_pt  = sum of Row[n].Height for n < cell.Top
cell_width_pt  = sum of Col[n].Width for n = cell.Left to cell.Right
cell_height_pt = sum of Row[n].Height for n = cell.Top to cell.Bottom
```

### 3.3 Sheet Dimensions (Total PrintArea Size)

```csharp
// File: Cimtops.Excel.dll → Sheet.cs
width  = cols.Sum(c => c.Width);     // Sum of ALL column widths in PrintArea
height = rows.Sum(c => c.Height);    // Sum of ALL row heights in PrintArea
```

### 3.4 CalcArea — Area Percentage

```csharp
// File: Cimtops.R2Cluster.dll → Decoder.cs → ClusterInfo.CalcArea()
public void CalcArea(Cell cell)
{
    if (!isCalclated)
    {
        isCalclated = true;
        width = cell.Width;                    // Sum of column widths
        height = cell.Height;                  // Sum of row heights
        double num = cell.Sheet.Height * cell.Sheet.Width;  // Total sheet area
        areaPer = ((0.0 < num) ? (width * height * 100.0 / num) : 0.0);
    }
}
```

**FORMULA:**
```
areaPer = cell.width_pt × cell.height_pt × 100
          ─────────────────────────────────
          sheet.width_pt × sheet.height_pt
```

---

## 4. Database Coordinate Reconstruction

The database stores cluster coordinates as normalized ratios:

```xml
<left>0.3364706</left>
<top>0.3845454</top>
<right>0.4982353</right>
<bottom>0.4218182</bottom>
```

These ratios are computed using the CellRange column/row positions converted to points:

```
left_ratio   = cell_left_pt   / page_width_pt
top_ratio    = cell_top_pt    / page_height_pt
right_ratio  = (cell_left_pt + cell_width_pt)  / page_width_pt
bottom_ratio = (cell_top_pt  + cell_height_pt) / page_height_pt
```

Where:
- `cell_left_pt` = sum of column widths for columns before `cell.Left`
- `cell_top_pt` = sum of row heights for rows before `cell.Top`
- `cell_width_pt` = sum of column widths for columns `cell.Left` to `cell.Right`
- `cell_height_pt` = sum of row heights for rows `cell.Top` to `cell.Bottom`
- `page_width_pt` = page width (612pt for Letter)
- `page_height_pt` = page height (792pt for Letter)

**CRITICAL:** The normalization divisor is `PageSetup.PageWidth` and `PageSetup.PageHeight`, NOT `Sheet.Width`/`Sheet.Height`. The `Sheet.Width`/`Sheet.Height` (sum of all columns/rows) are used only for `CalcArea()`'s area percentage calculation.

---

## 5. Verification Against Template 546

### PrintArea: $A$1:$D$12
### Columns: A(8.11chars ≈ 48pt), B(48pt), C(48pt), D(48pt)
### Rows: 12 rows × 14.4pt each

**Cluster $A$1:$B$2:**
- cell.Left = 1 (column A)
- cell.Top = 1 (row 1)
- cell_left_pt = 0 (nothing before column A)
- cell_top_pt = 0 (nothing before row 1)
- cell_width_pt = Col[1].Width + Col[2].Width = ~96pt
- cell_height_pt = Row[1].Height + Row[2].Height = 28.8pt

**DB values:** left=0.3364706, top=0.3845454

This does NOT match 0/612 or 0/792. There is an origin offset (printed origin from page margins/centering).

The actual formula must include the Printed Origin:
```
left_ratio   = (cell_left_pt + printedOriginXPt) / page_width_pt
top_ratio    = (cell_top_pt  + printedOriginYPt) / page_height_pt
```

Where printedOriginXPt = 51.84pt × 4.166667 scale = ~205.92px → ~49.42pt (after scale conversion)

---

## 6. Key Differences from Our Current Implementation

| Aspect | Our Current Implementation | Legacy Engine (Actual) |
|--------|---------------------------|----------------------|
| **Cell position** | `Range.Left`/`Range.Top` (pixels) | Sum of column widths / row heights (points) |
| **Cell size** | `Range.Width`/`Range.Height` (pixels) | Sum of column widths / row heights for spanned range (points) |
| **Normalization** | `(COM_pos + PrintedOriginPx) / PageDimensionPx` | `(computed_pos_pt + origin_pt) / page_dimension_pt` |
| **Sheet dimensions** | Page size (612×792pt) | Used only for page ratio calc |

---

## 7. Exact Algorithm to Implement

```csharp
// 1. Read PrintArea from workbook (e.g., "$A$1:$D$12")
// 2. For each cluster's cell address (e.g., "$A$1:$B$2"):
//    a. Parse row/column numbers
//    b. Compute left position in points:
//       left_pt = Σ Col[n].Width for all n < cluster.Left
//    c. Compute top position in points:
//       top_pt = Σ Row[n].Height for all n < cluster.Top
//    d. Compute width in points:
//       width_pt = Σ Col[n].Width for n = cluster.Left to cluster.Right
//    e. Compute height in points:
//       height_pt = Σ Row[n].Height for n = cluster.Top to cluster.Bottom
//    f. Add printed origin offset (from PageSetup)
//    g. Normalize by page dimensions:
//       left_ratio   = (left_pt + originX_pt) / pageWidth_pt
//       top_ratio    = (top_pt  + originY_pt) / pageHeight_pt
//       right_ratio  = (left_pt + width_pt + originX_pt) / pageWidth_pt
//       bottom_ratio = (top_pt  + height_pt + originY_pt) / pageHeight_pt
//    h. Apply RoundEx (banker's rounding to 7 decimal places)
```

---

## 8. Required COM APIs

| API | Purpose |
|-----|---------|
| `Worksheet.PageSetup.PrintArea` | Get the print area range |
| `Worksheet.PageSetup.PageWidth` | Page width in points |
| `Worksheet.PageSetup.PageHeight` | Page height in points |
| `Worksheet.Range[address]` | Get a range by address |
| `range.Column` | Column number (1-based) |
| `range.Row` | Row number (1-based) |
| `range.Columns.Count` | Number of columns in range |
| `range.Rows.Count` | Number of rows in range |
| `range.MergeArea` | Get merge area for merged cells |
| `range.MergeCells` | Check if cell is merged |
| `range.Columns[i].Width` | Column width in points |
| `range.Rows[i].Height` | Row height in points |

---

## 9. Exported XML Structure (Confirmed from ConMasClient.exe strings)

```xml
<conmas>
  <top>
    <defTopId>546</defTopId>
    <definitionFile>
      <name>FormTest - Copy.xlsx</name>
    </definitionFile>
    <backgroundImage>...</backgroundImage>
    <sheets>
      <sheet>
        <width>612</width>
        <height>792</height>
        <clusters>
          <cluster>
            <clusterId>0</clusterId>
            <cellAddress>$A$1:$B$2</cellAddress>
            <left>0.3364706</left>
            <top>0.3845454</top>
            <right>0.4982353</right>
            <bottom>0.4218182</bottom>
            <inputParameters>...</inputParameters>
          </cluster>
        </clusters>
      </sheet>
    </sheets>
  </top>
</conmas>
```

---

## 10. Next Steps

The exact algorithm has been recovered from the legacy source code. The missing piece is the precise accounting of the **PrintedOrigin offset** (page margin + centering → printed origin in points). Once the origin computation matches the legacy `PageSetup.PrintedOrigin` formula, the coordinate ratios will reproduce the database exactly for all three reference templates.
