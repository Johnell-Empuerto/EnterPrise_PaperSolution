# Coordinate Algorithm — Complete Mathematical Formula

## Units

| Unit | Description | Value |
|------|-------------|-------|
| 1 point (pt) | 1/72 inch | 0.013889 inches |
| 1 inch | 72 points | — |
| Letter page | 612 × 792 pt | 8.5 × 11 inches |
| Standard DPI | 96 DPI | 96 pixels per logical inch |

## Source Data (from Excel COM)

### Column Widths and Row Heights
```
Col[n].Width  = column width in pt (from COM: Range.Columns[i].Width)
Row[n].Height = row height in pt (from COM: Range.Rows[i].Height)
```

Both obtained via `ExcelWorksheetCom.GetRange(row, col)` from LibExcelController or direct COM interop.

### Page Dimensions
```
pageWidth_pt  = PageSetup.PageWidth  (e.g., 612.0 for Letter)
pageHeight_pt = PageSetup.PageHeight (e.g., 792.0 for Letter)
```

### Print Area
```
PrintArea = PageSetup.PrintArea (e.g., "$A$1:$D$12")
```

## Step 1: Cell Position in Points

For a cluster spanning columns `colStart` to `colEnd` and rows `rowStart` to `rowEnd`:

```
cell_left_pt   = SUM(Col[n].Width)  for all n where colStart ≤ n ≤ colEnd and n < colStart? 
                 Wait — the left position is: SUM of widths of ALL columns to the LEFT of colStart

Actually from CellRange/ClusterInfo:
  Left   = colStart (1-based column index)
  Right  = colEnd   (1-based column index)
  Top    = rowStart (1-based row index)
  Bottom = rowEnd   (1-based row index)

cell_left_pt   = SUM(Col[k].Width)  for k = 1 to colStart-1
cell_top_pt    = SUM(Row[k].Height) for k = 1 to rowStart-1
cell_width_pt  = SUM(Col[k].Width)  for k = colStart to colEnd
cell_height_pt = SUM(Row[k].Height) for k = rowStart to rowEnd
```

## Step 2: Printed Origin

The `ExportPdf()` / `ExportAsFixedFormat()` operation renders the print area onto a PDF page. The printed content is positioned within the page based on:
- Page margins (PageSetup.LeftMargin, .RightMargin, .TopMargin, .BottomMargin)
- Header/Footer sizes (PageSetup.HeaderMargin, .FooterMargin)
- CenterHorizontally (PageSetup.CenterHorizontally)
- CenterVertically (PageSetup.CenterVertically)

The printed origin represents the offset from the top-left of the page to the top-left of the printed content:

```
originX_pt = PageSetup.LeftMargin           (if not centered horizontally)
             OR
             (pageWidth_pt - contentWidth_pt) / 2   (if centered)

originY_pt = PageSetup.TopMargin + HeaderMargin + HeaderSize
             (if not centered vertically)
             OR
             (pageHeight_pt - contentHeight_pt) / 2  + (if centered)
```

**CONFIRMED:** All 3 templates (546, 547, 548) use:
- Zoom = 100%
- No centering
- Default margins (0.75" left/right = 54pt, 1" top/bottom = 72pt)
- No header/footer

But the actual transform is NOT just margins — the decompiled `GetClusterSize()` in `LibExcelController.Lib.ImageUtility` takes a PDF path as input and calculates cluster sizes from the rendered PDF. This means the coordinate algorithm uses the PDF as ground truth for page geometry.

## Step 3: Coordinate Normalization

```
left_ratio   = (cell_left_pt   + originX_pt + gapH) / pageWidth_pt
top_ratio    = (cell_top_pt    + originY_pt + gapV) / pageHeight_pt
right_ratio  = (cell_left_pt   + cell_width_pt  + originX_pt + gapH) / pageWidth_pt
bottom_ratio = (cell_top_pt    + cell_height_pt + originY_pt + gapV) / pageHeight_pt
```

Where `gapH` and `gapV` are unknown correction factors (see Unknowns.md).

## Step 4: RoundEx

```
DB_value = Math.Round(ratio, 7, MidpointRounding.ToEven)
```

## Step 5: GetClusterSize Algorithm (inferred from API)

`LibExcelController.Lib.ImageUtility.GetClusterSize(strPdfPath, page, clusterCount)`:

1. Exports Excel sheet to PDF via `ExportPdf(sourceXlsPath, destPdfPath)`
2. Opens the PDF and extracts page dimensions (width/height in points)
3. For each cluster:
   a. Reads the COM cell position (Range.Left, Range.Top in points)
   b. Applies the printed origin correction
   c. Returns List<ClusterRect> with float coordinates

The fact that GetClusterSize takes a PDF path suggests it uses the PDF page dimensions as the normalization divisor (pageWidth_pt / pageHeight_pt from PDF metadata), NOT the PageSetup.PageWidth/PageHeight.

## Template-Specific Measurements

### Template 546 (FormTest - Copy.xlsx)
- Page: Letter (612 × 792 pt)
- PrintArea: $A$1:$D$12
- Columns: 4 (widths: ~48pt each)
- Rows: 12 (heights: ~14.4pt each)
- 6 clusters (comment-based)
- originX: ≈ 0 pt (leftmost cluster starts at ≈0 ratio = origin aligns with page left)
- originY: ≈ 0 pt
- GapV: ~0.87 pt needed
- GapH: ~2.04 pt needed

### Template 547 ([V3.1_Sample]アンケート用紙.xlsx)
- Page: Letter (612 × 792 pt)
- PrintArea: $B$2:$M$46
- Columns: 12 (B through M)
- Rows: 45 (2 through 46)
- 5 clusters (comment-based, names auto-generated: "Cluster0"-"Cluster4")
- originX: Significant (columns start at B, not A)
- GapV: ~4.44 pt + 0.18 pt per row (page-wrap compensation)
- GapH: ~0 pt

### Template 548 (Sample A.xlsx)
- Page: Letter (612 × 792 pt)
- PrintArea: $A$3:$G$11
- Columns: 7 (A through G)
- Rows: 9 (3 through 11)
- 2 clusters
- GapV: ~0 pt
- GapH: ~2.22 pt

## Reverse Formula (DB → COM Position)

```
COM_left_pt   = DB_left_ratio   × pageWidth  - originX - gapH
COM_top_pt    = DB_top_ratio    × pageHeight - originY - gapV
COM_right_pt  = DB_right_ratio  × pageWidth  - originX - gapH
COM_bottom_pt = DB_bottom_ratio × pageHeight - originY - gapV
```

## PDF-Based Coordinate Verification

The `GetClusterSize()` method takes the exported PDF as input. This means the actual coordinate transform uses the PDF page dimensions (from PDF metadata) as the authoritative page size, and measures cluster positions relative to the PDF page's rendered content area. This explains why simple margin-based formulas don't match — Excel's `ExportAsFixedFormat` may apply additional rendering adjustments (printer DPI, font metrics, column width rounding) that can only be recovered by analyzing the PDF output directly.

**Key takeaway:** The ONLY way to get 100% matching coordinates is to:
1. Export Excel to PDF via `ExportAsFixedFormat(xlTypePDF)` 
2. Measure actual positions of rendered content in the PDF
3. Compute ratios from PDF measurements
