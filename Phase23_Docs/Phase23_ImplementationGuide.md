# Phase 23: Implementation Guide — Replicating GetClusterSize() in ASP.NET Core

**Date**: 2026-07-13
**Goal**: 100% coordinate match with legacy pixel-scanning algorithm

---

## 1. The Challenge

The legacy `GetClusterSize()` uses a pixel-scanning approach:
1. Export sanitized Excel → PDF (via Excel COM)
2. Render PDF → 200 DPI bitmap (via PdfiumViewer)
3. Morphological close (via OpenCV)
4. Scan pixels for black rectangles → normalized ratios

In ASP.NET Core, we CANNOT use Excel COM, PdfiumViewer, or OpenCvSharp (Windows-only). We need a cross-platform alternative.

---

## 2. Recommended Approach: Column-Width Replication with Calibration

### 2.1 Why Not Full Pixel Scanning?

| Requirement | Legacy | ASP.NET Core Alternative | Feasibility |
|-------------|--------|-------------------------|-------------|
| Excel COM | `Application.Workbooks.Open` | ClosedXML/DocumentFormat.OpenXml | ✅ |
| PDF Export | COM `ExportAsFixedFormat` | ClosedXML PDF export or server Excel | ⚠️ Complex |
| PDF Render | PdfiumViewer 200 DPI | PdfiumViewer (Linux via SkiaSharp) | ⚠️ Heavy |
| Morphology | OpenCvSharp 3×3 close | ImageSharp manual kernel | ⚠️ Complex |
| Pixel scan | Native byte array | Pure C# | ✅ Easy |

Full replication requires 4+ heavy dependencies. **Not recommended**.

### 2.2 The Mathematical Alternative

The pixel scanner inherently captures:
1. Page margins (from PageSetup)
2. Centering (horizontal/vertical)
3. Column widths + row heights
4. Page dimensions

These are ALL accessible from the Excel OpenXML file without rendering anything.

### 2.3 Formula Derivation

From the pixel-scan data we know:

```
Left   = (MarginLeft + CenteringOffsetX + Σ(ColumnWidths[1..col-1])) / PageWidth
Right  = (MarginLeft + CenteringOffsetX + Σ(ColumnWidths[1..col+numCols-1])) / PageWidth  
Top    = (MarginTop + CenteringOffsetY + Σ(RowHeights[1..row-1])) / PageHeight
Bottom = (MarginTop + CenteringOffsetY + Σ(RowHeights[1..row+numRows-1])) / PageHeight
```

Where:
- `MarginLeft`, `MarginTop` = from `PageSetup.LeftMargin`, `PageSetup.TopMargin` (in points, default 0.75in = 54pt)
- `CenteringOffsetX` = `max(0, (PageWidth - MarginLeft - MarginRight - ContentWidth) / 2)` when `PageSetup.CenterHorizontally = true`
- `CenteringOffsetY` = `max(0, (PageHeight - MarginTop - MarginBottom - ContentHeight) / 2)` when `PageSetup.CenterVertically = true`
- `ContentWidth` = sum of all column widths in print area
- `ContentHeight` = sum of all row heights in print area

### 2.4 Column Width: OpenXML to Points Conversion

OpenXML stores column width in "character units" — the count of '0' characters that fit. COM `Range.Width` returns points. The conversion:

```csharp
// Standard Excel internal formula:
// pixelWidth = columnWidth * maxDigitWidth + (5 pixel padding)
// points = pixelWidth * 72 / 96  (at 96 DPI screen reference)
static double ColumnWidthToPoints(double columnWidth, double maxDigitWidth = 7)
{
    double pixels = columnWidth * maxDigitWidth + 5;     // Standard Excel formula
    double points = pixels * 72.0 / 96.0;                // Convert 96 DPI → points
    return points;
}

// Row height is stored directly in points in OpenXML
static double RowHeightToPoints(double rowHeight) => rowHeight;
```

The `maxDigitWidth` depends on the Normal style font:
- Calibri 11pt (default): 7 pixels
- Arial 10pt: ~7 pixels  
- MS Mincho 10pt (CJK): 10 pixels

### 2.5 Page Dimensions

Get page dimensions from the PDF file (not the Excel):
```csharp
// Use any PDF library (PdfPig, PdfiumViewer, etc.)
double pageWidth  = GetPdfPageWidth(pdfPath);   // e.g., 612 pts (Letter portrait)
double pageHeight = GetPdfPageHeight(pdfPath);  // e.g., 792 pts (Letter portrait)
```

---

## 3. Implementation Plan

### Phase 23a: OpenXML Reading Layer

```csharp
public class OpenXmlReader
{
    // Read column widths from worksheet
    double[] GetColumnWidths(string xlsxPath, string sheetName);
    
    // Read row heights from worksheet  
    double[] GetRowHeights(string xlsxPath, string sheetName);
    
    // Read print area
    (int firstCol, int lastCol, int firstRow, int lastRow) GetPrintArea(string xlsxPath, string sheetName);
    
    // Read PageSetup
    PageSetupInfo GetPageSetup(string xlsxPath, string sheetName);
    
    // Read merged cells with content
    List<MergeAreaInfo> GetMergeAreas(string xlsxPath, string sheetName);
}

public class PageSetupInfo
{
    public double LeftMargin { get; set; }    // points (default 54)
    public double RightMargin { get; set; }   // points (default 54)
    public double TopMargin { get; set; }     // points (default 54)
    public double BottomMargin { get; set; }  // points (default 54)
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
}
```

### Phase 23b: Coordinate Calculator

```csharp
public class ClusterSizeCalculator
{
    List<ClusterRect> Calculate(string xlsxPath, string pdfPath, int pageNumber)
    {
        // 1. Read Excel data via OpenXML
        var columnWidths = GetColumnWidths(...);
        var rowHeights = GetRowHeights(...);
        var pageSetup = GetPageSetup(...);
        var mergeAreas = GetMergeAreas(...);
        
        // 2. Get page dimensions from PDF
        double pageWidth = GetPdfPageWidth(pdfPath);
        double pageHeight = GetPdfPageHeight(pdfPath);
        
        // 3. Compute content dimensions for centering
        double contentWidth = Sum(columnWidths[printArea.ColRange]);
        double contentHeight = Sum(rowHeights[printArea.RowRange]);
        
        // 4. Compute centering offsets
        double centerX = pageSetup.CenterHorizontally
            ? Math.Max(0, (pageWidth - pageSetup.LeftMargin - pageSetup.RightMargin - contentWidth) / 2)
            : 0;
        double centerY = pageSetup.CenterVertically
            ? Math.Max(0, (pageHeight - pageSetup.TopMargin - pageSetup.BottomMargin - contentHeight) / 2)
            : 0;
        
        // 5. For each merged cell, compute position
        var results = new List<ClusterRect>();
        foreach (var area in mergeAreas)
        {
            double left  = pageSetup.LeftMargin + centerX + Sum(columnWidths[1..area.FirstCol-1]);
            double right = pageSetup.LeftMargin + centerX + Sum(columnWidths[1..area.LastCol]);
            double top   = pageSetup.TopMargin + centerY + Sum(rowHeights[1..area.FirstRow-1]);
            double bottom = pageSetup.TopMargin + centerY + Sum(rowHeights[1..area.LastRow]);
            
            results.Add(new ClusterRect
            {
                Left   = (float)(left / pageWidth),
                Right  = (float)(right / pageWidth),
                Top    = (float)(top / pageHeight),
                Bottom = (float)(bottom / pageHeight)
            });
        }
        
        // 6. Round to 7 decimal places (matching DB storage)
        results = RoundCoordinates(results, 7);
        
        return results;
    }
}
```

### Phase 23c: Calibration (Critical!)

The mathematical approach WILL have small differences from pixel-scanning due to:
1. Column width conversion (OpenXML chars → points)
2. Margin interpretation (Excel internal padding)
3. Border widths (~1-2 pts per cell border)

**Calibration process**:
1. Run both legacy (pixel-scan) and new (mathematical) on same template
2. Compare coordinates point by point
3. Compute offset = legacy_value - computed_value
4. Apply as correction factor if consistent

**Expected difference**: < 0.5% ratio (≈ 3 pts at 612 page width)

---

## 4. ASP.NET Core Implementation Details

### Dependencies

```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.1.1" />
<PackageReference Include="PdfPig" Version="0.1.8" />  <!-- For PDF page dimensions -->
```

### ClusterRect Model

```csharp
public class ClusterRect : IComparable<ClusterRect>
{
    public float Top { get; set; }
    public float Bottom { get; set; }
    public float Left { get; set; }
    public float Right { get; set; }
    
    public int CompareTo(ClusterRect other) => Left.CompareTo(other.Left);
}
```

---

## 5. Verification Checklist

| Test | Expected | How to Verify |
|------|----------|---------------|
| Template 546 | Match stored cluster JSON | Compare all 6 clusters |
| Template 547 | Match stored values | Compare all clusters |
| Template 548 | Match stored values | Compare all clusters |
| Custom template | Reasonable positions | Visual inspection |
| Landscape page | Correct aspect ratio | Ratio * pageDim = expected |
| Non-centered template | Content flush with margin | Compare with pixel-scan |
| Different margins | Correct offset | Compare with pixel-scan |

---

## 6. Fallback: Full Pixel-Scanning

If mathematical approach differs by >1%, implement full pixel-scanning:

```csharp
// Using only cross-platform libraries
1. ClosedXML → Export to PDF (SaveAs)
2. PdfPig → Render page to raw bytes
3. ImageSharp → Manual morphological close (3×3 kernel)
4. Pure C# → Pixel scan (same algorithm as GetAddress)
```

This is more complex but guaranteed 100% identical output.
