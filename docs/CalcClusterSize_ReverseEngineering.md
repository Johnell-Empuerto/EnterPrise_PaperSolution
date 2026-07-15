# CalcClusterSize — Complete Coordinate Generation Algorithm

> **Source:** Decompiled IL from `LibExcelController.dll` (ConMas Designer) + `iReporterExcelAddIn.dll`  
> **Date:** 2026-07-15  
> **Status:** Complete

---

## 1. The Full Upload/Publish Pipeline

```
User initiates "Publish/Upload"
     │
     ▼
┌─────────────────────────────────────────────────────┐
│ [Phase A] MakeCluster(worksheet, sheetNo, oSheetObj)│
│   (iReporterExcelAddIn.dll / LibExcelController)    │
│                                                     │
│   1. Read comments from every cell in print area    │
│   2. For cells WITH a comment:                      │
│      - Set interior to BLACK (Color.Black)          │
│      - Clear cell value to empty                    │
│   3. For cells WITHOUT a comment:                   │
│      - Set interior to WHITE (Color.White)          │
│      - Clear cell value to empty                    │
│   4. Clear ALL borders, shapes, headers/footers     │
│   5. Build XElement with cluster metadata           │
│      (cellAddr, clusterType from comment 1st line)  │
│   6. Save sanitized workbook to temp.xlsx           │
│                                                     │
│   Returns: XElement <cluster> list (no coordinates) │
└───────────────────┬─────────────────────────────────┘
                    │
                    ▼
┌─────────────────────────────────────────────────────┐
│ [Phase B] CalcClusterSize(xConmas, workbook)         │
│   (LibExcelController.ExcelControllerInterop)        │
│                                                     │
│   1. ExportPdf(temp.xlsx → temp.pdf)                 │
│      via Excel COM ExportAsFixedFormat(xlTypePDF)    │
│                                                     │
│   2. CreateImageFromPdf(temp.pdf, page=0,            │
│                          applyMorphology=false)       │
│      → System.Drawing.Image at DpiPdfToImage (300)   │
│                                                     │
│   3. GetClusterSize(temp.pdf, page=0, clusterCount)  │
│      → GetAddress(image, clusterCount)               │
│        (pixel scanning → List<ClusterRect>)          │
│                                                     │
│   4. SortClusters(clusterRects, topThresh, leftThresh)│
│                                                     │
│   5. For each cluster element in XElement:           │
│      left_position   = rect.Left   / imageWidth      │
│      top_position    = rect.Top    / imageHeight     │
│      right_position  = rect.Right  / imageWidth      │
│      bottom_position = rect.Bottom / imageHeight     │
│      (format: 7 decimal places, e.g. "0.3364706")    │
│                                                     │
│   Writes ratios to XML attributes                    │
└───────────────────┬─────────────────────────────────┘
                    │
                    ▼
┌─────────────────────────────────────────────────────┐
│ [Phase C] ExportPdf(clean-copy.xlsx → output.pdf)    │
│   (for background_image_file storage)                │
└───────────────────┬─────────────────────────────────┘
                    │
                    ▼
┌─────────────────────────────────────────────────────┐
│ [Phase D] Database writes (single transaction)       │
│   INSERT def_cluster (with ratios)                   │
│   INSERT def_sheet                                   │
│   UPDATE def_top SET xml_data, background_image_file │
└─────────────────────────────────────────────────────┘
```

---

## 2. Complete Call Graph (IL-Verified)

```
ExcelControllerBase.ExportExcelDefinition(xlsPath, outputDir)
  │
  ├── new ExcelControllerInterop() / or this as base
  │
  └── ExcelControllerInterop.GetDefinition(progress, outputDir, xlsFileName)
        │
        ├── For each worksheet in workbook.Sheets:
        │     │
        │     └── MakeCluster(worksheet, sheetNo, oSheetObj)  ← 5354 bytes IL
        │           │  49 locals. Largest method in DLL.
        │           │
        │           ├── worksheet.ClearShapes()
        │           │
        │           ├── usedRange = worksheet.UsedRange()
        │           │
        │           ├── For row = usedRange.FirstRow → LastRow:
        │           │   For col = usedRange.FirstColumn → LastColumn:
        │           │     │
        │           │     ├── cell = worksheet.Cells(row, col)
        │           │     ├── IF cell.HasComment()  ← Infragistics API
        │           │     │     comment = cell.GetComment()
        │           │     │     commentText = comment.Text
        │           │     │     fieldType = commentText.Split()[0]  // first line
        │           │     │     mergeArea = cell.MergeArea ?? cell
        │           │     │
        │           │     │     // FILL with BLACK for PDF detection
        │           │     │     mergeArea.Interior.Color = Color.Black
        │           │     │     mergeArea.Value = ""  // clear content
        │           │     │
        │           │     └── ELSE:
        │           │           cell.Interior.Color = Color.White  // clear to white
        │           │           cell.Value = ""  // clear content
        │           │
        │           ├── // Remove ALL borders
        │           │   For each cell in usedRange:
        │           │     cell.Borders.LineStyle = xlNone
        │           │
        │           ├── // Clear shapes, headers, footers
        │           │   worksheet.PageSetup.CenterHeader = ""
        │           │   worksheet.PageSetup.CenterFooter = ""
        │           │
        │           └── Return XElement:
        │                 <sheet>
        │                   <sheetNo>{n}</sheetNo>
        │                   <cluster>
        │                     <clusterId>{n}</clusterId>
        │                     <name>{commentLine0}</name>
        │                     <type>{commentLine1}</type>
        │                     <cellAddress>$A$1:$B$2</cellAddress>
        │                     <inputParameters>{...}</inputParameters>
        │                   </cluster>
        │                   ...
        │                 </sheet>
        │
        ├── workbook.SaveAs(temp.xlsx)  ← Save sanitized workbook
        │
        └── CalcClusterSize(xConmas, workbook, ref progress)  ← 776 bytes IL
              13 locals.
              │
              ├── string pdfPath = Path.Combine(
              │       Path.GetTempPath(), $"ConMas_{Guid.NewGuid()}.pdf")
              │
              ├── ImageUtility.ExportPdf(temp.xlsx, pdfPath)
              │   │  ← 350 bytes IL, 3 locals
              │   │
              │   └── Select version:
              │         ├── If Excel >= 2010: ExportPdf2010()
              │         │     excel.Workbooks.Open(sourceXlsPath)
              │         │     workbook.ExportAsFixedFormat(
              │         │       Type: xlTypePDF (=0),
              │         │       Filename: destPdfPath,
              │         │       Quality: xlQualityStandard (=0),
              │         │       IncludeDocProps: Missing,
              │         │       IgnorePrintAreas: Missing (=false),
              │         │       From: Missing, To: Missing,
              │         │       OpenAfterPublish: Missing)
              │         │     workbook.Close(false)
              │         │
              │         └── If Excel 2007: ExportPdf2007()
              │               (same, different params)
              │
              ├── Image pdfImage = CreateImageFromPdf(pdfPath, 0, false)
              │   │
              │   │  Renders PDF page 0 at DpiPdfToImage (300 DPI)
              │   │  Returns System.Drawing.Image
              │   │  imageWidth  = pdfImage.Width   ← SOURCE OF TRUTH
              │   │  imageHeight = pdfImage.Height  ← SOURCE OF TRUTH
              │   │
              │   └── When applyMorphology=true (AUTO-DETECT only):
              │         OpenCV: BGRA→GRAY + MorphologyEx(Close, 3×3 rect kernel)
              │
              ├── List<ClusterRect> rects = GetClusterSize(pdfPath, 0, clusterCount)
              │   │  ← 71 bytes IL bridge, 3 locals
              │   │
              │   └── GetClusterSize(strPdfPath, page, clusterCount):
              │         │
              │         ├── Image img = CreateImageFromPdf(strPdfPath, page, false)
              │         │   // 300 DPI, NO morphology
              │         │
              │         └── return GetAddress(img, clusterCount)
              │               ← 754 bytes IL, 16 locals
              │
              ├── rects = SortClusters(rects, topThreshold, leftThreshold)
              │   ← 534 bytes IL, 10 locals
              │
              └── For each XElement cluster in xConmas.Descendants("cluster"):
                    │
                    ├── Match cluster to rect by sorted order
                    │
                    ├── Write normalized ratios:
                    │     cluster.SetAttributeValue("left_position",
                    │       (rect.Left / (float)imageWidth).ToString("0.0#######"))
                    │     cluster.SetAttributeValue("top_position",
                    │       (rect.Top / (float)imageHeight).ToString("0.0#######"))
                    │     cluster.SetAttributeValue("right_position",
                    │       (rect.Right / (float)imageWidth).ToString("0.0#######"))
                    │     cluster.SetAttributeValue("bottom_position",
                    │       (rect.Bottom / (float)imageHeight).ToString("0.0#######"))
                    │
                    └── // rect.Left/Top/Right/Bottom are PIXEL coordinates
                        // from the 300 DPI rasterized image
```

---

## 3. `GetAddress` — The Core Pixel-Scanning Algorithm

**Source:** `LibExcelController.Lib.ImageUtility.GetAddress` (754 bytes IL, 16 locals)

This is the image processing core that detects black rectangles on a white background (from the sanitized workbook export).

### 3.1 Input Image

The input is a **grayscale bitmap** of the sanitized workbook PDF page:
- **Black pixels** = cluster fields (cells that had comments, filled with black in MakeCluster)
- **White pixels** = everything else (cells without comments, cleared to white)
- **No borders, no text, no shapes** — all removed in MakeCluster
- **300 DPI** (from `DpiPdfToImage` property)
- **No morphology applied** (`applyMorphology: false`)

### 3.2 Scanning Algorithm (IL-Reconstructed)

```
// Pixel scan: top→bottom, left→right
for y = 0 to image.Height - 1:
    for x = 0 to image.Width - 1:
        pixel = image.GetPixel(x, y)

        // Detect TOP-LEFT CORNER of a black rectangle
        if pixel IS BLACK
           AND pixel(x-1, y) IS WHITE     // left neighbor is white
           AND pixel(x, y-1) IS WHITE     // top neighbor is white
        then
            // This is the TOP-LEFT corner of a cluster

            // --- 6-PIXEL VERIFICATION (noise filter) ---
            // Check 6 pixels right of (x, y-1) are all WHITE
            // Check 6 pixels left of (x+6, y) are all WHITE
            // Check 6 pixels above (x, y+6) are all WHITE
            // If any check fails → skip (noise)

            // --- EXPAND RIGHT to find right edge ---
            scanX = x + 1
            while scanX < image.Width AND pixel(scanX, y) IS BLACK:
                scanX++
            // scanX is now first WHITE pixel

            // --- EXPAND DOWN to find bottom edge ---
            // Walk each column from x to scanX, downward
            maxY = y
            for col = x to scanX:
                scanY = y + 1
                while scanY < image.Height AND pixel(col, scanY) IS BLACK:
                    scanY++
                maxY = max(maxY, scanY - 1)

            // --- MINIMUM SIZE FILTER ---
            if (scanX - 1 - x) >= 6 AND (maxY - y) >= 6:
                rect = new ClusterRect()
                rect.Left   = (float)x
                rect.Top    = (float)y
                rect.Right  = (float)(scanX - 1)
                rect.Bottom = (float)maxY
                Add rect to result list

            // Mark scanned pixels (skip to avoid re-detection)

// Sort results: top-to-bottom, left-to-right (within topThreshold)
return SortClusters(result, topThreshold, leftThreshold)
```

### 3.3 Key Properties

| Property | Value | Evidence |
|----------|-------|----------|
| Scan direction | Top→bottom, left→right | IL: 2 nested for loops |
| Content detection | BLACK pixel with WHITE neighbors above/left | IL: pixel comparison ops |
| Right edge | First WHITE pixel moving right from content | IL: while loop scanning right |
| Bottom edge | Max Y of black pixels in each column of the content | IL: inner loop per column |
| Minimum size | ≥6 pixels in both dimensions | IL: comparison with constant 6 |
| Noise filter | 6-pixel verification of surrounding white | IL: multiple GetPixel calls |
| Morphology | NONE (applyMorphology=false in manual path) | IL: parameter passed as false |
| Post-processing | SortClusters (row grouping + column sort) | IL: after main loop |

---

## 4. Exact Answers to Each Question

### Q1: What is the starting coordinate?

**Not from COM.** The starting coordinate is the **pixel position** of the top-left corner of a black rectangle detected by `GetAddress()`:

```
rect.Left = x_pixel    // first column where BLACK pixel found
rect.Top  = y_pixel    // first row where BLACK pixel found
```

Where `(x_pixel, y_pixel)` is the first pixel detected as: `pixel(x,y)==BLACK AND pixel(x-1,y)==WHITE AND pixel(x,y-1)==WHITE`.

**This is NOT `Range.Left` or `MergeArea.Left`.** COM properties are NEVER used as coordinates. The pixel scanning of the rendered PDF is the only source.

MakeCluster's COM range reading (`MergeArea.Row`, `MergeArea.Column`) is only used for:
- Identifying which cells are clusters (comment-based)
- Determining the MergeArea to fill with black
- Building the cell address string (`$A$1:$B$2`)

**IL evidence:** `GetAddress` contains no calls to `Range.Left`, `Range.Top`, `MergeArea.Left`, or any COM geometry API. It only calls `image.GetPixel(x, y)`.

### Q2: How is width calculated?

By **pixel expansion rightward** from the top-left corner:

```
scanX = x + 1
while scanX < imageWidth AND IsBlack(pixel(scanX, y)):
    scanX++

rect.Right = (float)(scanX - 1)
width_px = rect.Right - rect.Left
```

Width is NOT `Range.Width`, `MergeArea.Width`, or sum of column widths. It's the pixel extent of the contiguous black region on the rendered image.

### Q3: How is height calculated?

By **per-column downward expansion** from the top-left corner:

```
maxY = y
for col = x to scanX:    // each column from left to right edge
    scanY = y + 1
    while scanY < imageHeight AND IsBlack(pixel(col, scanY)):
        scanY++
    maxY = max(maxY, scanY - 1)

rect.Bottom = (float)maxY
height_px = rect.Bottom - rect.Top
```

The algorithm finds the tallest black column within the detected region. Each column is scanned independently downward, and the maximum extent across all columns is used as the bottom edge.

### Q4: What image dimensions are used as denominator?

The **image dimensions** (pixel width and height) of the `System.Drawing.Image` returned by `CreateImageFromPdf()`:

```csharp
Image pdfImage = CreateImageFromPdf(pdfPath, 0, false);
int imageWidth  = pdfImage.Width;    // ← denominator for left/right
int imageHeight = pdfImage.Height;   // ← denominator for top/bottom
```

These are the **full page dimensions** at 300 DPI. For Letter (612×792 pt at 300 DPI):
- `imageWidth = 612 × 300/72 = 2550 px`
- `imageHeight = 792 × 300/72 = 3300 px`

**These are NOT:**
- PrintArea dimensions
- UsedRange dimensions
- Paper size in points (the raw pt value is converted through DPI)
- Composite dimensions from column/row sums

**IL evidence:** `CalcClusterSize` declares locals `[7] int imageWidth` and `[8] int imageHeight` and assigns them from `pdfImage.Width` / `pdfImage.Height`.

### Q5: How is the PDF size obtained?

From **`CreateImageFromPdf()`** which renders the PDF at `DpiPdfToImage` (300 DPI):

```
pdfImage = CreateImageFromPdf(pdfPath, pageNumber=0, applyMorphology=false)
```

This renders the PDF page at **300 DPI** using `System.Drawing.Graphics`. The resulting bitmap dimensions are the denominator values (`pdfImage.Width`, `pdfImage.Height`).

The PDF itself is generated by `ExportAsFixedFormat(xlTypePDF)` via Excel COM on the **sanitized** workbook (where clusters are black rectangles). However, the PDF page size is identical to the original workbook's page setup (same paper size, margins, etc.).

**Important:** The `background_image_file` stored in the database is a PDF of the **CLEAN** (unsanitized) workbook, not the sanitized one. But both have the same page dimensions, so the ratios computed from the sanitized PDF apply correctly to the clean PDF background.

### Q6: Is there any transformation before normalization?

**No transformations are applied to the pixel coordinates.** The `ClusterRect` values (Left, Top, Right, Bottom) are used directly in the ratio formula.

Specifically, the following are **NOT** applied:
- ❌ Margin subtraction — PDF includes full page including margins
- ❌ Printer offset compensation — Excel COM ExportAsFixedFormat handles this internally
- ❌ Zoom correction — The PDF export uses whatever zoom/scaling the workbook defines
- ❌ FitToPage adjustment — Same, respected by ExportAsFixedFormat
- ❌ Centering compensation — The PDF renders with PageSetup centering applied
- ❌ Gap correction — No post-hoc adjustments
- ❌ Merged cell adjustment — Merge areas are handled in MakeCluster (black fill applied to the entire MergeArea)

**However**, three things DO affect the detected coordinates indirectly (they are part of the rendering, not post-hoc adjustments):

1. **Black fill during MakeCluster**: The cluster's MergeArea is filled solid black. This means the detected rectangle covers the ENTIRE merge area, not individual cells within it.

2. **6-pixel noise filter**: Very small content (<6 px) is discarded. This sets a minimum rectangle size.

3. **Per-column bottom-edge detection**: The bottom edge is the MAXIMUM Y of black pixels across all columns in the rectangle. If one column is taller than others, the rectangle extends to cover it.

**IL evidence:** `CalcClusterSize` contains no arithmetic operations on pixel coordinates before the division — the values go directly from `ClusterRect.Left/Top/Right/Bottom` into the ratio computation.

### Q7: What is the exact normalization formula?

**Confirmed from IL (776 bytes, 13 locals):**

```
left_position   = (float)rect.Left   / (float)imageWidth
top_position    = (float)rect.Top    / (float)imageHeight
right_position  = (float)rect.Right  / (float)imageWidth
bottom_position = (float)rect.Bottom / (float)imageHeight
```

Where:
- `rect.Left/Top/Right/Bottom` = pixel coordinates from `GetAddress()` (type: `float` in `ClusterRect`)
- `imageWidth/imageHeight` = pixel dimensions of the PDF rendered at **300 DPI** from `CreateImageFromPdf()`

**No constants, no offsets, no adjustments.** The formula is literally `ratio = pixel_coord / pixel_dimension`.

The ratios are then formatted to 7 decimal places using `ToString("0.0#######")` (or `"F7"` — format may vary slightly but produces values like `"0.3364706"`).

### Q8: Does the algorithm depend on column widths, row heights, print area, page margins, merged cells, or only on rendered pixel coordinates?

**Only on rendered pixel coordinates.** The coordinate detection in `GetAddress` is entirely image-based.

However, the **MakeCluster sanitization** phase (which runs BEFORE CalcClusterSize) does depend on:

- **Merged cells**: The `MergeArea` of a commented cell determines the extent of the black fill. If a cell belongs to a merge range, the entire merge area is filled black.
- **Print area**: Only cells in the print area are scanned for comments (per `GetFirstPagePrintArea`).
- **Column widths & row heights**: These affect the rendered PDF dimensions indirectly (Excel COM renders at the workbook's page setup, which respects column/row geometry). The pixel scan then measures the RENDERED result, not the raw geometry.

**The coordinate calculation itself** (`GetAddress` → ratio) depends ONLY on:
- Pixel positions of black regions in the rendered image
- Image dimensions (width, height) at 300 DPI

**Evidence:** `GetAddress` IL contains no calls to any COM range property, no worksheet geometry API, and no OpenXML parsing. It only calls `Bitmap.GetPixel(x, y)`.

### Q9: Complete Call Chain (definitive)

```
Publish/Upload Button Click
  │
  └── iReporterExcelAddIn.ThisAddIn.ExportExcelDefinition(xlsPath, outputDir)
        │
        ├── LibExcelController.ExcelControllerBase.GetDefinition(progress, outputDir, xlsFileName)
        │     │  [4036 bytes IL, 51 locals]
        │     │
        │     ├── Load workbook (Infragistics primary, COM fallback)
        │     │
        │     ├── For each worksheet in workbook.Sheets:
        │     │     │
        │     │     └── MakeCluster(worksheet, sheetNo, oSheetObj)
        │     │           │  [5354 bytes IL, 49 locals]
        │     │           │
        │     │           ├── For each cell in print area:
        │     │           │     IF cell.HasComment() (Infragistics API):
        │     │           │       mergeArea = cell.MergeArea ?? cell
        │     │           │       mergeArea.Interior.Color = BLACK
        │     │           │       mergeArea.Value = ""
        │     │           │     ELSE:
        │     │           │       cell.Interior.Color = WHITE
        │     │           │       cell.Value = ""
        │     │           │
        │     │           ├── Clear all borders, shapes, headers/footers
        │     │           │
        │     │           └── Build XElement (cellAddr, type, name)
        │     │
        │     ├── workbook.SaveAs(temp.xlsx)
        │     │
        │     └── CalcClusterSize(xConmas, workbook, ref progress)
        │           │  [776 bytes IL, 13 locals]
        │           │
        │           ├── ImageUtility.ExportPdf(temp.xlsx, temp.pdf)
        │           │     │  [350 bytes IL, 3 locals]
        │           │     │
        │           │     └── excel.Workbooks.Open(temp.xlsx)
        │           │           workbook.ExportAsFixedFormat(xlTypePDF, temp.pdf,
        │           │             xlQualityStandard, IgnorePrintAreas=false)
        │           │           workbook.Close(false)
        │           │
        │           ├── Image pdfImage = CreateImageFromPdf(temp.pdf, 0, false)
        │           │     Render PDF page 0 at DpiPdfToImage (300 DPI)
        │           │     NO morphology (applyMorphology=false)
        │           │     imageWidth = pdfImage.Width
        │           │     imageHeight = pdfImage.Height
        │           │
        │           ├── List<ClusterRect> rects = GetClusterSize(temp.pdf, 0, clusterCount)
        │           │     │  [71 bytes IL, 3 locals]
        │           │     │
        │           │     └── Image img = CreateImageFromPdf(pdfPath, 0, false)
        │           │           GetAddress(img, clusterCount)
        │           │             │  [754 bytes IL, 16 locals]
        │           │             │
        │           │             └── For y = 0 to img.Height:
        │           │                   For x = 0 to img.Width:
        │           │                     IF pixel(x,y)==BLACK
        │           │                        AND pixel(x-1,y)==WHITE
        │           │                        AND pixel(x,y-1)==WHITE
        │           │                     THEN:
        │           │                       Expand right → rightX
        │           │                       Expand down → bottomY
        │           │                       IF width≥6 AND height≥6:
        │           │                         rect.Left = x
        │           │                         rect.Top = y
        │           │                         rect.Right = rightX
        │           │                         rect.Bottom = bottomY
        │           │                         Add rect to list
        │           │
        │           ├── rects = SortClusters(rects, topThreshold, leftThreshold)
        │           │     [534 bytes IL, 10 locals]
        │           │
        │           └── For each cluster in xConmas XML:
        │                 left_position   = rect.Left   / imageWidth
        │                 top_position    = rect.Top    / imageHeight
        │                 right_position  = rect.Right  / imageWidth
        │                 bottom_position = rect.Bottom / imageHeight
        │
        ├── ExportPdf(clean.xlsx, output.pdf)  ← for DB storage
        │
        └── Database transaction
              ├── INSERT def_cluster (ratios)
              ├── INSERT def_sheet
              ├── UPDATE def_top SET xml_data
              └── UPDATE def_top SET background_image_file = PDF bytes
```

### Q10: Reconstructed Pseudocode

```python
def publish_workbook(xlsx_path: str) -> dict:
    """
    Reconstructed ConMas Upload/Publish algorithm.
    
    Returns cluster data with left_position, top_position,
    right_position, bottom_position normalized ratios.
    """
    import win32com.client
    from PIL import Image
    
    DPI = 300
    
    # ===== Phase 1: Sanitize workbook (MakeCluster) =====
    
    # Open workbook via Excel COM
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(xlsx_path)
    
    clusters_xml = []  # list of dicts with cellAddr, type, name
    
    for ws in wb.Worksheets:
        used = ws.UsedRange
        print_area = ws.PageSetup.PrintArea  # or compute from UsedRange
        
        for row in range(used.Row, used.Row + used.Rows.Count):
            for col in range(used.Column, used.Column + used.Columns.Count):
                cell = ws.Cells(row, col)
                
                # Check for cell comment
                comment = None
                try:
                    comment = cell.Comment  # COM per-cell property
                except:
                    pass
                
                if comment is not None:
                    # This cell has a ConMas field comment
                    merge_area = cell.MergeArea if cell.MergeArea else cell
                    
                    # Record cluster metadata
                    lines = str(comment.Text()).replace("\r\n", "\n").split("\n")
                    cluster = {
                        "cellAddr": merge_area.Address(False, False),
                        "name": lines[0].strip() if lines else "",
                        "type": lines[1].strip() if len(lines) > 1 else "Text",
                    }
                    clusters_xml.append(cluster)
                    
                    # Fill merge area with BLACK
                    merge_area.Interior.Color = 0x000000  # Black
                    merge_area.Value = ""
                else:
                    # Clear to WHITE
                    cell.Interior.Color = 0xFFFFFF  # White
                    cell.Value = ""
        
        # Remove borders, shapes, headers/footers
        ws.PageSetup.CenterHeader = ""
        ws.PageSetup.CenterFooter = ""
        # Clear shapes: for each shape in ws.Shapes: shape.Delete()
    
    # ===== Phase 2: Export sanitized workbook to PDF =====
    
    import tempfile, os
    temp_dir = tempfile.mkdtemp(prefix="conmas_")
    sanitized_xlsx = os.path.join(temp_dir, "sanitized.xlsx")
    wb.SaveAs(sanitized_xlsx)
    
    # Export to PDF via Excel COM (same as original)
    pdf_path = os.path.join(temp_dir, "output.pdf")
    wb2 = excel.Workbooks.Open(sanitized_xlsx)
    wb2.ExportAsFixedFormat(0, pdf_path)  # xlTypePDF=0
    wb2.Close(False)
    wb.Close(False)
    excel.Quit()
    
    # ===== Phase 3: Render PDF to image at 300 DPI =====
    
    import fitz  # PyMuPDF
    doc = fitz.open(pdf_path)
    page = doc[0]
    zoom = DPI / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    image_width = pix.width    # e.g., 2550 for Letter
    image_height = pix.height  # e.g., 3300 for Letter
    doc.close()
    
    # Convert to numpy array for pixel scanning
    import numpy as np
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    
    # Threshold: pixel is "BLACK" if all RGB channels < threshold
    def is_black(r, g, b, threshold=50):
        return r < threshold and g < threshold and b < threshold
    
    def is_white(r, g, b, threshold=200):
        return r > threshold and g > threshold and b > threshold
    
    # ===== Phase 4: Pixel scan for black rectangles (GetAddress) =====
    
    MIN_SIZE = 6  # minimum 6 pixels in each dimension
    
    rects = []
    visited = np.zeros((image_height, image_width), dtype=bool)
    
    for y in range(1, image_height - 1):
        for x in range(1, image_width - 1):
            if visited[y, x]:
                continue
            
            r, g, b = img[y, x]
            
            # Detect top-left corner: BLACK pixel with WHITE above and left
            if not is_black(r, g, b):
                continue
            
            r_left, g_left, b_left = img[y, x - 1]
            r_up, g_up, b_up = img[y - 1, x]
            
            if not (is_white(r_left, g_left, b_left) and is_white(r_up, g_up, b_up)):
                continue
            
            # 6-pixel noise filter: verify clean white along top and left edges
            # (IL: checks 6 pixels to the right of (x, y-1), 6 pixels above (x, y+6), etc.)
            clean = True
            for dx in range(1, 7):
                if x + dx < image_width:
                    rr, gg, bb = img[y - 1 if y > 0 else 0, x + dx]
                    if not is_white(rr, gg, bb):
                        clean = False
                        break
            if not clean:
                continue
            
            # Expand RIGHT: scan from (x+1, y) rightward until WHITE
            right_x = x + 1
            while right_x < image_width:
                rr, gg, bb = img[y, right_x]
                if is_white(rr, gg, bb):
                    break
                right_x += 1
            right_x -= 1  # back to last BLACK pixel
            
            # Expand DOWN: find tallest black column
            max_bottom = y
            for col in range(x, right_x + 1):
                scan_y = y + 1
                while scan_y < image_height:
                    rr, gg, bb = img[scan_y, col]
                    if is_white(rr, gg, bb):
                        break
                    scan_y += 1
                max_bottom = max(max_bottom, scan_y - 1)
            
            # Minimum size filter
            rect_width = right_x - x + 1
            rect_height = max_bottom - y + 1
            if rect_width >= MIN_SIZE and rect_height >= MIN_SIZE:
                rects.append({
                    "Left": float(x),
                    "Top": float(y),
                    "Right": float(right_x),
                    "Bottom": float(max_bottom),
                })
                
                # Mark pixels as visited to avoid re-detection
                visited[y:max_bottom + 1, x:right_x + 1] = True
    
    # ===== Phase 5: Sort clusters =====
    # Sort: by Top (grouped within topThreshold), then by Left
    rects.sort(key=lambda r: (r["Top"] // 5, r["Left"]))
    
    # ===== Phase 6: Normalize to ratios =====
    
    for i, rect in enumerate(rects):
        rect["left_position"]   = rect["Left"]   / image_width
        rect["top_position"]    = rect["Top"]    / image_height
        rect["right_position"]  = rect["Right"]  / image_width
        rect["bottom_position"] = rect["Bottom"] / image_height
    
    # ===== Phase 7: Match clusters to XML elements =====
    # (Both lists are sorted the same way, so index-based matching works)
    for cluster, rect in zip(clusters_xml, rects):
        cluster["left_position"]   = f"{rect['left_position']:.7f}"
        cluster["top_position"]    = f"{rect['top_position']:.7f}"
        cluster["right_position"]  = f"{rect['right_position']:.7f}"
        cluster["bottom_position"] = f"{rect['bottom_position']:.7f}"
    
    # Cleanup
    import shutil
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    return clusters_xml
```

---

## 5. Key Data Structures

### ClusterRect (from `LibExcelController.Lib.ClusterRect`)

```csharp
// ValueType (struct)
class ClusterRect {
    float Left   { get; set; }   // pixel X of left edge (0-based)
    float Top    { get; set; }   // pixel Y of top edge (0-based)
    float Right  { get; set; }   // pixel X of right edge (inclusive)
    float Bottom { get; set; }   // pixel Y of bottom edge (inclusive)
}
```

All values are **pixel coordinates on the 300 DPI rasterized image** of the sanitized PDF.

### CalculateCell (from `LibExcelController.CalculateCell`)

```csharp
class CalculateCell {
    string CellAdress { get; set; }  // e.g., "Sheet1!$A$1:$B$2"
    int SheetNo { get; set; }
    int Row { get; set; }
    int Col { get; set; }
    string CommentText { get; set; }
    string FieldType { get; set; }
    object Value { get; set; }
    bool HasComment { get; set; }
}
```

Used internally by `MakeCluster` as intermediate data holder. Never stored to DB.

---

## 6. DPI Confusion Resolved

| Document | Claimed DPI | Evidence Level |
|----------|-------------|----------------|
| `Phase32_LegacyPublishPipeline.md` | **200 DPI** — "DpiPdfToImage = 200; // Hardcoded" | ❌ Contradicted by empirical ratio math |
| `HowConMasLegacyWorks.md` | **300 DPI** — "DpiPdfToImage { get; set; } // = 300 (default)" | ✅ Matches stored ratio × dimension |
| `final_validation.py` (Experiment 6) | **300 DPI** — Ratio × 300 DPI = 0px delta | ✅ Empirical proof |

**Definitive answer: 300 DPI.**

Proof: Stored ratio `0.3364706 × 2550 = 858 px` (300 DPI Letter width). Our render at 300 DPI produces zero delta for Template 546. If DPI were 200, the image width would be 1700 and `0.3364706 × 1700 = 572 px` — which does NOT match the correct field position.

The 200 DPI claim in Phase32 appears to be an error — possibly from a different DLL version, the tablet code path, or a misinterpretation of a local variable initialization that gets overridden by the base class property.

---

## 7. What MakeCluster Does NOT Do

- ❌ Does NOT set `left_position`, `top_position`, `right_position`, `bottom_position`
- ❌ Does NOT read `Range.Left`/`Range.Top`/`Range.Width`/`Range.Height` for coordinate purposes
- ❌ Does NOT compute pixel coordinates from column widths / row heights
- ❌ Does NOT read the `_Fields` sheet
- ❌ Does NOT matter what DPI or page size is used — the coordinate output is always ratios that scale with any DPI/page size

MakeCluster only:
- Reads cell comments (field type, name, parameters)
- Fills cluster cells with BLACK (for PDF detection)
- Clears everything else to WHITE
- Removes borders/shapes/headers
- Creates XML structure with cell addresses and field metadata

The actual coordinate computation happens ENTIRELY in `CalcClusterSize` → `GetClusterSize` → `GetAddress`.

---

## 8. Critical Implications

1. **Coordinates come from pixel scanning, not COM geometry.** Any attempt to compute coordinates from `Range.Left`/`Range.Width`/columns/rows will produce different results.

2. **The sanitization step is essential.** Without it, the PDF would contain borders, text, and colors that would confuse pixel scanning. The solid BLACK fill of cluster cells is what makes the rectangles detectable.

3. **The formula is trivial.** `ratio = pixel / image_dimension`. All the complexity is in MakeCluster (sanitization) and GetAddress (pixel scanning), not in the normalization math.

4. **DPI independence.** The ratios work at ANY DPI because both numerator and denominator scale with DPI. `(x * 300/72) / (pageWidth * 300/72) = x / pageWidth` at point level.

5. **Match between GetClusterSize PDF and background_image_file PDF.** Both use `ExportAsFixedFormat` on the same workbook. The GETCLUSTER PDF is the sanitized version (black rectangles), the background_image_file PDF is the original workbook. Both have the same page dimensions, so ratios computed from one apply correctly to the other.

6. **No calibration constants needed.** The ratio formula has no hidden offsets, no gap corrections, no margin compensations. The PDF render (through `ExportAsFixedFormat`) naturally includes all page setup transformations (margins, centering, printer compensation).
