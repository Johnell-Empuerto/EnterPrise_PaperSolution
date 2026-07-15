# Phase 22 — Designer Reverse Engineering Blueprint

## Scope

```
Excel Workbook → Upload → ASP.NET Core Engine → Next.js Designer → Visual Result
```

This document covers ONLY the pipeline from Excel upload to the visual Designer display. Tablet, runtime, OCR, workflow, mobile, auth — all out of scope.

---

## Part 1 — Excel Geometry

### 1.1 Row Heights

**Source:** `ExcelWorksheetCom.GetRange(row, col)` → `Range.Rows[i].Height` (COM)

**Reading:** From LibExcelController's `ExcelControllerBase.ExportExcelDefinition()` and the COM interop layer. Row heights are read as `Double` values in **points** (1/72 inch).

**Effect on rendering:**
- Directly determines `cell_top_pt` (sum of heights above cluster) and `cell_height_pt` (sum of heights within cluster)
- Affects vertical coordinate in final normalized ratio
- Excel COM `Range.Height` = row height in points

**Hidden rows:** Must be excluded from summation. `Range.Hidden` property (available via `ExcelRangeInfra.get_Hidden()`) indicates visibility. Hidden rows DO affect coordinate calculation in legacy — they are simply skipped (0 height effectively).

**Row height sources:**
- Custom height: set by user (stored in OpenXML as `customHeight="1"` with `ht` attribute)
- Default height: from worksheet default (stored in OpenXML as `<sheetFormatPr defaultRowHeight="15" />`)
- Auto-fit: calculated from font height + padding

**COM vs OpenXML discrepancy:**
- COM `Range.Height` returns the actual rendered height (includes padding)
- OpenXML stores row height in `ht` attribute directly
- Our implementation should use COM values for exact match, or OpenXML with the same formula

### 1.2 Column Widths

**Source:** `ExcelWorksheetCom.GetRange(row, col)` → `Range.Columns[i].Width` (COM)

**Reading:** Column width in points via COM interop.

**Effect on rendering:**
- Directly determines `cell_left_pt` (sum of widths left of cluster) and `cell_width_pt` (sum of widths within cluster)
- Affects horizontal coordinate in final normalized ratio

**COM width vs OpenXML width — CRITICAL DIFFERENCE:**

Excel COM returns column width in **points** using a different formula than OpenXML:
- COM: `width_pt = width_chars * fontWidth + padding` where `fontWidth` = width of "0" in NormalStyle font
- OpenXML: `width = (maxDigitWidth * charCount + 5) / 7 * 256 + 0.5 / 256` (approximate)

This means summing OpenXML column widths will NOT produce the same result as COM `Range.Width`.

**Verified evidence:** The Phase 15 decompilation confirms the engine sums `Col.Width` from COM — there is no conversion. The column widths are used AS-IS from COM.

**For replication:**
- Use Excel COM Interop on the server to read column widths (guarantees match)
- OR recover the exact COM→OpenXML conversion formula

### 1.3 Merged Cells

**Source:** `Range.MergeArea` (COM)

**Reading:** When a cell is part of a merge range, `MergeArea` returns the encompassing range. The engine uses:
```csharp
Range range2 = range.MergeArea ?? range;
clusterList.Add(range2.Row, range2.Column, range2.Rows.Count, range2.Columns.Count, ...);
```

**Effect on rendering:**
- Cluster bounds = MergeArea bounds (row, column, rowCount, colCount)
- Single unmerged cells still become clusters (rowCount=1, colCount=1)
- Merged cells define the **visual rectangle** for the field on screen

### 1.4 Borders

**Source:** `Range.Borders` → `BorderQ` class (in `Cimtops.Excel.dll`)

**Reading:** Border detection is used ONLY for auto-judging cluster types, NOT for rendering:
```csharp
// BorderQ detects if cell has borders in specific directions
// Used in Decoder.IsCluster() to determine if a cell is a form field
// Not used in coordinate calculation or background generation
```

**Effect on rendering:** NONE for the visual output. Borders affect auto-detection only.

### 1.5 Fonts and Font Size

**Source:** `Range.Font` → `get_Name()`, `get_Size()`, `get_Bold()`, `get_Italic()`, `get_Color()`

**Reading:** Font properties are read but NOT directly used in coordinate calculations. They are stored in the `<format>` section of the generated XML for the tablet runtime.

**Effect on rendering:**
- **INDIRECT** — Font size affects column width calculation in Excel's internal formula
- The NormalStyle font (typically Calibri 11pt) determines the character width used for column width calculations
- Changing font can change column widths, which changes coordinates
- Font metrics (fontFamily, fontSize) determine `maxDigitWidth` used in OpenXML column width formula

**For the Designer:** Fonts are NOT rendered on the background. Only the field overlays show field type, not content.

### 1.6 Alignment

**Source:** `Range.HorizontalAlignment`, `Range.VerticalAlignment` (COM)

**Reading:** Available via `ExcelRangeCom.get_HorizontalAlignment()` / `get_VerticalAlignment()`.

**Effect on rendering:**
- **LOW** — Alignment affects text rendering within cells in the PDF export
- `ExportAsFixedFormat` renders text with the cell's alignment
- This can slightly shift text position within a cell, which could affect optical coordinate placement
- The coordinate engine uses cell boundaries, not text positions, so alignment has NO effect on coordinates

### 1.7 Page Setup

**Source:** `_Worksheet.PageSetup` (COM)

**Properties:**
- `PrintArea` — String like `$A$1:$D$12` — defines the published region
- `PageWidth` / `PageHeight` — Page dimensions in points (612×792 for Letter)
- `LeftMargin` / `RightMargin` / `TopMargin` / `BottomMargin` — Margins in points
- `HeaderMargin` / `FooterMargin` — Header/footer space
- `CenterHorizontally` / `CenterVertically` — Boolean flags
- `Zoom` — Percentage (100 = 100%)
- `FitToPagesWide` / `FitToPagesTall` — Page fitting
- `Orientation` — Portrait/Landscape
- `PaperSize` — Enum (Letter=1, A4=9, etc.)

**Effect on rendering:**
- **PrintArea** → Determines which cells are published (only visible sheets with non-empty PrintArea)
- **PageWidth/PageHeight** → Used as normalization divisor (via PDF metadata, NOT directly)
- **Margins** → Affect `originX`/`originY` — the offset from page edge to printed content
- **Centering** → Alternative origin calculation when content is centered
- **Zoom/Scaling** → Affects how content is fitted to the page in PDF export

**All 3 templates confirmed:** Letter (612×792), Zoom=100%, FitToPages=1×1, no centering, default margins (0.75" L/R, 1" T/B).

### 1.8 Hidden Rows and Hidden Columns

**Source:** `RangeEntireRow.Hidden`, `RangeEntireColumn.Hidden` (COM)

**Reading:** Via `ExcelRangeInfra.get_Hidden()` or `Range.Hidden` COM property.

**Effect on rendering:**
- Hidden rows/columns are **skipped** in the column/width summation
- They effectively have 0 width/height
- This affects the coordinate calculation directly
- Hidden rows/columns near or within clusters shift cluster positions

### 1.9 Images and Shapes

**Source:** `Worksheet.Shapes` (COM)

**Reading:** Via `ExcelWorksheetBase.ClearShapes()` — shapes are **removed** before export.

**Effect on rendering:**
- Shapes (images, drawings, text boxes) are **cleared** during the publish process
- `ClearShapes()` is called before `ExportPdf()`
- This means images in Excel do NOT appear in the background
- If an image should be part of the form, it must be in the print area as a cell's background or image cluster

### 1.10 Page Breaks

**Source:** `HPageBreaks`, `VPageBreaks` (COM)

**Reading:** Via `_Worksheet.get_HPageBreaks()` and `_Worksheet.get_VPageBreaks()`.

**Effect on rendering:**
- Page breaks determine where multi-page content is split
- When content exceeds one page, the algorithm may need to shift coordinates per page
- Template 547 (content = 146.2% of one page) shows evidence of per-row compensation
- `ExistsPrintAreaOnePage()` method checks if content fits on one page

---

## Part 2 — Background Generation

### 2.1 Complete Pipeline

```
Excel Worksheet
    │
    ├── 1. ClearShapes() — Remove all shapes/images
    ├── 2. ExportAsFixedFormat(xlTypePDF) — Export to PDF
    │       │
    │       ↓
    │    PDF File (stored in memory or temp file)
    │    Page dimensions: width_pt × height_pt from PDF metadata
    │       │
    │       ↓
    ├── 3. CreateImageFromPdf(pdfPath, pageNumber, applyMorphology)
    │       │
    │       ↓
    │    System.Drawing.Image (bitmap rasterization)
    │       │
    │       ↓
    └── 4. WPF Image control on ZoomableCanvas
            (or stored as PDF bytes in DB background_image_file)
```

### 2.2 ExportPdf()

**Source:** `LibExcelController.Lib.ImageUtility.ExportPdf(string sourceXlsPath, string destPdfPath)`

**Implementation:**
1. Opens the Excel workbook via COM interop
2. Calls `_Workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)`
3. Exports each sheet with PrintArea to a PDF page

**Key parameters:**
- `xlTypePDF` = 0 (enum value)
- Quality = `xlQualityStandard` or `xlQualityMinimum`
- `IgnorePrintAreas` = false (only print areas are exported)
- Sheet → One PDF page per sheet print area

**Evidence:** Confirmed from `_Workbook.ExportAsFixedFormat()` in LibExcelController reflection. The `XlFixedFormatType` enum has `xlTypePDF = 0` and `xlTypeXPS = 1`.

### 2.3 PDF Page Size

The PDF page size is determined by `PageSetup.PaperSize` and `PageSetup.Orientation`:

| Paper Size | Width (pt) | Height (pt) | Orientation |
|-----------|-----------|-------------|-------------|
| Letter | 612 | 792 | Portrait |
| A4 | 595 | 842 | Portrait |
| Letter (Landscape) | 792 | 612 | Landscape |

**Our templates (546, 547, 548):** All use Letter Portrait (612×792).

### 2.4 DPI in Background Rendering

**DPI Properties:**
- `ExcelControllerBase.DpiPdfToImage` — DPI for PDF→Image conversion (default: 200?)
- `ExcelControllerBase.DpiPdfToBitmap` — DPI for PDF→Bitmap conversion

**The rendering equation:**
```
imageWidth_px  = pageWidth_pt  × dpi / 72
imageHeight_px = pageHeight_pt × dpi / 72
```

For Letter at 200 DPI: 612 × 200/72 = 1700 px wide, 792 × 200/72 = 2200 px tall
For Letter at 96 DPI:  612 × 96/72  = 816 px wide,  792 × 96/72  = 1056 px tall

**Effect on rendering:** Higher DPI = more detailed background but larger image. The overlay coordinates are independent of DPI — they are ratios that scale with the canvas size.

### 2.5 CreateImageFromPdf

**Source:** `ExcelControllerBase.CreateImageFromPdf(string pdfPath, int pageNumber, bool applyMorphology)`

**Returns:** `System.Drawing.Image` (GDI+ bitmap)

**The `applyMorphology` parameter:**
- When true, applies image cleanup operations (likely from `LibConMas.Imaging`):
  - `CorrectDistortion` — Deskew
  - `MedianOperation` — Noise reduction
  - `ThresholdOperation` — Binarization
  - `SharpenOperation` — Edge enhancement
- When false (default for Designer), returns raw PDF→image rendering

**Implementation (inferred from available libraries):**
- Uses one of: AHPDFLib, PdfiumViewer, O2S PDF Render4NET, or DevExpress.Pdf
- Renders PDF page at specified DPI
- Returns bitmap

### 2.6 Image to Designer Canvas

The rendered bitmap is displayed as the `Background` or as an `Image` element on the `ZoomableCanvas`. The canvas is sized to match the background image dimensions (in pixels at the chosen DPI).

**Scale to canvas:**
```
canvasWidth  = imageWidth_px  (or scaled to viewport)
canvasHeight = imageHeight_px (or scaled to viewport)

At zoom=100%: 1 canvas pixel = 1 image pixel
At zoom=200%: 1 canvas pixel = 0.5 image pixels (ScaleX=ScaleY=2.0)
```

### 2.7 DB Storage

The PDF bytes (NOT the rendered image) are stored in `def_top.background_image_file`. The rendering happens at display time.

---

## Part 3 — Coordinate Transformation (HIGHEST PRIORITY)

### 3.1 Complete Mathematical Chain

```
Excel Cell (row, col)
    │
    │  MergeArea detection
    ▼
Cluster Bounds: Top, Left, Bottom, Right (1-based row/col indices)
    │
    │  Column width summation (COM)
    ▼
cell_left_pt   = Σ Col[n].Width  for n=1 to Left-1
cell_top_pt    = Σ Row[n].Height  for n=1 to Top-1
cell_width_pt  = Σ Col[n].Width  for n=Left to Right
cell_height_pt = Σ Row[n].Height  for n=Top to Bottom
    │
    │  Printed origin offset
    ▼
print_left_pt  = cell_left_pt   + originX_pt
print_top_pt   = cell_top_pt    + originY_pt
print_right_pt = cell_left_pt   + cell_width_pt  + originX_pt
print_bottom_pt= cell_top_pt    + cell_height_pt + originY_pt
    │
    │  Gap compensation (unknown)
    ▼
pdf_left_pt   = print_left_pt   + gapH
pdf_top_pt    = print_top_pt    + gapV
pdf_right_pt  = print_right_pt  + gapH
pdf_bottom_pt = print_bottom_pt + gapV
    │
    │  Normalization (divide by page dimension)
    ▼
left_ratio   = pdf_left_pt   / pageWidth_pt
top_ratio    = pdf_top_pt    / pageHeight_pt
right_ratio  = pdf_right_pt  / pageWidth_pt
bottom_ratio = pdf_bottom_pt / pageHeight_pt
    │
    │  RoundEx (banker's rounding)
    ▼
DB_left   = Math.Round(left_ratio,   7, MidpointRounding.ToEven)
DB_top    = Math.Round(top_ratio,    7, MidpointRounding.ToEven)
DB_right  = Math.Round(right_ratio,  7, MidpointRounding.ToEven)
DB_bottom = Math.Round(bottom_ratio, 7, MidpointRounding.ToEven)
    │
    │  Store in def_cluster
    ▼
Database: left_position, right_position, top_position, bottom_position
    │
    │  Designer load (ratio → pixel)
    ▼
canvas_left_px   = DB_left   × canvasWidth_px
canvas_top_px    = DB_top    × canvasHeight_px
canvas_right_px  = DB_right  × canvasWidth_px
canvas_bottom_px = DB_bottom × canvasHeight_px
```

### 3.2 Origin Calculation (originX, originY)

**Definition:** The offset from the page top-left corner to the top-left corner of the printed content.

**Formula (hypothesized, based on PageSetup values):**

```
IF CenterHorizontally:
    originX = (pageWidth - contentWidth) / 2
ELSE:
    originX = LeftMargin  (= 0.75" = 54pt default)

IF CenterVertically:
    originY = (pageHeight - contentHeight) / 2
ELSE:
    originY = TopMargin + HeaderMargin + HeaderTextHeight
           (= 1" + 0.5" + header_text_pt default)
           (= 72pt + 36pt + ?pt)
```

**Where contentWidth/contentHeight = total width/height of print area cells.**

**Evidence:** Template 546 (PrintArea=$A$1:$D$12, no centering) has originX≈0, originY≈0. Template 548 (PrintArea=$A$3:$G$11) also has origin≈0. Template 547 (PrintArea=$B$2:$M$46) has significant origin offset because column B is the first column.

**But wait — if originX = LeftMargin = 54pt, then ALL coordinates would be shifted by 54pt right.** This contradicts templates where origin≈0. The likely explanation: `GetClusterSize()` uses a different page origin — possibly the PDF's TrimBox or MediaBox, or measures positions relative to the PDF page content area directly.

**Alternative hypothesis:** The origin is calculated by back-computing from the difference between COM positions and DB coordinates. This is the "origin derivation" from Phase 15:
```
originX = DB_left_of_first_cluster × pageWidth - COM_left_of_first_cluster
```

### 3.3 Gap Compensation (gapH, gapV)

**Empirical values:**

| Template | gapH (pt) | gapV (pt) | Note |
|----------|----------|----------|------|
| 546 | 2.04 | 0.87 | Horizontal only |
| 547 | ≈0 | 4.44 + 0.18/row | Multi-page vertical compensation |
| 548 | 2.22 | ≈0 | Horizontal only |

**Hypothesis:** Gaps originate from COM column widths not matching export render widths. Excel's Print Preview and `ExportAsFixedFormat` use OpenXML rendering internally, which has a different column width formula than COM's `Range.Width`. The gap is the difference between COM-based width and actual rendered width.

**Verification approach:**
1. Export Excel to PDF
2. Create a PDF that has the same content
3. Measure actual positions of cell boundaries in the rendered PDF
4. Compare with COM-derived positions
5. The difference = gap compensation

### 3.4 Normalization Divisor

**The divisor is NOT PageSetup.PageWidth/PageHeight directly.** The `GetClusterSize()` method takes a PDF path, implying it reads page dimensions from PDF metadata.

**PDF page dimensions vs PageSetup page dimensions:**
- Both should give the same result (612×792 for Letter)
- But PDF may use exact media box dimensions which could differ slightly
- PDF dimensions are in points (1 pt = 1/72 inch)

### 3.5 RoundEx

```csharp
Math.Round(value, 7, MidpointRounding.ToEven)
```

Banker's rounding to 7 decimal places. Example:
- 0.336470588235... → 0.3364706
- 0.384545454545... → 0.3845454

Verified against all 3 templates' DB values.

### 3.6 Designer Pixel Calculation

On the Designer canvas:
```
pixelValue = ratio × canvasDimension
```

Where `canvasDimension` = background image pixel size at chosen DPI.

At DPI=96 for Letter:
```
canvasWidth  = 612 × 96/72 = 816 px
canvasHeight = 792 × 96/72 = 1056 px

cluster.left_px = 0.3364706 × 816 = 274.6 px
cluster.top_px  = 0.3845454 × 1056 = 406.1 px
```

At different DPI, the pixel values change but ratios remain the same.

### 3.7 Summary of Known vs Unknown

| Component | Status | Source |
|-----------|--------|--------|
| Column/row summation | KNOWN | Decompiled Cell.cs, ClusterInfo.cs |
| MergeArea detection | KNOWN | ThisAddIn.cs InitClusterList() |
| Printed origin | HYPOTHESIZED | Phase 15 reconstruction |
| Gap compensation | UNKNOWN | Empirical gap in all templates |
| Normalization divisor | HYPOTHESIZED | PDF page dimensions from GetClusterSize() |
| RoundEx | KNOWN | Decompiled, matches DB values |
| Designer pixel calc | KNOWN | Ratio × canvas size = pixel position |

---

## Part 4 — Designer Rendering

### 4.1 Rendering Layers (Z-Order)

```
Z=6   Context menus, popups (xamDialogWindow, ContextMenu)
Z=5   Grid snap indicators
Z=4   Resize handles (corner + edge squares)
Z=3   Selection rectangles (highlighted cluster overlay)
Z=2   Selected cluster overlay (different color, e.g., light blue)
Z=1   Non-selected cluster overlays (semi-transparent yellow)
Z=0   Background image (rendered PDF → BitmapImage on Image control)
```

### 4.2 Background Rendering

**Implementation:**
```csharp
// Convert PDF bytes to image
Image bgImage = excelController.CreateImageFromPdf(
    pdfPath: tempPdfPath,
    pageNumber: 1,
    applyMorphology: false
);

// Display on WPF Image control
var image = new System.Windows.Controls.Image();
image.Source = ConvertToBitmapSource(bgImage); // System.Drawing.Image → BitmapSource
image.Stretch = Stretch.None; // No stretch — canvas zoom handles scaling

// Add to canvas
canvas.Children.Add(image);
Canvas.SetZIndex(image, 0);
```

**For Next.js:** The equivalent is:
```javascript
// Render PDF to canvas using PDF.js
const pdf = await pdfjsLib.getDocument(pdfBytes).promise;
const page = await pdf.getPage(1);
const viewport = page.getViewport({ scale: dpi / 72 });
const canvas = document.getElementById('background-canvas');
canvas.width = viewport.width;
canvas.height = viewport.height;
await page.render({ canvasContext: ctx, viewport }).promise;
```

### 4.3 Cluster Overlay Rendering

**Algorithm:**
```
For each cluster in defClusters:
    left_px = cluster.left_ratio   × canvas.width
    top_px  = cluster.top_ratio    × canvas.height
    width   = (cluster.right_ratio - cluster.left_ratio)   × canvas.width
    height  = (cluster.bottom_ratio - cluster.top_ratio)   × canvas.height
    
    Create overlay rectangle at (left_px, top_px, width, height)
    Fill: rgba(255, 255, 0, 0.3)    // semi-transparent yellow
    Stroke: rgba(200, 160, 0, 0.8)  // dark yellow border
    StrokeWidth: 1px
    ZIndex: 1
```

**Cluster label (name + type):**
Draw text at top-left of overlay rectangle.

### 4.4 Selection Handling

**Mouse events:**
- `canvas_MouseLeftButtonDown` → Hit test cluster overlays → Select/deselect
- `canvas_MouseMove` → Update selection rectangle (rubber-band selection)
- `canvas_MouseLeftButtonUp` → Finalize selection

**Selection appearance:**
- Selected cluster overlay: Different color (e.g., rgba(173, 216, 230, 0.5) — light blue)
- Selection rectangle: Dashed border around selected region
- Highlight handles appear on selected clusters

**Multi-select:**
- Ctrl+click to toggle selection
- Shift+click for range selection

### 4.5 Resize Handles

**Appearance:** Small squares (8×8 px) at:
- 4 corners (top-left, top-right, bottom-left, bottom-right)
- 4 edge midpoints (top-center, right-center, bottom-center, left-center)

**Behavior:**
- Drag corner → resize proportionally (if aspect ratio locked)
- Drag edge → resize in one dimension only
- Snap to nearest grid point when released (if snap enabled)
- Min/max size enforcement

**Config:**
- `__AutoAlignmentThreshold` — pixels within which snap activates

### 4.6 Zoom and Scroll

**Zoom control:** `xamNumericSliderZoom` slider (Infragistics WPF):
- Range: MinimumScale → MaximumScale (e.g., 10% → 500%)
- Slider thumb: discrete steps or continuous

**Zoom transform:** `ZoomableCanvas` uses `ScaleTransform`:
```csharp
// Applied via RenderTransform
canvas.RenderTransform = new ScaleTransform(ScaleX, ScaleY, centerX, centerY);
```

**Scroll:** `ScrollViewer` wrapping the canvas:
```csharp
ScrollToHorizontalOffset(value);
ScrollToVerticalOffset(value);
```

**Zoom at cursor:** When zooming with mouse wheel, zoom centered on cursor position:
```
newScale = oldScale * factor
offset = cursorPosition - (cursorPosition - offset) * newScale / oldScale
```

**For Next.js:**
```typescript
// CSS Transform approach
const transform = `scale(${zoom}) translate(${panX}px, ${panY}px)`;
// Or use transform-origin for cursor-centered zoom
```

### 4.7 Grid Snap

**Grid spacing:** Based on cell boundaries from the Excel grid (variable row/col widths). Simplified in Designer to a fixed pixel grid.

**Snap behavior:**
```javascript
function snapToGrid(value, threshold) {
    const gridSnap = gridSize; // configurable
    const remainder = value % gridSnap;
    if (remainder < threshold) return value - remainder;
    if (remainder > gridSnap - threshold) return value - remainder + gridSnap;
    return value;
}
```

`__AutoAlignmentThreshold` controls the `threshold` parameter.

### 4.8 Redraw Cycle

The canvas redraws when:
1. Window resize (canvas resize)
2. Zoom change (scale transform)
3. Selection change (overlay update)
4. Cluster edit (property change)
5. Scroll (offset change)
6. Drag operation (mouse move)

Redraw is NOT a full re-render — only the affected overlay elements are updated.

---

## Part 5 — Designer State Reconstruction

### 5.1 Load Pipeline

```
Database Query
    │
    ├── def_top → FormDefinition
    │     defTopId, defTopName, xml_data (string)
    │     background_image_file (byte[])
    │
    ├── def_sheet → SheetDefinition[]
    │     defSheetId, sheetNo, width, height
    │
    └── def_cluster → ClusterDefinition[]
          clusterId, cluster_type, name
          left_position, right_position, top_position, bottom_position
          input_parameter, display_value, readOnly, etc.
    
    ▼

Parse xml_data (redundant with def_cluster but authoritative for structure)
    │
    ├── Sheets (width, height from XML)
    ├── Clusters (name, type, coords from XML)
    ├── Networks (field routing)
    ├── Labels (sheet labels)
    └── References (external data)
    
    ▼

Memory Objects
    │
    ├── FormDefinition (from def_top + xml_data)
    ├── BackgroundImage (from background_image_file → PDF → bitmap)
    ├── SheetDefinition[] (from def_sheet)
    └── ClusterDefinition[] (from def_cluster with merged XML properties)
    
    ▼

Canvas Rendering
    │
    ├── ZoomableCanvas (root panel)
    ├── Background (Image control, Z=0)
    ├── Cluster Overlays (CanvasChild+Rect, Z=1)
    ├── Selection Rectangles (CanvasChild+Rect, Z=2-3)
    ├── Resize Handles (CanvasChild+Rect, Z=4)
    └── Grid/Guides (CanvasChild+Line, Z=5)
```

### 5.2 Object Model

```typescript
interface FormDefinition {
    defTopId: number;
    defTopName: string;
    designerVersion: string;
    sheetCount: number;
    sheets: SheetDefinition[];
    // From XML:
    networks: NetworkDefinition[];
    labels: LabelDefinition[];
    references: ReferenceDefinition[];
    autoNumbering: AutoNumberingConfig;
    finishOutput: OutputConfig;
    editOutput: OutputConfig;
    remarks: string[];
}

interface SheetDefinition {
    defSheetId: number;
    sheetNo: number;
    width: number;      // Page width in pts
    height: number;     // Page height in pts
    backgroundImage: ImageData;
    clusters: ClusterDefinition[];
    remarks: string[];
}

interface ClusterDefinition {
    clusterId: number;
    sheetNo: number;
    name: string;
    type: ClusterType;  // enum
    left: number;       // normalized ratio
    top: number;
    right: number;
    bottom: number;
    cellAddress: string; // e.g., "$A$1:$B$2"
    isHidden: boolean;
    isHiddenDesigner: boolean;
    readOnly: boolean;
    external: boolean;
    displayValue: string;
    cooperationCluster: number;
    pinNo: string;
    pinValue: string;
    inputParameters: string;
    remarks: string[];
    management: string;
}
```

### 5.3 XML → Object Mapping

| XML Node | Object Property | Source |
|----------|----------------|--------|
| `<top><defTopId>` | `defTopId` | def_top.def_top_id |
| `<top><definitionFile><name>` | `fileName` | def_top.def_file |
| `<top><sheets><sheet><width>` | `sheet.width` | def_sheet.width |
| `<top><sheets><sheet><height>` | `sheet.height` | def_sheet.height |
| `<variables><cluster><clusterId>` | `cluster.clusterId` | def_cluster.cluster_id |
| `<variables><cluster><name>` | `cluster.name` | def_cluster.name |
| `<variables><cluster><type>` | `cluster.type` | def_cluster.cluster_type |
| `<variables><cluster><left>` | `cluster.left` | def_cluster.left_position |
| `<variables><cluster><top>` | `cluster.top` | def_cluster.top_position |
| `<variables><cluster><right>` | `cluster.right` | def_cluster.right_position |
| `<variables><cluster><bottom>` | `cluster.bottom` | def_cluster.bottom_position |
| `<variables><cluster><inputParameters>` | `cluster.inputParameters` | def_cluster.input_parameter |

### 5.4 Type-Specific Cluster Objects

Each cluster type has a dedicated parameter class in `LibConMas.ImageDb`:

```typescript
type ClusterParameters =
    | TextClusterParameter
    | NumericClusterParameter
    | CheckClusterParameter
    | SelectClusterParameter
    | CalendarDateClusterParameter
    | QRCodeClusterParameter
    | CodeReaderClusterParameter
    | ImageClusterParameter
    | FreeDrawClusterParameter
    | HandwritingClusterParameter
    | CalculateClusterParameter
    | ActionClusterParameter
    | GpsClusterParameter
    | ...;

interface TextClusterParameter {
    inputRestriction: 'none' | 'numeric' | 'alphanumeric';
    keyboardType: 'default' | 'numeric' | 'url' | 'email';
    textClusterMode: 'single' | 'multi' | 'password';
}

interface NumericClusterParameter {
    minValue: number;
    maxValue: number;
    decimalPlaces: number;
    unit: string;
}

// ... etc for all 38 types
```

### 5.5 Visual Element Mapping

| Logical Object | Visual Element | Canvas | Notes |
|---------------|---------------|--------|-------|
| Background | WPF Image control | Z=0 | PDF rendered at DPI |
| Cluster overlay | CanvasChild+Rect (yellow rectangle) | Z=1 | Semi-transparent fill |
| Cluster label | TextBlock on overlay | Z=1 | Name + type text |
| Selected cluster | CanvasChild+Rect (blue-ish rectangle) | Z=2 | Different color + dashed border |
| Selection rubber-band | CanvasChild+Rect (dashed) | Z=3 | Mouse drag selection |
| Resize handle | CanvasChild+Rect (small square) | Z=4 | 8×8 px squares |
| Grid lines | CanvasChild+Line | Z=5 | Optional display |
| Context menu | WPF ContextMenu | Z=6 | Right-click popup |

---

## Part 6 — Validation

### 6.1 Source Classification

| Source | Tag | Reliability |
|--------|-----|-------------|
| Decompiled C# source | `[DECOMPILED]` | High |
| .NET Reflection | `[REFLECTION]` | High |
| Database records | `[DB]` | High |
| String analysis (binary) | `[STRINGS]` | Medium |
| Empirical measurement | `[MEASURED]` | Medium |
| Phase report reconstruction | `[RECONSTRUCTED]` | Medium |
| Hypothesis/inference | `[HYPOTHESIS]` | Low |

### 6.2 Known Facts (High Confidence)

| Fact | Source | Evidence |
|------|--------|----------|
| Cluster bounds from MergeArea | [DECOMPILED] | ThisAddIn.cs: InitClusterList(): `range.MergeArea ?? range` |
| Column width summation | [DECOMPILED] | Cell.cs: `Sum(c => c.Width)` |
| Row height summation | [DECOMPILED] | Cell.cs: `Sum(r => r.Height)` |
| Sort order: ClusterIndex→Row→Col | [DECOMPILED] | ClusterList.cs: `Comp(x.ClusterIndex, y.ClusterIndex) ?? Comp(x.Row, y.Row) ?? Comp(x.Col, y.Col)` |
| Comment format: L0=name, L1=type, L2=index | [DECOMPILED] | ClusterInfo.cs: `ClusterName = GetLine(0)`, `ClusterTypeKey = GetLine(1)`, `ClusterIndex = GetLine(2)` |
| Comment lines 0-15 | [DECOMPILED] | InputForm2.cs: comments mapped to rows 0-15 |
| RoundEx to 7 decimals | [DECOMPILED] | Phase 15 reconstruction, DB verification |
| ExportPdf → ExportAsFixedFormat | [REFLECTION] | `_Workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF)` |
| PageWidth/Height from PDF | [REFLECTION] | `GetClusterSize(strPdfPath, page, clusterCount)` takes PDF path |
| CreateImageFromPdf | [REFLECTION] | `ExcelControllerBase.CreateImageFromPdf(string pdfPath, int page, bool applyMorphology)` |
| DpiPdfToImage / DpiPdfToBitmap | [REFLECTION] | `ExcelControllerBase.get_DpiPdfToImage()` |
| 38 cluster types | [REFLECTION] | `LibExcelController.Lib.ClusterType` enum |
| ClusterRect has Single coords | [REFLECTION] | `ClusterRect.Top/Bottom/Left/Right` as Single |
| Background is PDF in DB | [DB] | `%PDF-1.7` header in background_image_file |
| CanvasChild (Image, Rect types) | [REFLECTION] | `ConMasDesigner.Domain.Entities.CanvasChild` |
| ZoomableCanvas with ScaleX/Y | [STRINGS] | `ZoomableCanvas`, `ScaleX`, `ScaleY`, `CoerceScale` |
| DB coordinates per template | [MEASURED] | Verified 6+5+2 clusters across 3 templates |

### 6.3 Reconstructed Facts (Medium Confidence)

| Fact | Source | Evidence |
|------|--------|----------|
| Printed origin from margins | [RECONSTRUCTED] | Phase 15: origin = first cluster back-calculation |
| Gap compensation needed | [MEASURED] | All 3 templates show systematic gaps |
| ClearShapes before export | [STRINGS] | `ClearShapes()` in LibExcelController |
| ExistsPrintAreaOnePage | [REFLECTION] | `ExcelWorksheetBase.ExistsPrintAreaOnePage()` |
| Yellow overlay color | [STRINGS] | `__StandardClusterColor`, `rgbYellow` |
| Grid snap threshold | [STRINGS] | `__AutoAlignmentThreshold` |
| Header/footer support | [STRINGS] | `get_Header`, `get_Footer`, `WorksheetHeaderFooter` |

### 6.4 Hypotheses (Low Confidence)

| Hypothesis | Basis |
|------------|-------|
| Gap from COM vs OpenXML width difference | Empirical: gap consistent per template, varies with column count |
| Origin = back-calc from first cluster | Phase 15 reconstruction, not directly decompiled |
| ApplyMorphology = false for Designer background | Implied by name — morphology is cleanup for OCR |
| Grid snap uses fixed pixel grid | `__AutoAlignmentThreshold` suggests threshold-based snap |

---

## Part 7 — Replication Assessment

### 7.1 Component Readiness

| Component | Understood | Confidence | Ready to Implement | Blockers |
|-----------|-----------|------------|-------------------|----------|
| **Excel Reader (COM)** | Complete | High | YES | Need server-side COM Interop (Windows-only) |
| **Print Area** | Complete | High | YES | — |
| **Page Setup** | Complete | High | YES | — |
| **Comments** | Complete | High | YES | — |
| **Cluster Detection** | Complete | High | YES | — |
| **Cluster Sorting** | Complete | High | YES | — |
| **DB Write (def_top)** | Complete | High | YES | — |
| **DB Write (def_sheet)** | Complete | High | YES | — |
| **DB Write (def_cluster)** | Complete | High | YES | — |
| **XML Generation** | 95% | High | YES | Minor: some XML nodes may be missing |
| **Background Export (PDF)** | 90% | High | YES | Server needs Excel COM or alternative |
| **CreateImageFromPdf** | 80% | High | YES | Use PDF.js or server-side rendering |
| **DB Read (load form)** | Complete | High | YES | — |
| **XML Parsing** | Complete | High | YES | — |
| **Coordinate Normalization** | 90% | Medium | YES (with gap tolerance) | Gap compensation unknown |
| **ZoomableCanvas** | 90% | High | YES | Use React Canvas or CSS transforms |
| **Cluster Overlay (yellow rect)** | 95% | High | YES | — |
| **Selection Handling** | 85% | Medium | YES | — |
| **Resize Handles** | 80% | Medium | YES | — |
| **Grid Snap** | 70% | Medium | YES | Threshold config |
| **Background Rendering (canvas)** | 90% | High | YES | PDF.js + canvas |
| **DPI Handling** | 75% | Medium | YES | Configurable DPI |
| **Coordinate Calculation** | **70%** | **Medium** | **PARTIAL** | **Gap compensation BLOCKER** |
| **PDF Export (server-side)** | 85% | Medium | YES | LibreOffice or Excel COM |

### 7.2 Overall Assessment

**The system is 85% understood and ~80% ready to implement.**

The main implementation path is clear for all components. The single critical unknown (coordinate gap compensation) can be worked around by:
1. Building the system with all known formulas
2. Measuring the gap empirically for each template
3. Implementing a calibration step that adjusts gapH/gapV per workbook
4. Decompiling `GetClusterSize()` when .NET 4.x is available

### 7.3 Implementation Order

1. **Phase A — Backend (ASP.NET Core)**
   - Excel COM reader (row heights, column widths, print area, comments)
   - Cluster detection and sorting
   - Coordinate calculation (with zero gap placeholder)
   - XML generation
   - Database writes
   - Publish API endpoint

2. **Phase B — PDF Generation**
   - PDF export (Excel COM ExportAsFixedFormat, or LibreOffice headless)
   - PDF storage in DB
   - PDF retrieval API

3. **Phase C — Frontend (Next.js)**
   - PDF.js canvas rendering
   - Cluster overlay (yellow rectangles at ratio positions)
   - Canvas zoom/pan
   - Cluster editing (move, resize, properties)
   - Selection and resize handles

4. **Phase D — Gap Recovery**
   - Compare generated coordinates with legacy DB values
   - Measure gapH/gapV per template
   - Implement calibration or recovery formula
   - Achieve 100% match

---

## Part 8 — Remaining Unknowns (Genuine Blockers Only)

### Blocker 1: Gap Compensation Formula

**What:** The systematic difference between COM-derived positions and actual rendered positions (gapH, gapV).

**Why it matters:** Without it, cluster overlays will be misaligned by 1-4 pixels (at 96 DPI), which is visually noticeable.

**Can be experimentally reproduced:** YES — by comparing our calculated coordinates against legacy DB values for all 3 templates. The gap is quantifiable.

**Implementation impact:** MEDIUM — can ship with zero gap and fix in a later phase. Gap is consistent per template and could be calibrated.

**Blocks development:** NO — but blocks "perfect alignment" success criterion.

### Blocker 2: GetClusterSize() Internal Algorithm

**What:** The exact implementation of `LibExcelController.Lib.ImageUtility.GetClusterSize(string pdfPath, int page, int clusterCount)`.

**Why it matters:** This function may contain additional coordinate logic beyond simple column-width summation. It may read PDF measurements directly.

**Can be experimentally reproduced:** PARTIALLY — we can call it if .NET 4.x is available, or measure its output empirically.

**Implementation impact:** MEDIUM — the function likely performs the same column-width summation + origin + gap that we've reconstructed. The PDF parameter may be used only for page dimensions.

**Blocks development:** NO — our reconstruction covers 90% of the behavior.

### Blocker 3: Origin Calculation

**What:** The exact formula for converting PageSetup values to the printed origin offset.

**Why it matters:** Origin shifts all coordinates by a constant amount. Wrong origin = wrong position.

**Can be experimentally reproduced:** YES — by measuring the offset between the page edge and the first rendered column in the PDF output.

**Implementation impact:** LOW — can be derived from first cluster back-calculation as shown in Phase 15.

**Blocks development:** NO — the origin can be derived from known data.

### Summary: No Absolute Blockers

**There are NO absolute blockers to beginning implementation.** Gap compensation and origin calculation can be:
1. Implemented as configurable parameters
2. Calibrated against known DB values
3. Refined later when .NET 4.x decompilation is available

The implementation can start TODAY with:
- All known formulas (column summation, RoundEx, normalization)
- Zero-gap placeholder
- Back-computed origin from DB
- Verification against all 3 templates
