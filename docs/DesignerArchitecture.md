# Designer Architecture — Complete Pipeline

## Overview

The PaperLess system has three distinct applications working together:

```
Excel Workbook (.xlsx)
    │
    ├───► [iReporterExcelAddIn] — Excel COM Add-In (WinForms)
    │         Defines clusters via Excel Comments
    │         Auto-judges field types from borders/colors
    │         
    ├───► [ConMasClient.exe] — Designer (WPF + Infragistics)
    │         Loads forms from DB for editing
    │         Publishes forms to DB
    │         Exports PDF background
    │
    └───► [ConMasExcelClient.dll] — Headless Excel Controller
              Handles Excel COM operations
              ExportPdf, GetClusterSize, ExportExcelDefinition
```

## Assembly Map

### LibExcelController.dll (81 KB, .NET 4.8)
Abstract layer over Excel COM interop. Two implementations:
- `Com` variant: Uses `Microsoft.Office.Interop.Excel` directly
- `Infra` variant: Uses `Infragistics.Documents.Excel` (OpenXML, no Excel required)

Key classes:
- `ExcelControllerBase` — Abstract base. Properties: `DpiPdfToImage`, `DpiPdfToBitmap`. Methods: `ExportExcelDefinition()`, `CreateImageFromPdf()`
- `ImageUtility` — Static utility. Methods: `ExportPdf(sourceXlsPath, destPdfPath)`, `GetClusterSize(pdfPath, page, clusterCount)`, `CanExportPdf()`, `GetExcelType()`
- `ClusterRect` — Float-based rect: Top/Bottom/Left/Right (all Single). Implements `IComparable<ClusterRect>`.
- `CalculateCell` — DTO: SheetNo, Row, Col, CellAddress

### LibConMas.dll (1.6 MB, .NET 4.8)
Core business logic and image processing library.

Namespaces:
- `LibConMas.Imaging` — Image processing pipeline (35+ operations: deskew, threshold, resize, OCR prep, barcode, contour detection)
- `LibConMas.ImageDb` — All cluster parameter entities (TextClusterParameter, NumericClusterParameter, CheckClusterParameter, etc. — 60+ types)
- `ConMasDesigner.Domain.Enums` — 50+ enum types (BGImageFormat, PageClusterType, KeyboardType, TextClusterMode, etc.)
- `ConMasDesigner.Domain.Entities` — `CanvasChild`, `ClusterRow`, `ClusterTable`, `ColorDefinition`, `FontSetting`, `InputParameterModel`, etc.
- `ConMasDesigner.Domain.Helpers` — `ClusterMethods`, `EdgeOCRMethods`, `PageClusterMethods`, `ParameterMapper`, `XElementMapper`

### ConMasClient.exe (3.8 MB, .NET 4.8 WPF)
The Designer application. Uses:
- `ZoomableCanvas` (custom WPF canvas with ScaleX/ScaleY/Offset transforms)
- `CanvasChild` (Image, Rect types for rendering overlays)
- `InfragisticsWPF4.v10.3` (xamDataGrid, xamSlider, xamColorPicker, xamDialogWindow, xamDataTree, xamMenu)
- PDF rendering via AH/PdfTk libraries
- DPI awareness via `SetProcessDPIAWARE`, `DpiPdfToImage`, `DpiPdfToBitmap`

## Data Flow: Excel → Database → Designer

### Phase 1: Form Design (Excel Add-In)
```
1. User creates Excel workbook with form layout
2. User inserts Comments on cells that are form fields
   Comment Line 0 = Cluster Name (e.g., "Machine")
   Comment Line 1 = Cluster Type Key (e.g., "KeyboardText")
   Comment Line 2 = Cluster Index (sort order)
   Comment Lines 3-15 = Extension flags, remarks 1-10
   Comment Line 16 = Table number (if part of a table)
3. AutoJudge() reads borders/colors to auto-detect clusters
4. InputForm2 (WinForms) lets user edit cluster properties
```

### Phase 2: Publish (Designer → Database)
```
1. ConMasClient.exe calls LibExcelController:
   a. ExportExcelDefinition() — analyzes workbook
   b. ExportPdf() — exports each sheet to PDF
   c. GetClusterSize(pdfPath, page, clusterCount) — calculates coordinates

2. Database writes (single transaction):
   a. INSERT def_top (form metadata)
   b. INSERT def_sheet (sheet metadata, width, height)
   c. INSERT def_cluster (cluster with normalized coords)
   d. UPDATE def_top SET xml_data (full form XML)
   e. UPDATE def_top SET background_image_file (PDF bytes)
   f. INSERT def_current (publish completion marker)

3. Coordinate calculation (via GetClusterSize):
   a. Read column widths and row heights from Excel COM
   b. Sum widths left of cluster → left_pt
   c. Sum heights above cluster → top_pt
   d. Apply printed origin shift (originX, originY)
   e. Normalize: ratio = pt / pageDimension
   f. RoundEx to 7 decimal places (MidpointRounding.ToEven)
```

### Phase 3: Designer Load (Database → Screen)
```
1. Load from DB:
   a. def_top → form metadata, xml_data
   b. def_sheet → sheet dimensions (width, height in points)
   c. def_cluster → cluster list with normalized coords
   d. background_image_file → PDF bytes

2. Parse XML (redundant with def_cluster but authoritative for form structure):
   a. Sheet dimensions (width/height)
   b. Cluster properties (name, type, coordinates, inputParameters)

3. Render background:
   a. PDF bytes → temporary file
   b. ConvertPdfToImage() or PDF library (AHPDFLib / O2S / Pdfium)
   c. Scale to canvas size using DpiPdfToImage/DpiPdfToBitmap settings

4. Create field overlays:
   a. For each def_cluster record:
      left_px  = ratio_left  * canvasWidth
      top_px   = ratio_top   * canvasHeight
      right_px = ratio_right * canvasWidth
      bottom_px = ratio_bottom * canvasHeight
   b. Create CanvasChild elements (Rect border overlay, type-specific controls)
   c. Color: yellow for standard, configurable for selected

5. Enable editing:
   a. Grid snapping
   b. Resize handles
   c. Zoom via ZoomableCanvas (ScaleX/ScaleY)
   d. Selection rectangles
```

### Key Insight: Two Coordinate Paths
The Designer uses BOTH `def_cluster` AND `xml_data`:
- `def_cluster` is the authoritative source for coordinates (ratios in left_position, right_position, top_position, bottom_position)
- `xml_data` is the authoritative source for form structure (cluster names, types, inputParameters, networks, labels)

## Technology Stack Summary

| Component | Technology |
|-----------|------------|
| Excel Add-In | VSTO, WinForms, .NET 4.8 |
| Designer (ConMasClient.exe) | WPF, InfragisticsWPF4 v10.3, .NET 4.8 |
| Excel Controller | Microsoft.Office.Interop.Excel + Infragistics.Excel |
| Image Processing | LibConMas.Imaging (OpenCV-based via OpenCvSharp) |
| PDF Export | Microsoft.Office.Interop.Excel ExportAsFixedFormat |
| PDF Rendering | AHPDFLib / O2S PDF Render4NET / Pdfium |
| Database | PostgreSQL (via Npgsql) |
| Background Storage | PDF bytes in background_image_file column |
| Config Format | XML (conmas schema) in xml_data column |
