# Designer Rendering Pipeline

## Overview

The Designer (ConMasClient.exe) is a WPF application that renders published forms using a `ZoomableCanvas` with overlay elements. The rendering pipeline:

```
Database (def_top, def_sheet, def_cluster, background_image_file)
    │
    ├──► Load def_top → Form metadata (xml_data)
    ├──► Load def_sheet → Sheet dimensions
    ├──► Load def_cluster → Cluster positions (ratios) + metadata
    └──► Load background_image_file → PDF bytes
    
    ▼
    
Canvas (ZoomableCanvas)
    ├──► Background Image (PDF rendered to Bitmap)
    ├──► Cluster Overlays (yellow rectangles)
    ├──► Selection Rectangles
    ├──► Resize Handles
    └──► Grid Snapping Layer
```

## Canvas Architecture

### ZoomableCanvas (Custom WPF Panel)
The Designer uses a custom `ZoomableCanvas` (from `ConMasDesigner.Domain.Entities` namespace in LibConMas.dll):

Properties:
- `ScaleX` / `ScaleY` — Zoom level (via `xamNumericSliderZoom`)
- `OffsetX` / `OffsetY` — Pan offset
- `MinimumScale` / `MaximumScale` — Zoom limits
- `CoerceScale()` — Clamps scale to valid range
- `RealizeOverride()` / `ArrangeOverride()` — Custom layout logic
- `RenderTransform` / `LayoutTransform` — Transform pipeline
- `GeneralTransformGroup` — Combined transforms

Events:
- `canvas_MouseLeftButtonDown` — Selection start
- `canvas_MouseMove` — Selection/drag update
- `canvas_MouseLeftButtonUp` — Selection/drag end
- `canvas_MouseLeave` — Cancel operations
- `canvas_ContextMenuOpening` — Context menu

### CanvasChild Elements
From `ConMasDesigner.Domain.Entities.CanvasChild` in LibConMas.dll:

Types (discriminated by `CanvasChildType`):
- `Image` — Background or cluster image
- `Rect` — Selection/highlight rectangle

Each CanvasChild has:
- Position (from ratio × canvas size)
- Size (from ratio extents × canvas size)
- Z-order (layer ordering)
- Selection state

## Background Rendering

### Pipeline:
```
PDF bytes (from DB background_image_file)
    │
    ├──► Write to temporary file
    │
    ├──► ConvertPdfToImage(pdfPath, pageNumber, applyMorphology)
    │     Via LibExcelController.ExcelControllerBase.CreateImageFromPdf()
    │     OR via O2S PDF Render4NET / PdfiumViewer / AHPDFLib
    │
    ├──► DPI scaling: DpiPdfToImage (default), DpiPdfToBitmap
    │     LOGPIXELSX / LOGPIXELSY (GDI) for device pixel mapping
    │
    └──► Set as Background or Image on canvas
```

Background image format:
- The `BGImageFormat` enum (from `ConMasDesigner.Domain.Enums`) defines supported formats
- String analysis confirms: PDF, XPS, GIF, TIFF, PNG, JPG

## Cluster Overlay Rendering

### Coordinate Conversion
Each def_cluster stores normalized ratios:
```
left_px   = left_ratio   × canvasWidth
top_px    = top_ratio    × canvasHeight
right_px  = right_ratio  × canvasWidth
bottom_px = bottom_ratio × canvasHeight
```

### Rendering Layers (Z-order)
```
Z=0: Background Image (PDF raster)
Z=1: Cluster overlays (yellow semi-transparent rectangles)
Z=2: Selected cluster overlay (different color, e.g., light blue)
Z=3: Resize handles (small squares at corners/edges)
Z=4: Grid lines / snap indicators
Z=5: Context menus, popups
```

### Colors
- `StandardClusterColor` (config key: `__StandardClusterColor`) — Default overlay color (yellow)
- `SelectedClusterColor` (config key: `__SelectedClusterColor`) — Selected overlay color
- RGB color definitions found in strings: `rgbYellow`, `rgbLightGoldenrodYellow`, `rgbYellowGreen`

### Selection
- `xamDataGridClusters` — Grid control showing all clusters as a list
- Selection in grid → highlights cluster on canvas
- Selection on canvas → selects in grid
- Multi-select supported (`SelectionMode.MultiExtended`)
- Selection rectangles drawn as `CanvasChild+Rect` elements

### Resize
- Resize handles at corners and edge midpoints
- Constrained by grid snapping
- `AutoAlignmentThreshold` config key controls snap sensitivity

### Grid Snapping
- Configurable via settings (detected in ConMasClient.exe strings)
- Snaps to device pixels (`SnapsToDevicePixels`)
- Alignment threshold controlled by `__AutoAlignmentThreshold`

### Zoom and Scroll
- Zoom controlled via `xamNumericSliderZoom`
- `statusBarItemZoomRatio` displays current zoom percent
- `labelZoomRatio` shows ratio text
- `ScrollToHorizontalOffset` / `ScrollToVerticalOffset` for panning
- `HorizontalOffset` / `VerticalOffset` properties
- `MinimumScale` / `MaximumScale` bound zoom range

## Cluster Type Display

Each cluster type determines how the field is displayed on the canvas:
- `KeyboardText` — TextBox overlay
- `InputNumeric` — NumericBox overlay
- `Check` — CheckBox overlay
- `Select` — ComboBox overlay
- `CalendarDate` — DatePicker overlay
- `Image` — Image placeholder
- `FreeDraw` — InkCanvas overlay
- `FixedText` / `FreeText` — Label (read-only text)

## WPF Controls Used

From the InfragisticsWPF4 v10.3 suite:
- `xamDataGrid` (for cluster lists, tables, networks)
- `xamNumericSlider` (for zoom slider)
- `xamColorPicker` (for color selection)
- `xamDialogWindow` (for modal dialogs)
- `xamDataTree` (for form structure tree)
- `xamMenu` (for context menus)

From standard WPF:
- `ZoomableCanvas` (custom)
- `InkCanvas` (for FreeDraw)
- `Image` (for background, Image clusters)
- `Canvas` (base overlay container)

## DPI Handling

The Designer handles high-DPI displays:
- `SetProcessDPIAWARE` — makes process DPI-aware
- `DpiPdfToImage` / `DpiPdfToBitmap` — controls PDF rendering resolution
- `LOGPIXELSX` / `LOGPIXELSY` — GDI constants for device pixel mapping
- `SnapsToDevicePixels` — WPF property for pixel-perfect rendering
- `maxPixelsOfLongSide` / `minPixelsOfShortSide` — image size constraints
