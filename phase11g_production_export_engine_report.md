# Phase 11G — Production Export Engine (PNG/PDF)

**Date:** July 2026  
**Build:** 0 errors, 20 warnings  
**Status:** ✅ Complete

---

## Objective

Transform the FormLess Rendering Engine into a production-quality export engine capable of generating PNG and PDF output that matches Microsoft Excel's printed output with pixel-level accuracy.

This phase builds on the completed rendering pipeline and does not redesign the architecture — it creates the final export layer.

---

## Current Rendering Pipeline

```
Workbook
    ↓
OpenXmlParser
    ↓
StyleResolver (Phase 11E)
    ↓
GeometryBuilder
    ↓
PrintLayoutEngine (Phase 11D)
    ↓
PageRenderer         ← NEW
    ↓
RenderingContext
    ↓
FillEngine          Layer 1
    ↓
GridlineLayer       Layer 2
    ↓
BorderEngine        Layer 3
    ↓
TextEngine          Layer 4
    ↓
ImageEngine         Layer 5
    ↓
ShapeEngine         Layer 6
    ↓
ExportCoordinator   ← NEW (public export API)
    ↓
PNG / PDF
```

---

## Files Created (3 new)

| File | Purpose |
|:-----|:--------|
| `Rendering/ExportOptions.cs` | Configurable export settings — DPI (96/150/300/600), transparent background, font embedding, multi-page limits, page numbering, PNG compression, metadata fields |
| `Rendering/PageRenderer.cs` | Multi-page splitting engine — divides worksheet into printable pages respecting PrintArea, margins, scaling, FitToPage, centering, orientation. Returns `List<PageInfo>` with per-page `RenderingContext` |
| `Rendering/ExportCoordinator.cs` | Public export API — `ParseWorkbook()` for shared setup, `ExportPng()` → list of file paths, `ExportPdf()` → single PDF with metadata. Orchestrates parse → geometry → pages → render → output |

## Files Modified (4)

| File | Change |
|:-----|:-------|
| `Generators/PreviewGenerator.cs` | Now delegates to `ExportCoordinator.ExportPng()` — no rendering code remains. Falls back to overlay-only if XLSX unavailable. |
| `Generators/PdfGenerator.cs` | Now delegates to `ExportCoordinator.ExportPdf()` — no rendering code remains. Falls back to overlay-only if XLSX unavailable. |
| `Rendering/RenderingContext.cs` | `Workbook` and `Sheet` properties changed from `init` to `set` for post-construction assignment by ExportCoordinator |
| `Program.cs` | Registered `ExportOptions`, `PageRenderer`, `ExportCoordinator` as singletons |

---

## ExportCoordinator Architecture

```
ExportCoordinator (public API)
    │
    ├── ParseWorkbook(xlsxPath)
    │   ├── OpenXmlParser.Parse()
    │   └── GeometryBuilder.ComputeGeometry()
    │
    ├── ExportPng(xlsxPath, form, sheetId, options)
    │   ├── ParseWorkbook()
    │   ├── PageRenderer.ComputePages() → List<PageInfo>
    │   └── ForEach page:
    │       ├── SKBitmap + SKCanvas
    │       ├── RendererCoordinator.RenderToCanvas()
    │       ├── DrawPageNumber() [optional]
    │       └── Encode + Save to PNG
    │
    └── ExportPdf(xlsxPath, form, sheetId, outputPath, options)
        ├── ParseWorkbook()
        ├── PageRenderer.ComputePages() → List<PageInfo>
        ├── SKDocument.CreatePdf() with metadata
        └── ForEach page:
            ├── pdfDocument.BeginPage()
            ├── Stage bitmap + RenderToCanvas()
            ├── DrawImage() to PDF canvas with scale
            └── pdfDocument.EndPage()
```

---

## PageRenderer Multi-Page Algorithm

```
Input:
  - Page settings (paper size, orientation, margins)
  - Content dimensions (totalContentWidthPt, totalContentHeightPt)
  - DPI, max pages

Algorithm:
  1. Compute base layout via PrintLayoutEngine
  2. Calculate scaled content size (content * scaleFactor)
  3. Determine pagesWide = fitToPagesWide OR ceil(scaledWidth / printableWidth)
  4. Determine pagesTall = fitToPagesTall OR ceil(scaledHeight / printableHeight)
  5. For each (pageX, pageY):
     - originXPt = layout.originX + pageX * printableWidthPt
     - originYPt = layout.originY + pageY * printableHeightPt
     - clipLeft = originX + pageX * printableWidthPt
     - clipRight = min(clipLeft + printableWidth, originX + scaledContentWidth)
     - Build per-page RenderingContext
     - Return PageInfo { PageNumber, TotalPages, Context, PixelWidth, PixelHeight }

Output: List<PageInfo> — one per printable page
```

---

## ExportOptions

| Option | Type | Default | Description |
|:-------|:----:|:-------:|:------------|
| `Dpi` | int | 300 | Rendering DPI (96/150/300/600) |
| `TransparentBackground` | bool | false | Transparent instead of white |
| `EmbedFonts` | bool | false | Font embedding in PDF (future) |
| `RasterImagesOnly` | bool | true | Only raster images in PDF (future) |
| `CompressPdf` | bool | true | PDF compression (future) |
| `IncludeMetadata` | bool | true | PDF metadata (title, author, etc.) |
| `Title` / `Author` / `Subject` / `Keywords` | string? | null | PDF metadata overrides |
| `HighQualityAntialiasing` | bool | true | Enable HQ antialiasing |
| `LcdTextRendering` | bool | false | LCD-optimized text (future) |
| `SubpixelText` | bool | true | Subpixel text positioning (future) |
| `MaxPages` | int | 0 | Max pages (0 = unlimited) |
| `IncludePageNumbers` | bool | false | Draw page numbers |
| `PageNumberFormat` | string | "Page {0} of {1}" | Page number template |
| `PageNumberFontSizePt` | double | 8 | Page number font size |
| `PngCompressionLevel` | int | 100 | PNG quality (0-100) |
| `CropToPrintArea` | bool | false | Crop PNG to PrintArea (future) |
| `FilePrefix` | string | "page" | Output file prefix |
| `OutputDirectory` | string? | null | Output directory |

---

## Known Limitations

| Issue | Status |
|:------|:-------|
| PDF vector output | Current implementation uses staged bitmap approach. Future: render directly to PDF canvas for selectable/searchable text and vector borders/shapes. |
| PrintArea clipping | PageRenderer uses worksheet content bounds. Future: accept `PrintAreaInfo` for exact PrintArea clipping. |
| Quality settings (`ApplyQualitySettings`) | Method is a placeholder — SkiaSharp quality defaults are used. |
| PDF CreationDate/ModifiedDate | Removed due to API differences in SkiaSharp 3.x. Metadata still includes Title, Author, Subject, Keywords, Producer. |
| Multi-page overlays | Cluster overlays from PreviewGenerator/PdfGenerator fallback are single-page only. |
| Dead ExportOptions | `CropToPrintArea`, `EmbedFonts`, `CompressPdf`, `RasterImagesOnly`, `LcdTextRendering`, `SubpixelText` are defined but not yet wired to rendering logic. |

---

## Files Summary

### New Files

```
ExcelAPI/ExcelAPI/Rendering/
├── ExportOptions.cs          # Configurable export settings
├── PageRenderer.cs           # Multi-page splitting engine
└── ExportCoordinator.cs      # Public export API (PNG + PDF)
```

### Modified Files

```
ExcelAPI/ExcelAPI/Generators/
├── PreviewGenerator.cs       # Delegates to ExportCoordinator.ExportPng()
└── PdfGenerator.cs           # Delegates to ExportCoordinator.ExportPdf()

ExcelAPI/ExcelAPI/Rendering/
└── RenderingContext.cs        # Workbook/Sheet: init → set

ExcelAPI/
└── Program.cs                 # DI registrations (ExportOptions, PageRenderer, ExportCoordinator)
```

---

## Roadmap After Phase 11G

```
✅ Phase 11A — Rendering Core
✅ Phase 11B — Rendering Pipeline
✅ Phase 11C — Text Engine
✅ Phase 11D — Coordinate & Print Layout
✅ Phase 11E — Style Resolution & Theme Engine
✅ Phase 11F — Image & Shape Engine
✅ Phase 11G — Production Export Engine
────────────────────────────────────
⬜ Phase 11H — Regression & Pixel-Diff Validation
⬜ Phase 11I — Form Runtime Engine
⬜ Phase 11J — Field Overlay Engine
⬜ Phase 11K — Designer Enhancements
⬜ Phase 11L — Production Release
```

The Rendering Engine is now feature-complete. The next milestone (11H) will validate rendering accuracy against Excel-generated output.
