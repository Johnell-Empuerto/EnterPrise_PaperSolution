# Phase 11B — Rendering Pipeline Integration

**Date:** July 9, 2026  
**Status:** ✅ Complete — Build: 0 errors, 15 warnings  
**Location:** `ExcelAPI/ExcelAPI/Rendering/` + `Generators/`  

---

## Overview

Phase 11B refactors the FormLess Rendering Core into a modular, pluggable pipeline and integrates it into the existing PreviewGenerator and PdfGenerator.

### Architecture Before

```
OpenXmlParser (parses + computes geometry)
  ↓
FillEngine / BorderEngine (standalone, no shared context)
  ↓  
RendererCoordinator (hardcoded layer order)
  ↓
PreviewGenerator / PdfGenerator (duplicate rendering logic)
```

### Architecture After

```
OpenXmlParser → only parses XLSX
  ↓
GeometryBuilder → computes geometry (separated from parser)
  ↓
RenderingContext → shared context (workbook, sheet, origin, DPI)
  ↓
IRenderLayer pipeline (plug-and-play layers):
  FillEngine (IRenderLayer)     → Layer 1: backgrounds
  GridlineLayer (IRenderLayer)  → Layer 2: gridlines
  BorderEngine (IRenderLayer)   → Layer 3: borders
  TextLayer (future)            → Layer 4: text
  ImageLayer (future)           → Layer 5: images
  ↓
RendererCoordinator → foreach(layer) layer.Render(canvas, context)
  ↓
PreviewGenerator / PdfGenerator → delegates to RendererCoordinator
```

---

## Files Created (3 new)

### `Rendering/RenderingContext.cs`
Shared context passed to every IRenderLayer:
- `Workbook` — parsed workbook model
- `Sheet` — current sheet
- `OriginXPt`, `OriginYPt` — printable origin in points
- `PageWidthPt`, `PageHeightPt` — page dimensions
- `Dpi` — rendering DPI (300)
- `PointsToPixels` — computed multiplier (`Dpi / 72 = 4.1667`)
- `PixelWidth`, `PixelHeight` — canvas pixel dimensions

### `Rendering/IRenderLayer.cs`
Pluggable layer interface:
```csharp
public interface IRenderLayer
{
    void Render(SKCanvas canvas, RenderingContext context);
    string Name { get; }
}
```
Future layers (TextEngine, ImageEngine, ShapeEngine) simply implement this interface and register in DI — no coordinator changes needed.

### `Rendering/GeometryBuilder.cs`
Geometry computation extracted from OpenXmlParser:
- `ComputeGeometry(RenderSheet)` — cumulative column/row positions, merge bounds
- `CharWidthToPoints()` — delegates to `CellGeometryEngine` as canonical source (§4)

---

## Files Modified (7 updated)

### `Rendering/OpenXmlParser.cs`
- **Removed**: `ComputeGeometry()` method and `Padding`/`ScreenDpi`/`PointsPerInch` constants
- **Added**: Optional `GeometryBuilder` constructor parameter
- **Kept**: `ComputeGeometry` as public convenience method that delegates to `GeometryBuilder`
- Parser now only parses — geometry computation lives in `GeometryBuilder`

### `Rendering/FillEngine.cs`
- **Added**: `IRenderLayer` implementation
- `Render(context)` delegates to existing `RenderFills()` using `context.Sheet` + `context.OriginXPt/YPt`
- `Name = "FillLayer"`

### `Rendering/BorderEngine.cs`
- **Added**: `IRenderLayer` implementation
- `Render(context)` delegates to existing `RenderBorders()` via shared context
- `Name = "BorderLayer"`

### `Rendering/CellGeometryEngine.cs`
- `CharWidthToPoints()` is now the **canonical implementation** (§4 compliance)
- Other classes delegate here — no formula duplication

### `Rendering/RendererCoordinator.cs` (rewritten)
- **New**: `Prepare()` method — shared setup for PNG/PDF: parses, computes geometry, creates `RenderingContext`
- **New**: `RenderToCanvas()` — executes the IRenderLayer pipeline via `foreach(layer) layer.Render(canvas, context)`
- `RenderToPng()` and `RenderToPdf()` both call `Prepare()` + `RenderToCanvas()`
- Constructor takes `IEnumerable<IRenderLayer>` — DI injects all registered layers

### `Generators/PreviewGenerator.cs` (rewritten)
- **New**: Delegates to `RendererCoordinator.RenderToPng()` when XLSX path is available
- **Fallback**: Overlay-only preview for cases without XLSX
- Removed: `DrawPrintArea()` — no longer duplicates grid/print-rendering logic
- `DrawClusterOverlay()` kept for field annotation overlays

### `Generators/PdfGenerator.cs` (rewritten)
- **New**: Delegates to `RendererCoordinator.RenderToPdf()` when XLSX path is available
- **Fallback**: Overlay-only PDF for cases without XLSX
- Removed: `DrawPrintAreaContent()` — no longer duplicates cell rendering
- `DrawCluster()` kept for field annotation overlays

### `Program.cs`
- **Added**: `GeometryBuilder`, `GridlineLayer`, `IRenderLayer` DI registrations
- **Registration order** determines render order:
  1. `FillEngine` (fills)
  2. `GridlineLayer` (gridlines)
  3. `BorderEngine` (borders)
- `RendererCoordinator` gets `IEnumerable<IRenderLayer>` auto-injected

### `Generators/GridlineLayer.cs` (new)
- Extracted gridline rendering from `RendererCoordinator.RenderGridlines()`
- Implements `IRenderLayer`
- Layer 2 in the pipeline (between fills and borders)

---

## Requirement Compliance

| Req # | Requirement | Status |
|:-----:|:------------|:------:|
| 1 | Refactor parser: only parse, no rendering logic | ✅ Geometry extracted to GeometryBuilder |
| 2 | Integrate RendererCoordinator into PreviewGenerator | ✅ Delegates rendering, keeps overlays |
| 3 | Integrate RendererCoordinator into PdfGenerator | ✅ Delegates rendering, keeps overlays |
| 4 | Single source for CharWidthToPoints | ✅ CellGeometryEngine is canonical |
| 5 | Create RenderingContext | ✅ All engines receive same context |
| 6 | Create IRenderLayer interface | ✅ + pluggable pipeline in coordinator |
| 7 | Remove rendering from ExcelCaptureService | ✅ Already delegates to COM + PDFtoImage |
| 8 | Update Preview API pipeline | ✅ Pipeline matches spec |
| 9 | Maintain Next.js frontend | ✅ No API changes |
| 10 | Prepare for TextEngine (M4) | ✅ IRenderLayer ready, no coordinator changes needed |

---

## Build Validation

| Metric | Result |
|:-------|:------:|
| Build errors | **0** ✅ |
| Build warnings | **15** (minor: unused field warnings, CA suggestions) |
| Breaking API changes | **None** — DI resolves new constructor params automatically |
| Frontend compatibility | ✅ No changes to Next.js or API contracts |

---

## Remaining Items (M4 Preparation)

| Item | Status |
|:-----|:-------|
| TextEngine as IRenderLayer | 🔜 Milestone M4 |
| StyleResolver as separate service | 🔜 Future — ApplyCellStyle still in parser |
| Vector PDF output (not bitmap-wrapped) | 🔜 Future milestone |
| Theme color resolution | 🔜 Future — currently returns null |
