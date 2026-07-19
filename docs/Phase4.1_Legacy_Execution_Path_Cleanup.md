# Phase 4.1 — Eliminate Remaining Legacy Execution Paths

**Date**: 2026-07-19  
**Status**: ✅ Complete  
**Build**: 0 errors, 0 new warnings  

---

## Summary

Completed the transition to the canonical architecture by removing the final legacy execution paths that are no longer in use. WorkbookDefinition is now the only internal execution model, with legacy models existing only as thin compatibility projections for public APIs.

---

## Files Modified

| File | Change |
|------|--------|
| `Rendering/WbDefConverter.cs` | Removed `ToRenderWorkbook()` method (~70 lines). Updated header comment. `ToResolvedCellStyle` and `ToPrintLayoutResult` remain. |
| `Controllers/FormController.cs` | Added `[Obsolete]` attribute to `ConvertCaptureToForm()` — WorkbookDefinition is the canonical path. |
| `Runtime/FieldDetector.cs` | Removed unused `ptsToPx = dpi / 72.0` declaration (confirmed zero references in the method). |
| `Models/WorkbookDefinition/WorkbookDefinitionConverter.cs` | Replaced inline `300.0 / 72.0` with canonical `Rectangle.PtToPx(1, 300)`. |
| `Application/CoordinateTransformer.cs` | Replaced inline `dpi / 72.0` fallback with canonical `WbDef.Rectangle.PtToPx(1, dpi)`. |
| `Validation/PipelineValidator.cs` | Updated `ValidateRenderingBridge` to validate WbDef properties directly instead of via removed `ToRenderWorkbook` adapter. |

---

## Dead Code Eliminated

| Code Removed | Reason |
|-------------|--------|
| `WbDefConverter.ToRenderWorkbook()` | Only used by PipelineValidator (test code). Rendering now consumes WbDef directly. |
| `FieldDetector.ptsToPx` declaration | Unused variable — `CoordinateEngine.GetCellOrMergePixelBounds()` already returns pixel values. |
| `CoordinateTransformer` inline `dpi / 72.0` | Replaced with canonical `WbDef.Rectangle.PtToPx(1, dpi)`. |

---

## Code Marked for Removal (Next Phase)

| Code | Marker | Action in Phase 4.2 |
|------|--------|---------------------|
| `ConvertCaptureToForm()` | `[Obsolete]` | Remove when WbDef convergence confirmed for all templates. |

---

## Remaining Legacy Compatibility Code

The following remain for the OpenXml fallback path (legacy templates without `.runtime.json` metadata):

- `FormRuntimeBuilder.Build(RenderWorkbook)` — OpenXml fallback
- `FormRuntimeBuilder.BuildInternal()` — called by `Build()` above
- `FormController.GetRuntime()` — retains OpenXmlParser fallback path
- `ConvertCaptureToForm()` — marked `[Obsolete]`, retained as fallback

These cannot be removed until all legacy templates have `.runtime.json` metadata (post-migration).

---

## Updated Dependency Graph

```
Excel Workbook
        │
        ▼
  WorkbookDefinition
        │
 ├─────────────┬──────────────┐
 │             │              │
 ▼             ▼              ▼
Runtime    Rendering       Save/Export
 │             │              │
 ▼             ▼              ▼
RuntimeForm    PNG/PDF       Excel
        │
        ▼
    API DTOs
        │
        ▼
    Frontend
```

**No internal dependency on**: `CaptureResult`, `FormDefinition`, `RenderWorkbook` (except as compatibility projections for public APIs).

---

## Coordinate Math Cleanup

| Location | Before | After |
|----------|--------|-------|
| `WorkbookDefinitionConverter.cs` | `ptsToPx = 300.0 / 72.0` | `Rectangle.PtToPx(1, 300)` |
| `CoordinateTransformer.cs` | `dpi / 72.0` | `WbDef.Rectangle.PtToPx(1, dpi)` |
| `FieldDetector.cs` | `ptsToPx = dpi / 72.0` | (removed — unused) |

Remaining inline `dpi/72.0` in `Designer/Capture` and `Designer/Generation` are in COM-capture and overlay-generator code (slated for Phase 4.2 removal) — not in the primary execution pipeline.

---

## DI Registration Audit

All registered services are still referenced in production code. No dead DI registrations to remove.

---

## Readiness for Phase 4.2

**Ready** ✅

The codebase is now ready for Phase 4.2 — **Native WorkbookDefinition Rendering**. The remaining work includes:

1. Replace Designer overlay generators (`PreviewGenerator`, `PdfGenerator`) with WbDef-native alternatives
2. Simplify `FormRuntimeBuilder.Build(RenderWorkbook)` to always require WbDef
3. Remove `[Obsolete]` `ConvertCaptureToForm()` and the OpenXml fallback in `GetRuntime()`
4. Consolidate remaining Designer/Generation legacy code
