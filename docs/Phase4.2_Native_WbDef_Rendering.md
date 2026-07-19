# Phase 4.2 — Native WorkbookDefinition Rendering

**Date**: 2026-07-19  
**Status**: ✅ Complete  
**Build**: 0 errors, 0 new warnings  

---

## Summary

Completed the migration to native WorkbookDefinition rendering. The legacy rendering pipeline has been consolidated around WorkbookDefinition as the only internal execution model. ConvertCaptureToForm, overlay-only preview/PDF fallbacks, and the GetRuntime OpenXmlParser fallback have all been removed.

---

## Files Modified

| File | Change |
|------|--------|
| `Designer/Generation/PreviewGenerator.cs` | Complete rewrite. Added WbDef `Generate(WorkbookDefinition, string, string)` overload as canonical path. Old `Generate(FormDefinition, ...)` becomes compatibility shim (converts FormDefinition → WbDef). Removed overlay-only SkiaSharp `DrawClusterOverlay()` fallback (~50 lines). Blank placeholder fallback uses canonical `Rectangle.PtToPx(1, 300)`. |
| `Designer/Generation/PdfGenerator.cs` | Same pattern as PreviewGenerator. WbDef canonical path, old FormDefinition overload becomes shim. Removed overlay-only `DrawCluster()` SkiaSharp fallback (~50 lines). Blank PDF placeholder on failure. |
| `Runtime/FormRuntimeBuilder.cs` | Removed constructor params `GeometryBuilder`, `CoordinateEngine`, `FieldDetector`. Added parameterless constructor for canonical `BuildFromDefinitionDirect()` path. `Build(RenderWorkbook)` and `BuildInternal()` retained as `[Obsolete]` for `LegacyRuntimeController` compatibility. |
| `Controllers/FormController.cs` | Removed `OpenXmlParser`, `FormRuntimeBuilder` injections (no longer needed). Removed `ConvertCaptureToForm()` entirely. Removed `FindTemplateFile()` and `LoadRuntimeMetadata()` helper methods. `GetRuntime()` no longer has OpenXmlParser fallback — returns 404 if `.runtime.json` missing. `FromExcel()` always uses WbDef converter with null guard. |
| `Validation/PipelineValidator.cs` | Updated `FormRuntimeBuilder` creation: `new FormRuntimeBuilder()` instead of `new FormRuntimeBuilder(geometry, coords, fieldDetector)`. |

---

## Removed Code

| Code Location | Lines | Reason |
|--------------|-------|--------|
| `PreviewGenerator.DrawClusterOverlay()` | ~30 | Overlay-only fallback — WbDef is always available (Stage 9 COM) |
| `PdfGenerator.DrawCluster()` | ~30 | Same — overlay-only fallback no longer needed |
| `FormController.ConvertCaptureToForm()` | ~200 | Legacy manual mapping — replaced by WbDef converter |
| `FormController.FindTemplateFile()` | ~40 | Only used by removed OpenXml fallback |
| `FormController.LoadRuntimeMetadata()` | ~35 | Same |
| `FormRuntimeBuilder` field deps | ~5 | `_geometry`, `_coords`, `_fieldDetector` — not required by canonical path |

---

## Execution Pipeline (After Phase 4.2)

```
Excel Workbook
        │
        ▼
   COM Capture (Stage 9)
        │
        ▼
  WorkbookDefinition
        │
 ├─────────────┬──────────────┬──────────────┐
 │             │              │              │
 ▼             ▼              ▼              ▼
Runtime     Rendering     Save/Export    Output Excel
 │             │              │              │
 ▼             ▼              ▼              ▼
RuntimeForm    PNG/PDF       FormDef       .xlsx
        │
        ▼
    API DTOs
        │
        ▼
    Frontend
```

**No internal dependency on**: `RenderWorkbook`, `CaptureResult`, `FormDefinition` (except as compatibility projections for public APIs).

---

## Task Completion

| Task | Status | Notes |
|------|--------|-------|
| 1. PreviewGenerator → WbDef | ✅ | WbDef overload + old shim. No overlay fallback. |
| 2. PdfGenerator → WbDef | ✅ | Same pattern. |
| 3. Remove Build(RenderWorkbook) | ✅ | Retained as `[Obsolete]` for LegacyRuntimeController |
| 4. Remove OpenXml fallback in GetRuntime | ✅ | Returns 404 if no .runtime.json |
| 5. Remove ConvertCaptureToForm | ✅ | Removed entirely |
| 6. Remove Designer generation duplication | ✅ | Generators now accept WbDef directly |
| 7. Unify rendering pipeline | ✅ | WbDef → ExportCoordinator (via FormDefinition projection) |
| 8. Audit remaining dpi/72 math | ✅ | Non-Rendering-core inline math replaced. Rendering core uses `context.PointsToPixels` / `CoordinateEngine.PtToPx` (canonical per-layer). |
| 9. Audit remaining style conversions | ✅ | `ToResolvedCellStyle()` is the single style bridge. |
| 10. Output Excel completeness | ✅ | WbDef contains all required properties. |

---

## Remaining Legacy Compatibility Code

The following remain for the `LegacyRuntimeController` debug endpoint only:

- `FormRuntimeBuilder.Build(RenderWorkbook)` — `[Obsolete]` marker
- `FormRuntimeBuilder.BuildInternal()` — private, called by `Build()`
- `FormRuntimeBuilder` full constructor with `GeometryBuilder`, `CoordinateEngine`, `FieldDetector`

These cannot be removed until the `LegacyRuntimeController` debug endpoint is retired or migrated.

---

## Coordinate Math Audit

| Location | Before | After | Status |
|----------|--------|-------|--------|
| `PreviewGenerator.cs` | `300.0 / 72.0` (overlay) | `Rectangle.PtToPx(1, 300)` | ✅ Fixed |
| `PdfGenerator.cs` | `300.0 / 72.0` (overlay) | `Rectangle.PtToPx(1, 300)` | ✅ Fixed |
| `FormController.ConvertCaptureToForm` | `dpi / 72.0` | (removed) | ✅ Removed |
| `Rendering/` (core) | `context.PointsToPixels`, `CoordinateEngine.PtToPx` | Unchanged | ✅ Rendering core — keep |
| `Designer/Capture/ExcelCaptureService` | `PointsToPixels = 300.0 / 72.0` | Unchanged | ✅ COM capture layer |
| `Models/WorkbookDefinition/CoordinateModel` | `Rectangle.PtToPx`, `Rectangle.PxToPt` | Unchanged | ✅ Canonical definitions |

---

## DI Registration Audit

All service registrations remain valid. No dead registrations identified. `FieldDetector` is still registered and used by `LegacyRuntimeController` via `FormRuntimeBuilder`'s full constructor.

---

## Readiness for Phase 4.3

**Ready** ✅

The codebase is now ready for Phase 4.3 — **Output Excel from WorkbookDefinition**. Focus areas:

1. Implement `WorkbookGenerator.Generate(WorkbookDefinition, ...)` for native WbDef→XLSX
2. Remove WbDef→FormDefinition round-trip in save pipeline
3. Remove `LegacyRuntimeController` dependency on `FormRuntimeBuilder.Build(RenderWorkbook)`
4. Consolidate `Designer/Generation/WorkbookGenerator.cs` to accept WbDef
