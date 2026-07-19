# Phase 4.0 — Legacy Pipeline Retirement (Safe Cleanup) ✅

## Status: ✅ Complete

**Build: 0 errors, 0 new warnings**

---

## Summary

Phase 4.0 removes the remaining legacy execution pipeline from the backend. After this phase, `WorkbookDefinition` is the **only** internal source of truth for the application. Legacy models (`SaveMetadata`, `BuildFromDefinition`, manual `ConvertCaptureToForm`) are removed.

**No API contracts changed. No frontend changes required. No rendering output changed.**

---

## What Was Removed

### 1. `SaveMetadata()` — Legacy Runtime Persistence

**File:** `RuntimeCoordinateGenerator.cs`

**Removed:** ~90 lines of duplicate serialization logic that wrote `.runtime.json` from `CaptureResult.Fields` directly.

**Impact:** `SaveFromWbDef()` now requires `InternalWorkbookDefinition` to be populated. If it's null (should never happen for new uploads — Stage 9 always populates it), a warning is logged and the method returns early.

**Rationale:** `SaveFromDefinition()` (the WbDef path) produces the same `.runtime.json` schema with richer data (sheet names, field styles, multi-sheet support). Phase 3.7 validated equivalence.

### 2. `BuildFromDefinition()` — Legacy Adapter Path

**File:** `FormRuntimeBuilder.cs`

**Removed:** ~90 lines of adapter logic that converted WbDef→RenderWorkbook→BuildInternal→field override.

**Impact:** `BuildFromDefinitionDirect()` is the canonical path. `Build(RenderWorkbook)` and `BuildInternal()` are retained for the legacy OpenXml fallback path in `GetRuntime`.

**Rationale:** The adapter path wasted computation (empty adapter sheet → geometry compute → FieldDetector → clear fields → repopulate). Phase 3.7 validated that `BuildFromDefinitionDirect` produces identical output.

### 3. Manual `ConvertCaptureToForm()` — FormDefinition Projection (de-emphasized)

**File:** `FormController.cs`

**Changed:** The `FromExcel` endpoint now uses `WorkbookDefinitionConverter.ToFormDefinition()` as the primary path. The manual `ConvertCaptureToForm()` method is retained as a fallback when `InternalWorkbookDefinition` is null.

**Added:** `ApplyFormDefinitionIO()` — extracts the IO operations (thumbnail generation, background image persistence) from `ConvertCaptureToForm` into a separate method that's only called when using the WbDef converter path. This prevents double-processing.

**Impact:** Same FormDefinition output, cleaner separation of concerns.

---

## Inline Coordinate Math Consolidation

| File | Before | After |
|------|--------|-------|
| `RuntimeCoordinateGenerator.SaveFromDefinition` | `field.BoundsPt.Left * (dpi / 72.0)` | `WbDef.Rectangle.PtToPx(field.BoundsPt.Left, dpi)` |
| `RuntimeCoordinateGenerator.SaveFromDefinition` | `Math.Round(dpi / 72.0, 6)` | `Math.Round(WbDef.Rectangle.PtToPx(1, dpi), 6)` |
| `FormRuntimeBuilder.BuildFromDefinitionDirect` | `double ptsToPx = dpi / 72.0` | `WbDef.Rectangle.PtToPx(pageWidthPt, dpi)` |
| `FormRuntimeBuilder.BuildInternal` | `sheet.TotalWidthPt * ptsToPx` | `WbDef.Rectangle.PtToPx(sheet.TotalWidthPt, dpi)` |

**Remaining inline formulas (acceptable):**
- `RenderingContext.PointsToPixels` — centralized property, single authoritative source for rendering
- `CoordinateEngine.PtToPx` / `PxToPt` — canonical rendering conversions (delegate to same math)
- `ConvertCaptureToForm` fallback — rarely executes, documented as safe

---

## Architecture After Cleanup

```
Excel Workbook
        │
        ▼
   WorkbookDefinition          (single source of truth)
        │
        ├──────────────┬──────────────┬──────────────┐
        │              │              │              │
        ▼              ▼              ▼              ▼
   Runtime        Rendering        Save           API Response
   (WbDef→        (WbDef→          (WbDef→         (WbDef→
    RuntimeForm)   RenderWorkbook)  FormDefinition) FormDefinition
```

**Removed from execution pipeline:**
- ❌ `SaveMetadata()` — replaced by `SaveFromDefinition()`
- ❌ `BuildFromDefinition()` — replaced by `BuildFromDefinitionDirect()`
- ❌ Manual field mapping in `ConvertCaptureToForm()` — replaced by WbDef converter
- ❌ Inline `dpi/72.0` throughout Runtime/ and Controllers/ — replaced by `Rectangle.PtToPx()`

---

## Pipeline Validator Update

The `PipelineValidator` was updated to reflect the cleanup:
- **Runtime Pipeline (Axis 2):** No longer compares `BuildFromDefinition` vs `BuildFromDefinitionDirect` (the former no longer exists). Now validates `BuildFromDefinitionDirect` output against expected values from the test WbDef using `WbDef.Rectangle.PtToPx()` and `field.BoundsPt.ToPixels(300)`.

---

## Phase 4.1 Recommendations

1. **Remove `WbDefConverter.ToRenderWorkbook()`** — now production-dead (only used by PipelineValidator). Move the validator's rendering bridge test to a different approach.
2. **Mark `ConvertCaptureToForm` as `[Obsolete]`** — since `InternalWorkbookDefinition` is always populated by Stage 9, the fallback should never execute.
3. **Remove `Build(RenderWorkbook)` fallback path** from `GetRuntime` — once all legacy templates have `.runtime.json` files (migration completed), the OpenXml fallback becomes dead code.
4. **Remove `ServicesRegistration.cs` manual registrations** — simplify DI by removing no-longer-needed service registrations.
