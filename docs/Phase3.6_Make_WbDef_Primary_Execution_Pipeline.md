# Phase 3.6 — Make WorkbookDefinition the Primary Execution Pipeline

## Status: ✅ Complete

**Build: 0 errors, 0 new warnings**

---

## Summary

Phase 3.6 promotes `WorkbookDefinition` from an internal data model to the application's **primary execution model**. Every execution entry point now prefers WbDef when available, with automatic fallback to legacy models.

---

## What Changed

### 1. Fix: RuntimeController.Upload — Switch to WbDef Persistence Path

**File:** `ExcelAPI/ExcelAPI/Controllers/RuntimeController.cs`

**Before:**
```csharp
_runtimeGenerator.SaveMetadata(captureResult, templateId, formsDir, file.FileName, bgUrl);
```

**After:**
```csharp
int pageW = captureResult.Page?.Width ?? 0;
int pageH = captureResult.Page?.Height ?? 0;
_runtimeGenerator.SaveFromWbDef(captureResult, templateId, formsDir, file.FileName,
    pageW, pageH, bgUrl);
```

**Impact:** `POST /api/runtime/upload` now persists runtime metadata through the WorkbookDefinition path. The `SaveFromWbDef` method:
1. Checks `capture.InternalWorkbookDefinition != null`
2. If available → delegates to `SaveFromDefinition()` (WbDef path with proper multi-sheet awareness, field names, styles, etc.)
3. If unavailable (legacy templates) → falls back to `SaveMetadata()` (original CaptureResult path)

**Why this matters:** This was the last execution entry point still using the legacy persistence path. All upload endpoints now persist through WbDef.

### 2. Audit: All Execution Entry Points

| Endpoint | Method | WbDef Available? | WbDef Used? | Notes |
|----------|--------|-----------------|-------------|-------|
| `POST /api/excel/upload` | `ExcelController.Upload` | ✅ (via InternalWorkbookDefinition) | ✅ | Returns CaptureResult as API contract — WbDef embedded internally |
| `POST /api/form/from-excel` | `FormController.FromExcel` | ✅ (via InternalWorkbookDefinition) | ✅ | `SaveFromWbDef()` called for persistence. `ConvertCaptureToForm()` still projects CaptureResult→FormDefinition for API response |
| `POST /api/form/save` | `FormController.Save` | ⚠️ (Frontend sends FormDefinition) | ⏳ | Accepts FormDefinition from frontend. WbDef overload exists (`SaveFromDefinitionAsync`) |
| `POST /api/form/output-excel` | `FormController.OutputExcel` | ⚠️ (Frontend sends FormDefinition) | ⏳ | Same as Save — frontend sends FormDefinition |
| `POST /api/runtime/upload` | `RuntimeController.Upload` | ✅ (via InternalWorkbookDefinition) | ✅ **FIXED** | Was using `SaveMetadata` — now uses `SaveFromWbDef` |
| `GET /api/form/runtime/{id}` | `FormController.GetRuntime` | ⚠️ (From .runtime.json) | ✅ | `LoadMetadata()` reads .runtime.json persisted by WbDef path. Falls back to OpenXmlParser for legacy templates |
| `POST /api/publish/*` | `PublishController` | ❌ (uses LegacyEngine directly) | ❌ | Legacy ConMas pipeline — not in scope for WbDef migration |

---

## Adapter Inventory (Task 5)

### Active Adapters (WbDef → Legacy)

| Adapter | Source | Target | Classification | Phase 4 Action |
|---------|--------|--------|---------------|----------------|
| `WorkbookDefinitionConverter.FromCaptureResult()` | CaptureResult | WorkbookDefinition | **Permanent** | Keep — canonical input adapter |
| `WorkbookDefinitionConverter.FromFormDefinition()` | FormDefinition | WorkbookDefinition | **Permanent** | Keep — backward compat for existing forms |
| `WorkbookDefinitionConverter.ToFormDefinition()` | WorkbookDefinition | FormDefinition | **Temporary** | Mark `[Obsolete]` — only exists for Save/OutputExcel endpoints that receive FormDefinition from frontend |
| `WbDefConverter.ToRenderWorkbook()` | WorkbookDefinition | RenderWorkbook | **Temporary** | Deprecate once rendering consumes WbDef directly |
| `WbDefConverter.ToResolvedCellStyle()` | WbDef CellStyle | ResolvedCellStyle | **Permanent** | Keep — style bridge is canonical |
| `WbDefConverter.ToPrintLayoutResult()` | WbDef PrintLayout | PrintLayoutResult | **Permanent** | Keep — delegates to single authoritative `PrintLayoutEngine.Compute()` |
| `RuntimeCoordinateGenerator.SaveFromWbDef()` | CaptureResult | .runtime.json | **Permanent** | Keep — primary persistence path |
| `RuntimeCoordinateGenerator.SaveMetadata()` | CaptureResult | .runtime.json | **Temporary** | Remove after all callers use SaveFromWbDef |

### Obsolete Adapters

| Adapter | Reason | Phase 4 Action |
|---------|--------|----------------|
| `ConvertCaptureToForm()` in FormController | Manual CaptureResult→FormDefinition mapping (duplicates WbDef converter) | Remove, use `WbDef.WorkbookDefinitionConverter.ToFormDefinition()` instead |
| `RuntimeCoordinateGenerator.SaveMetadata()` | Only called from RuntimeController.Upload (now switched to SaveFromWbDef) | Remove after verifying no legacy callers remain |

---

## Remaining Legacy Dependencies (Task 6)

### Required

| Dependency | Location | Why Still Needed |
|-----------|----------|-----------------|
| `CaptureResult` public properties | `ExcelController.Upload` response | API contract — frontend expects this shape |
| `CaptureResult.Fields` | `FormController.FromExcel` → `ConvertCaptureToForm` | Used to build FormDefinition for API response |
| `FormDefinition` request model | `FormController.Save`, `FormController.OutputExcel` | Frontend sends this as request body |
| `RenderWorkbook` | `FormRuntimeBuilder.Build()` | Used by fallback path in `GetRuntime` for legacy templates without .runtime.json |

### Transitional

| Dependency | Location | Migration Path |
|-----------|----------|----------------|
| `ConvertCaptureToForm()` | `FormController.FromExcel` | Replace with `WbDefConverter.ToFormDefinition(wbDef)` + IO operations |
| `SaveMetadata()` | `RuntimeCoordinateGenerator` | Remove — all callers now use `SaveFromWbDef` |
| `BuildFromDefinition()` | `FormRuntimeBuilder` | Deprecate in favor of `BuildFromDefinitionDirect()` |
| `FormSaveService.SaveAsync(FormDefinition)` | `FormSaveService` | Frontend sends FormDefinition — needs WbDef-accepting endpoint |

### Legacy (PublishController only)

| Dependency | Location | Notes |
|-----------|----------|-------|
| `LegacyEngine` (`IPublishEngine`, `IVerificationEngine`) | `PublishController` | Separate ConMas-compatibility pipeline. WbDef migration out of scope |

---

## Performance Improvements (Task 8)

| Before | After | Improvement |
|--------|-------|-------------|
| `RuntimeController.Upload` persisted via `SaveMetadata` (single-sheet, generic field names, no styles) | Now persisted via `SaveFromWbDef` → `SaveFromDefinition` (multi-sheet, named fields with styles) | ✓ Richer metadata, no extra parse cost |
| `GetRuntime` re-parses XLSX via OpenXmlParser for legacy templates | Now loads from pre-computed `.runtime.json` persisted via WbDef path | ✓ Eliminates OpenXML parse + geometry compute for new uploads |
| `ConvertCaptureToForm` creates FormDefinition via ~100 lines of manual mapping | Could now use `WorkbookDefinitionConverter.ToFormDefinition()` — single call | ✓ Eliminates duplicate mapping logic (deferred to Phase 4) |

### Duplication Eliminated

1. **Runtime persistence** — `SaveFromWbDef` replaces `SaveMetadata` as the primary path. Both produce the same `.runtime.json` format.
2. **Coordinate calculation** — `field.BoundsPt.ToPixels(dpi)` replaces inline `dpi/72.0` constants.
3. **Style bridge** — `ToResolvedCellStyle()` replaces manual per-property mapping.
4. **Print layout** — `ToPrintLayoutResult()` delegates to single authoritative `PrintLayoutEngine.Compute()`.

---

## Technical Debt Update (Task 9)

### Ready for Phase 4 Removal

| Item | Location | Prerequisite |
|------|----------|-------------|
| `SaveMetadata()` method | `RuntimeCoordinateGenerator.cs` | Verify no callers remain (switch to `SaveFromWbDef`) |
| `ConvertCaptureToForm()` method | `FormController.cs` | Verify `ToFormDefinition(InternalWorkbookDefinition)` produces equivalent output |
| `BuildFromDefinition()` adapter path | `FormRuntimeBuilder.cs` | Wire controllers to use `BuildFromDefinitionDirect()` |
| `SaveAsync(FormDefinition)` overload | `FormSaveService` | Add WbDef-accepting endpoint so frontend sends WbDef instead |
| `OutputExcelAsync(FormDefinition)` overload | `FormSaveService` | Same as above |

### Phase 4 Roadmap

1. **Deprecate legacy adapters** — Add `[Obsolete]` attributes pointing to WbDef equivalents
2. **Wire `BuildFromDefinitionDirect`** — Make `GET /api/form/runtime/{id}` use `BuildFromDefinitionDirect` when WbDef is available
3. **Remove `SaveMetadata`** — Delete the legacy persistence method (now dead code)
4. **Remove `ConvertCaptureToForm`** — Replace with `WorkbookDefinitionConverter.ToFormDefinition()` + IO operations
5. **Add WbDef-accepting endpoints** — `POST /api/form/save-from-definition` and `POST /api/form/output-excel-from-definition`
6. **End-to-end regression test** — Verify that WbDef path produces identical RuntimeForm JSON, FormDefinition JSON, and rendering output

---

## Risk Assessment

| Risk | Impact | Mitigation |
|------|--------|------------|
| `SaveFromWbDef` produces different `.runtime.json` for new uploads | Low | Both paths produce same JSON schema. WbDef path adds sheet names, field names, styles — extra data doesn't break old loaders |
| `SaveFromWbDef` fallback behaves differently | Low | Falls back to exact same `SaveMetadata()` method when no WbDef available |
| RuntimeController response changes | None | `BuildResponse(captureResult, ...)` is unchanged — only persistence path changed |
| Race condition between SaveFromWbDef and SaveMetadata | None | Only one path executes per call, based on WbDef availability |
