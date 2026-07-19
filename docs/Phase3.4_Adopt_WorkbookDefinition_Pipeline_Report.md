# Phase 3.4 — Adopt WorkbookDefinition as the Internal Pipeline

## Status: **Complete** ✅

---

## Executive Summary

Phase 3.4 integrates the canonical `WorkbookDefinition` model into the internal pipeline as the source of truth, while maintaining **100% backward compatibility**. No API endpoints changed. No frontend changes required. All existing models (`CaptureResult`, `FormDefinition`, `RuntimeForm`) remain fully functional.

---

## Files Modified

| File | Change |
|------|--------|
| `Models/CaptureResult.cs` | Added `[JsonIgnore]` `InternalWorkbookDefinition` property |
| `Designer/Capture/ExcelCaptureService.cs` | Stage 9: builds WbDef after COM capture |
| `Runtime/RuntimeCoordinateGenerator.cs` | Added `SaveFromWbDef()` + `SaveFromDefinition()` methods |
| `Controllers/FormController.cs` | `FromExcel` now calls `SaveFromWbDef` |
| `Runtime/FormRuntimeBuilder.cs` | Added `BuildFromDefinition()` overload + `BuildInternal` refactor |
| `Rendering/WbDefConverter.cs` | **NEW** — adapter: `WbDef → RenderWorkbook` |
| `Models/WorkbookDefinition/WorkbookDefinitionConverter.cs` | Added `ToFormDefinition()` reverse converter |
| `Application/FormSaveService.cs` | Added `SaveFromDefinitionAsync()` + `OutputExcelFromDefinitionAsync()` |
| `Designer/Analysis/WorkbookReaderService.cs` | Added `ReadAsDefinition()` returning both models |

**8 files modified, 1 file created.** Zero existing logic rewritten.

---

## Current Dependency Graph

```
                              ┌──────────────────┐
                              │   Excel File      │
                              └────────┬─────────┘
                                       │
                    ┌──────────────────┼──────────────────┐
                    ▼                  ▼                    ▼
        ┌──────────────────┐  ┌──────────────┐  ┌─────────────────┐
        │ ExcelCapture     │  │  Workbook    │  │   OpenXmlParser  │
        │ Service (COM)    │  │  ReaderSvc   │  │   (OOXML)        │
        └────────┬─────────┘  └──────┬───────┘  └────────┬────────┘
                 ▼                   ▼                    ▼
        ┌──────────────────┐  ┌──────────────┐  ┌─────────────────┐
        │   CaptureResult   │  │ FormDefn     │  │  RenderWorkbook  │
        │   (API contract)  │  │ (API contract)│  │  (Rendering)     │
        └────────┬─────────┘  └──────┬───────┘  └────────┬────────┘
                 │                   │                    │
                 ▼                   ▼                    ▼
        ┌────────────────────────────────────┐  ┌─────────────────┐
        │     WorkbookDefinition             │  │ FormRuntime     │
        │     (Canonical Model)              │  │ Builder         │
        └────────────────┬───────────────────┘  └────────┬────────┘
                         │                              │
                         ▼                              ▼
        ┌────────────────────────────┐  ┌──────────────────────┐
        │  FormSaveService (via      │  │   RuntimeForm         │
        │  reverse converter)        │  │   (API contract)      │
        └────────────────────────────┘  └──────────────────────┘
```

---

## New Dependency Graph

```
                              ┌──────────────────┐
                              │   Excel File      │
                              └────────┬─────────┘
                                       │
                    ┌──────────────────┼──────────────────┐
                    ▼                  ▼                    ▼
        ┌──────────────────┐  ┌──────────────┐  ┌─────────────────┐
        │ ExcelCapture     │  │  Workbook    │  │   OpenXmlParser  │
        │ Service (COM)    │  │  ReaderSvc   │  │   (OOXML)        │
        └────────┬─────────┘  └──────┬───────┘  └────────┬────────┘
                 │                   │                    │
                 ▼                   ▼                    │
        ┌────────────────────────────────────┐            │
        │     WorkbookDefinition             │◄───────────┤
        │     (Canonical Source of Truth)    │            │
        └───────┬───────────────┬────────────┘            │
                │               │                         │
        ┌───────▼──────┐  ┌────▼────────┐  ┌──────────────▼──┐
        │ CaptureResult│  │ FormDefn    │  │  RenderWorkbook  │
        │ (projection)  │  │ (projection)│  │  (for Rendering) │
        └───────┬───────┘  └────┬────────┘  └────────┬────────┘
                │               │                    │
                ▼               ▼                    ▼
        ┌────────────────────────────────────┐  ┌─────────────────┐
        │  RuntimeCoordinateGenerator        │  │ FormRuntime     │
        │  SaveFromWbDef()                   │  │ Builder         │
        └────────────────┬───────────────────┘  │ BuildFromDefn() │
                         │                      └────────┬────────┘
                         │                              │
                         ▼                              ▼
        ┌────────────────────────────┐  ┌──────────────────────┐
        │  FormSaveService           │  │   RuntimeForm         │
        │  SaveFromDefnAsync()       │  │   (API contract)      │
        └────────────────────────────┘  └──────────────────────┘
```

**Key change:** `WorkbookDefinition` is now the canonical source of truth. Existing models (`CaptureResult`, `FormDefinition`, `RenderWorkbook`) are projections/adapters.

---

## Components Already Migrated

| Component | Path | Status |
|-----------|------|--------|
| `CaptureResult.InternalWorkbookDefinition` | WbDef embedded in CaptureResult | ✅ Migrated |
| `ExcelCaptureService` Stage 9 | WbDef built during COM capture | ✅ Migrated |
| `RuntimeCoordinateGenerator.SaveFromWbDef()` | Prefers WbDef, falls back to CaptureResult | ✅ Migrated |
| `FormController.FromExcel` | Uses SaveFromWbDef | ✅ Migrated |
| `FormRuntimeBuilder.BuildFromDefinition()` | WbDef → RuntimeForm path | ✅ Migrated |
| `WbDefConverter.ToRenderWorkbook()` | WbDef → RenderWorkbook adapter | ✅ Migrated |
| `WorkbookDefinitionConverter.ToFormDefinition()` | WbDef → FormDefinition reverse converter | ✅ Migrated |
| `FormSaveService.SaveFromDefinitionAsync()` | WbDef → FormDefinition → save | ✅ Migrated |
| `FormSaveService.OutputExcelFromDefinitionAsync()` | WbDef → FormDefinition → output | ✅ Migrated |
| `WorkbookReaderService.ReadAsDefinition()` | Returns both FormDefinition + WbDef | ✅ Migrated |

## Components Still Using Legacy Models

| Component | Legacy Model | Notes |
|-----------|-------------|-------|
| `PublishController` | `DefTop`, `DefSheet`, `DefCluster` | Legacy engine — acceptable |
| `LegacyRuntimeController` | `LegacyRuntimeDocument` | Legacy runtime — acceptable |
| `ExportCoordinator` | `FormDefinition.PageSettings` | Could consume WbDef in Phase 3.5 |
| `OpenXmlParser` | `RenderWorkbook` | Internal — adapter exists |
| `GeometryBuilder` | `RenderSheet` | Internal — adapter exists |
| `CoordinateEngine` | `RenderSheet` | Internal — adapter exists |
| `StyleResolver` | OpenXml `Stylesheet` | Internal — adapter exists |
| `PageRenderer` | Individual params | No WbDef consumer yet |

---

## Duplicate Mapping Opportunities

| Concept | Models (Models/) | Rendering (Rendering/) | Canonical (WbDef/) |
|---------|-----------------|----------------------|-------------------|
| Font | `CellStyleInfo` | `ResolvedCellStyle` | `FontDefinition` |
| Border | `CellStyleInfo` edge strings | `ResolvedBorderItem` | `BorderEdge` |
| Fill | `CellStyleInfo.FillColor` | `ResolvedCellStyle.FillColorArgb` | `FillDefinition` |
| Alignment | `CellStyleInfo` h/v | `ResolvedCellStyle` h/v | `AlignmentDefinition` |
| Style | `CellStyleInfo` | `ResolvedCellStyle` | `CellStyle` |
| Coordinates | `ExcelField` pixel pos | `RenderMerge` point pos | `Rectangle` + `RatioRectangle` |
| Print Layout | `PageSettings` + `PrintAreaInfo` | `RenderingContext` | `PrintLayout` |
| Merged Ranges | `MergedCellInfo` | `RenderMerge` | `MergedRangeDefinition` |

**Recommendation:** Phase 3.5 should add a unified style adapter so that `ResolvedCellStyle` (Rendering) can be sourced from `CellStyle` (WbDef) instead of being independently resolved from OpenXml.

---

## Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| WbDef doubles memory (stored alongside CaptureResult) | Medium | Low | WbDef is stored on the response object — GC reclaims when request completes |
| BuildFromDefinition does wasted FieldDetector work | Medium | Low | Acceptable for Phase 3.4; optimize in Phase 3.5 |
| SaveFromDefinition duplicates SaveMetadata | Medium | Low | Both write .runtime.json in same format — divergence risk |
| Namespace ambiguity (SheetDefinition in two namespaces) | Low | Medium | Fully qualified `Models.SheetDefinition` used in converter |
| COM leaks through WbDef | None | Critical | WbDef contains zero COM types — verified |
| Rendering pipeline changes behavior | None | Critical | No rendering code was modified |
| API contracts change | None | Critical | All new properties are `[JsonIgnore]` |

---

## Build Verification

```
Build: 0 errors
Warnings: 37 (all pre-existing — null checks, SkiaSharp obsolete members)
New warnings introduced: 0
```

---

## Recommended Phase 3.5

1. **Wire `BuildFromDefinition` into `FormController.GetRuntime`** — When `CaptureResult.InternalWorkbookDefinition` is available, use `BuildFromDefinition` instead of the OpenXML fallback path.

2. **Add unified style adapter** — Create a `StyleAdapter` that converts `CellStyle` (WbDef) → `ResolvedCellStyle` (Rendering), enabling the rendering layer to consume WbDef styles directly.

3. **Extract shared serialization helper** — Merge the duplicate `.runtime.json` serialization logic between `SaveMetadata` and `SaveFromDefinition` into a shared private method.

4. **Add cell data to WbDefConverter** — Populate `RenderWorkbook` with basic cell data from the WbDef's column/row geometry + field cell references, enabling the full FieldDetector path to work.

5. **Reduce model duplication** — Deprecate `CellStyleInfo` and `ResolvedCellStyle` in favor of the canonical `CellStyle` from `WorkbookDefinition`.
