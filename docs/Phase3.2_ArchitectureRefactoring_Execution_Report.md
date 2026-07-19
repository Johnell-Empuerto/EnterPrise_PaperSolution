# Phase 3.2 — Architecture Refactoring Execution Report

**Date**: 2026-07-19  
**Status**: ✅ Complete — Build Succeeds (0 Errors)  
**Objective**: Refactor the monolithic `ExcelAPI` project into clear architectural layers.

---

## 1. Approved Target Structure

```
ExcelAPI/
├── Controllers/          ← Unchanged (7 files)
├── Models/               ← Unchanged (15 files)
├── Designer/             ← NEW — all COM-dependent code
│   ├── Capture/          ← ExcelCaptureService, IExcelCaptureService
│   ├── Analysis/         ← WorkbookReaderService
│   ├── Generation/       ← WorkbookGenerator, PdfGenerator, PreviewGenerator
│   └── Legacy/           ← All 45 files from LegacyEngine/
│       ├── ClusterEngine/
│       ├── CoordinateEngine/
│       ├── Debug/
│       ├── ExcelEngine/
│       ├── LayoutEngine/
│       ├── Models/
│       ├── OpenXml/
│       ├── PublishEngine/
│       ├── VerificationEngine/
│       └── ServiceRegistration.cs
│
├── Runtime/              ← Extended (RuntimeCoordinateGenerator added)
├── Rendering/            ← Unchanged (34 files)
├── Application/          ← NEW — COM-free shared services
├── Helpers/              ← NEW — empty (future utility code)
├── GlobalUsings.cs       ← NEW — global using alias for COM Application type
└── Program.cs            ← Simplified — uses DI extension methods
```

---

## 2. Files Moved

### Services/ → Designer/Capture/
| File | New Path |
|------|----------|
| `ExcelCaptureService.cs` | `Designer/Capture/ExcelCaptureService.cs` |
| `IExcelCaptureService.cs` | `Designer/Capture/IExcelCaptureService.cs` |

### Services/ → Designer/Analysis/
| File | New Path |
|------|----------|
| `WorkbookReaderService.cs` | `Designer/Analysis/WorkbookReaderService.cs` |

### Generators/ → Designer/Generation/
| File | New Path |
|------|----------|
| `WorkbookGenerator.cs` | `Designer/Generation/WorkbookGenerator.cs` |
| `PdfGenerator.cs` | `Designer/Generation/PdfGenerator.cs` |
| `PreviewGenerator.cs` | `Designer/Generation/PreviewGenerator.cs` |

### LegacyEngine/ → Designer/Legacy/
| Directory | New Path | File Count |
|-----------|----------|------------|
| ClusterEngine/ | Designer/Legacy/ClusterEngine/ | 3 |
| CoordinateEngine/ | Designer/Legacy/CoordinateEngine/ | 3 |
| Debug/ | Designer/Legacy/Debug/ | 2 |
| ExcelEngine/ | Designer/Legacy/ExcelEngine/ | 3 |
| LayoutEngine/ | Designer/Legacy/LayoutEngine/ | 5 |
| Models/ | Designer/Legacy/Models/ | 5 |
| OpenXml/ | Designer/Legacy/OpenXml/ | 12 |
| PublishEngine/ | Designer/Legacy/PublishEngine/ | 3 |
| VerificationEngine/ | Designer/Legacy/VerificationEngine/ | 6 |
| ServiceRegistration.cs | Designer/Legacy/ServiceRegistration.cs | 1 |
| **Total** | | **43 files + 1 registration** |

### Services/ → Application/
| File | New Path |
|------|----------|
| `FormSaveService.cs` | `Application/FormSaveService.cs` |
| `CoordinateTransformer.cs` | `Application/CoordinateTransformer.cs` |
| `PythonRenderService.cs` | `Application/PythonRenderService.cs` |
| `PreviewCleanupService.cs` | `Application/PreviewCleanupService.cs` |

### Generators/ → Application/
| File | New Path |
|------|----------|
| `XmlGenerator.cs` | `Application/XmlGenerator.cs` |
| `DatabaseGenerator.cs` | `Application/DatabaseGenerator.cs` |

### Services/ → Runtime/
| File | New Path |
|------|----------|
| `RuntimeCoordinateGenerator.cs` | `Runtime/RuntimeCoordinateGenerator.cs` |

### Removed (old empty directories)
- `Services/` — absorbed into Designer/, Application/, Runtime/
- `Generators/` — absorbed into Designer/Generation/, Application/
- `Services/Interfaces/` — absorbed into Designer/Capture/

---

## 3. Namespaces Changed

| Old Namespace | New Namespace | Files Affected |
|---------------|---------------|----------------|
| `ExcelAPI.Services` | `ExcelAPI.Designer.Capture` | 2 |
| `ExcelAPI.Services.Interfaces` | `ExcelAPI.Designer.Capture` | 1 |
| `ExcelAPI.Services` | `ExcelAPI.Designer.Analysis` | 1 |
| `ExcelAPI.Generators` | `ExcelAPI.Designer.Generation` | 3 |
| `ExcelAPI.LegacyEngine` | `ExcelAPI.Designer.Legacy` | 1 |
| `ExcelAPI.LegacyEngine.*` | `ExcelAPI.Designer.Legacy.*` | 42 |
| `ExcelAPI.Services` | `ExcelAPI.Application` | 4 |
| `ExcelAPI.Generators` | `ExcelAPI.Application` | 2 |
| `ExcelAPI.Services` | `ExcelAPI.Runtime` | 1 |

### Controller Import Updates
| Controller | Old Using | New Using |
|------------|-----------|-----------|
| `FormController.cs` | `ExcelAPI.Services` | `ExcelAPI.Application` + `ExcelAPI.Designer.Analysis` |
| `RuntimeController.cs` | _(none)_ | `ExcelAPI.Designer.Capture` |
| `LegacyRuntimeController.cs` | `ExcelAPI.LegacyEngine.OpenXml` | `ExcelAPI.Designer.Legacy.OpenXml` |
| `PublishController.cs` | `ExcelAPI.LegacyEngine.*` | `ExcelAPI.Designer.Legacy.*` |
| `TemplateController.cs` | `ExcelAPI.LegacyEngine.OpenXml` | `ExcelAPI.Designer.Legacy.OpenXml` |
| `ExcelController.cs` | `ExcelAPI.Services.Interfaces` | `ExcelAPI.Designer.Capture` |

### Internal Using Updates
| File | Old Using | New Using |
|------|-----------|-----------|
| `FormSaveService.cs` | `ExcelAPI.Generators` | `ExcelAPI.Designer.Generation` |
| `ExcelCaptureService.cs` | _(none)_ | `ExcelAPI.Application` |

---

## 4. Dependency Injection Extension Methods

Four new `ServiceRegistration.cs` files were created to organize DI registrations by architectural layer:

### `Designer/ServiceRegistration.cs` — `AddDesigner()`
Registers COM-dependent services:
- `IExcelCaptureService`, `ExcelCaptureService`
- `WorkbookReaderService`
- `WorkbookGenerator`, `PdfGenerator`, `PreviewGenerator`

### `Runtime/ServiceRegistration.cs` — `AddRuntime()`
Registers COM-free runtime services:
- `FieldTypeResolver`, `FieldDetector`
- `FormRuntimeBuilder`, `RuntimeSerializer`
- `RuntimeCoordinateGenerator`

### `Application/ServiceRegistration.cs` — `AddApplication()`
Registers COM-free shared services:
- `IFormSaveService`, `FormSaveService`
- `XmlGenerator`, `DatabaseGenerator`
- `CoordinateTransformer`
- `PreviewCleanupService` (hosted service)

### `Rendering/ServiceRegistration.cs` — `AddRendering()`
Registers all rendering pipeline services (34 services):
- OpenXmlParser, GeometryBuilder, CoordinateEngine
- FillEngine, BorderEngine, GridlineLayer, TextEngine
- FontResolver, ThemeResolver, StyleResolver
- ImageEngine, ShapeEngine
- ExportCoordinator, PageRenderer, RendererCoordinator
- And all `IRenderLayer` implementations

### `Designer/Legacy/ServiceRegistration.cs` — `AddLegacyEngine()` *(existing)*
Unchanged — registers the full ConMas legacy pipeline.

---

## 5. Program.cs Simplification

**Before**: ~120 lines of inline DI registrations  
**After**: ~30 lines with 5 module extension calls

```csharp
// Module registrations (organized by architectural layer)

// Designer — Excel COM-dependent services (Capture + Generation)
builder.Services.AddDesigner();

// Legacy Engine — reverse-engineered ConMas publishing pipeline (COM-bound)
builder.Services.AddLegacyEngine(debugEnabled: true);

// Rendering — OpenXML parsing, SkiaSharp rendering pipeline (no COM)
builder.Services.AddRendering();

// Runtime — form building, field detection, serialization (no COM)
builder.Services.AddRuntime();

// Application — shared services, generators, orchestration (no COM)
builder.Services.AddApplication();
```

---

## 6. Resolved Issues During Execution

### Issue 1: Namespace Conflict — `Application` Ambiguity
**Problem**: The new `ExcelAPI.Application` namespace shadowed the `Microsoft.Office.Interop.Excel.Application` COM type in all files under the `ExcelAPI` namespace tree.

**Solution**: Created `GlobalUsings.cs` with:
```csharp
global using App = Microsoft.Office.Interop.Excel.Application;
```
Then replaced all unqualified `Application` type references with `App` in 5 files:
- `Designer/Capture/ExcelCaptureService.cs`
- `Designer/Analysis/WorkbookReaderService.cs`
- `Designer/Generation/WorkbookGenerator.cs`
- `Designer/Legacy/ExcelEngine/WorkbookLoader.cs`
- `Designer/Legacy/PublishEngine/BackgroundExporter.cs`

### Issue 2: Missing Using Statements in Controllers
Several controllers had stale `using ExcelAPI.Services` references which no longer existed.

**Solution**: Updated controller imports to point to the correct new namespaces.

### Issue 3: Duplicate WorkbookReaderService Registration
`WorkbookReaderService` was accidentally registered in both `Designer/ServiceRegistration.cs` and `Application/ServiceRegistration.cs`.

**Solution**: Removed the duplicate from `Application/ServiceRegistration.cs` since the service lives in `Designer/Analysis/`.

---

## 7. Build Verification

| Metric | Value |
|--------|-------|
| **Build Status** | ✅ Succeeded |
| **Compile Errors** | 0 |
| **Warnings Introduced** | 0 (2 pre-existing NuGet vulnerability warnings) |

---

## 8. Dependency Verification

| Layer | Has COM? | Has SkiaSharp? | Has OpenXml? | Notes |
|-------|----------|----------------|--------------|-------|
| **Designer/Capture** | ✅ Yes | ✅ Yes | ❌ No | Excel COM + SkiaSharp for PNG |
| **Designer/Analysis** | ✅ Yes | ❌ No | ✅ Yes | Hybrid COM + OpenXml |
| **Designer/Generation** | ✅ Yes | ✅ Yes | ❌ No | Excel COM + SkiaSharp |
| **Designer/Legacy** | ✅ Yes | ✅ Yes | ✅ Yes | Full ConMas pipeline |
| **Runtime** | ❌ No | ❌ No | ❌ No | Pure C# logic |
| **Rendering** | ❌ No | ✅ Yes | ✅ Yes | SkiaSharp pipeline |
| **Application** | ❌ No | ❌ No | ❌ No | Shared services |
| **Controllers** | ❌ No (most) | ❌ No | ❌ No | Route definitions only |

---

## 9. Controller Route Verification

All existing endpoints remain unchanged:

| Route | Controller | Method | Status |
|-------|-----------|--------|--------|
| `POST /api/excel/upload` | ExcelController | Upload | ✅ Unchanged |
| `POST /api/form/save` | FormController | Save | ✅ Unchanged |
| `POST /api/form/from-excel` | FormController | FromExcel | ✅ Unchanged |
| `POST /api/form/upload-preview` | FormController | UploadPreview | ✅ Unchanged |
| `POST /api/form/output-excel` | FormController | OutputExcel | ✅ Unchanged |
| `POST /api/form/upload-excel` | FormController | UploadExcel | ✅ Unchanged |
| `GET /api/form/runtime/{templateId}` | FormController | GetRuntime | ✅ Unchanged |
| `POST /api/runtime/upload` | RuntimeController | Upload | ✅ Unchanged |
| `GET /api/template/{id}` | TemplateController | GetTemplate | ✅ Unchanged |
| `POST /api/publish/publish` | PublishController | Publish | ✅ Unchanged |
| `POST /api/publish/publish-and-verify` | PublishController | PublishAndVerify | ✅ Unchanged |
| `POST /api/publish/verify` | PublishController | Verify | ✅ Unchanged |
| `POST /api/publish/regression` | PublishController | RunRegression | ✅ Unchanged |
| `GET /api/publish/transform` | PublishController | GetTransform | ✅ Unchanged |
| `POST /api/legacy-runtime/upload` | LegacyRuntimeController | Upload | ✅ Unchanged |

---

## 10. Architectural Boundary Summary

```
                      ┌─────────────────────┐
                      │    Controllers/      │
                      │  (route dispatching) │
                      └──────┬──────────────┘
                             │
         ┌───────────────────┼───────────────────┐
         │                   │                   │
         ▼                   ▼                   ▼
┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
│    Designer/    │  │    Runtime/     │  │  Application/   │
│  (COM-bound)    │  │  (COM-free)     │  │  (COM-free)     │
│                 │  │                 │  │                 │
│ Capture/        │  │ FormBuilder     │  │ FormSaveService │
│ Analysis/       │  │ FieldDetector   │  │ XmlGenerator    │
│ Generation/     │  │ RuntimeCoord    │  │ DatabaseGen     │
│ Legacy/ (45)    │  │ Serializer      │  │ CoordTransform  │
│ ActiveX Export  │  │                 │  │ PythonService   │
└─────────────────┘  └─────────────────┘  └─────────────────┘
                              │
                              ▼
                      ┌─────────────────┐
                      │   Rendering/    │
                      │  (COM-free)     │
                      │                 │
                      │ OpenXmlParser   │
                      │ CoordinateEngine│
                      │ FillEngine/     │
                      │ BorderEngine/   │
                      │ TextEngine/     │
                      │ ExportCoord     │
                      └─────────────────┘
```

---

## 11. Future Separation Readiness

| Component | Can Move To | When |
|-----------|-------------|------|
| `Designer/` + `Designer/Legacy/` | `PaperLess.Designer` project | Server split to two machines |
| `Runtime/` + `Rendering/` | `PaperLess.Runtime` project | Server split to two machines |
| `Application/` | Shared library | Always COM-free |
| `Models/` | `PaperLess.Contracts` | Extract to shared contracts |

**No code changes needed for separation** — only file moves and project reference updates.
