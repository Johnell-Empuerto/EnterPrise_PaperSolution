# Phase 3.1 — Refactor Current Implementation into the PaperLess Architecture

**Date**: 2026-07-19
**Objective**: Refactor the existing monolithic `ExcelAPI` project into clear architectural layers that naturally separate Designer (Excel/COM) concerns from Runtime (no-COM) concerns.

---

## 1. Current State Assessment

### What Exists Today

```
ExcelAPI (single project, ~128 files)
├── Controllers/       (7 — mixed Designer + Runtime endpoints)
├── Services/          (8 — COM + non-COM mixed)
├── Generators/        (5 — workbook, PDF, preview, XML, DB)
├── Models/            (17 — DTOs, form definitions, errors)
├── Rendering/         (34 — SkiaSharp rendering pipeline, no COM)
├── Runtime/           (7 — field detection, runtime form builder)
├── LegacyEngine/      (45 — COM-based ConMas pipeline)
└── Program.cs         (all DI registration in one place)
```

### Problem: No Separation of Concerns

| Component | Has COM? | Designer Concern? | Runtime Concern? | Current Location |
|-----------|----------|-------------------|------------------|-----------------|
| `ExcelCaptureService` | Yes | Capture print area | — | `Services/` |
| `WorkbookGenerator` | Yes | Generate output XLSX | — | `Generators/` |
| `PdfGenerator` | Yes | Generate PDF via COM | — | `Generators/` |
| `PreviewGenerator` | Yes | Generate preview PNG | — | `Generators/` |
| `PublishEngine` (LegacyEngine) | Yes | Full ConMas publish | — | `LegacyEngine/PublishEngine/` |
| `BackgroundExporter` | Yes | PDF→PNG export | — | `LegacyEngine/PublishEngine/` |
| `FormSaveService` | No | Save form pipeline | — | `Services/` |
| `WorkbookReaderService` | No | Read workbook | Yes (reads generated output) | `Services/` |
| `OpenXmlParser` | No | — | Yes (parse workbooks) | `Rendering/` |
| `CoordinateEngine` | No | — | Yes (geometry) | `Rendering/` |
| `ExportCoordinator` | No | — | Yes (render to PNG/PDF) | `Rendering/` |
| `FormRuntimeBuilder` | No | — | Yes (build runtime form) | `Runtime/` |
| `RuntimeSerializer` | No | — | Yes (JSON serialization) | `Runtime/` |
| `DatabaseGenerator` | No | Save to DB | Yes | `Generators/` |
| `FieldDetector` | No | — | Yes | `Runtime/` |
| All render layers (FillEngine, BorderEngine, TextEngine, etc.) | No | — | Yes | `Rendering/` |
| `RegressionTestRunner` | No | — | Yes (validation) | `Rendering/Validation/` |

---

## 2. Target Architecture

```
PaperLess.sln
│
├── PaperLess.Contracts/           ← Shared DTOs & Interfaces (zero dependencies)
│   ├── Models/
│   │   ├── FormDefinition.cs      ← Master form model (WorkbookDefinition, SheetDefinition, ClusterDefinition)
│   │   ├── RuntimeForm.cs         ← Runtime form (RuntimeSheet, RuntimeField)
│   │   ├── CaptureResult.cs       ← Excel capture output
│   │   ├── ApiResponse.cs         ← Generic API response
│   │   ├── ApiErrorResponse.cs    ← Error response
│   │   ├── ExcelField.cs          ← Single detected field
│   │   ├── PageInfo.cs            ← Page dimensions
│   │   ├── RuntimeUploadResponse.cs
│   │   └── TemplateModelDto.cs
│   └── Interfaces/
│       ├── IExcelCaptureService.cs    ← Contract (implemented by Designer)
│       ├── IRendererService.cs        ← Contract (implemented by Runtime)
│       └── IStorageService.cs         ← Contract (implemented by Storage)
│
├── PaperLess.Designer.Engine/     ← Excel COM automation (Windows-only)
│   ├── Services/
│   │   ├── ExcelCaptureService.cs  ← COM: Open, read print area, export PDF, detect fields
│   │   ├── WorkbookGenerator.cs    ← COM: Generate output XLSX
│   │   ├── PdfGenerator.cs         ← COM: Export PDF via ExportAsFixedFormat
│   │   └── PreviewGenerator.cs     ← COM: Generate preview PNG
│   ├── LegacyEngine/               ← ConMas backward compatibility
│   │   ├── PublishEngine/
│   │   ├── ClusterEngine/
│   │   ├── CoordinateEngine/
│   │   ├── LayoutEngine/
│   │   └── VerificationEngine/
│   ├── CoordinateTransformer.cs    ← ConMas centering formula
│   └── ServiceRegistration.cs      ← DI extension method
│
├── PaperLess.Designer.Api/        ← HTTP API for Designer (has Excel)
│   ├── Controllers/
│   │   ├── ExcelController.cs      ← POST /api/excel/upload (capture)
│   │   ├── PublishController.cs    ← POST /api/publish (ConMas pipeline)
│   │   └── LegacyRuntimeController.cs
│   ├── Program.cs
│   └── appsettings.json
│
├── PaperLess.Runtime.Engine/      ← Rendering & form building (no COM)
│   ├── Rendering/                  ← SkiaSharp pipeline (all 34 files)
│   │   ├── OpenXmlParser.cs
│   │   ├── GeometryBuilder.cs
│   │   ├── CoordinateEngine.cs
│   │   ├── ExportCoordinator.cs
│   │   ├── RendererCoordinator.cs
│   │   ├── PageRenderer.cs
│   │   ├── PrintLayoutEngine.cs
│   │   ├── FillEngine.cs, BorderEngine.cs, TextEngine.cs, ...
│   │   ├── ImageEngine.cs, ShapeEngine.cs
│   │   └── Validation/             ← Regression testing
│   ├── Runtime/                    ← Form building
│   │   ├── FormRuntimeBuilder.cs
│   │   ├── FieldDetector.cs
│   │   ├── FieldTypeResolver.cs
│   │   ├── RuntimeSerializer.cs
│   │   └── RuntimeCoordinateGenerator.cs
│   ├── Services/
│   │   ├── FormSaveService.cs      ← Orchestrates save pipeline
│   │   ├── WorkbookReaderService.cs
│   │   └── PythonRenderService.cs   ← External render client
│   ├── Generators/
│   │   ├── XmlGenerator.cs
│   │   └── DatabaseGenerator.cs
│   └── ServiceRegistration.cs
│
├── PaperLess.Runtime.Api/         ← HTTP API for frontend (no Excel)
│   ├── Controllers/
│   │   ├── FormController.cs       ← POST /api/form/save, GET /api/form/runtime/{id}
│   │   ├── RuntimeController.cs    ← POST /api/runtime/upload
│   │   └── TemplateController.cs   ← GET /api/template/{id}
│   ├── Program.cs
│   └── appsettings.json
│
├── PaperLess.Storage/             ← Database access (no COM, no HTTP)
│   ├── DatabaseGenerator.cs       ← SQL generation (Npgsql)
│   └── ServiceRegistration.cs
│
└── PaperLess.Shared/              ← Common utilities (optional)
    ├── PreviewCleanupService.cs    ← IHostedService for temp file cleanup
    └── StringExtensions.cs
```

### Project Dependencies

```
PaperLess.Contracts  ← Zero dependencies (pure DTOs + interfaces)
    ↑
PaperLess.Shared  ← Only depends on Contracts (shared utilities)
    ↑
PaperLess.Designer.Engine  ← Depends on Contracts + Microsoft.Office.Interop.Excel
    ↑
PaperLess.Designer.Api  ← Depends on Designer.Engine + Contracts
    |
PaperLess.Runtime.Engine  ← Depends on Contracts + DocumentFormat.OpenXml + SkiaSharp
    ↑
PaperLess.Runtime.Api  ← Depends on Runtime.Engine + Contracts
    |
PaperLess.Storage  ← Depends on Contracts + Npgsql
```

---

## 3. Migration Path: Today → Future

### Phase 3.1a — Create Projects (No Code Changes)

Create the solution structure with empty projects and move files:

| From | To |
|------|-----|
| `ExcelAPI/Contracts/` (extract) | `PaperLess.Contracts/` |
| `ExcelAPI/Services/ExcelCaptureService.cs` | `PaperLess.Designer.Engine/Services/` |
| `ExcelAPI/Generators/WorkbookGenerator.cs` | `PaperLess.Designer.Engine/Services/` |
| `ExcelAPI/Generators/PdfGenerator.cs` | `PaperLess.Designer.Engine/Services/` |
| `ExcelAPI/Generators/PreviewGenerator.cs` | `PaperLess.Designer.Engine/Services/` |
| `ExcelAPI/LegacyEngine/` | `PaperLess.Designer.Engine/LegacyEngine/` |
| `ExcelAPI/Services/CoordinateTransformer.cs` | `PaperLess.Designer.Engine/` |
| `ExcelAPI/Rendering/` (all 34 files) | `PaperLess.Runtime.Engine/Rendering/` |
| `ExcelAPI/Runtime/` (all 7 files) | `PaperLess.Runtime.Engine/Runtime/` |
| `ExcelAPI/Services/FormSaveService.cs` | `PaperLess.Runtime.Engine/Services/` |
| `ExcelAPI/Services/WorkbookReaderService.cs` | `PaperLess.Runtime.Engine/Services/` |
| `ExcelAPI/Services/PythonRenderService.cs` | `PaperLess.Runtime.Engine/Services/` |
| `ExcelAPI/Generators/XmlGenerator.cs` | `PaperLess.Runtime.Engine/Generators/` |
| `ExcelAPI/Generators/DatabaseGenerator.cs` | `PaperLess.Storage/` |
| `ExcelAPI/Services/RuntimeCoordinateGenerator.cs` | `PaperLess.Runtime.Engine/Runtime/` |
| `ExcelAPI/Services/PreviewCleanupService.cs` | `PaperLess.Shared/` |

### Phase 3.1b — Fix Namespaces

Update all `using ExcelAPI.*` references to the new project namespaces:
- `using PaperLess.Contracts`
- `using PaperLess.Designer.Engine`
- `using PaperLess.Runtime.Engine`
- `using PaperLess.Storage`

### Phase 3.1c — Separate Controllers

The current controllers contain a mix of concerns. Split them:

**Designer API Controllers** (keep in `PaperLess.Designer.Api`):
- `ExcelController` — upload + capture (COM)
- `PublishController` — ConMas publish pipeline (COM)
- `LegacyRuntimeController` — legacy runtime (COM)

**Runtime API Controllers** (move to `PaperLess.Runtime.Api`):
- `FormController` — save, from-excel, output-excel, runtime/{id}, upload-preview
- `RuntimeController` — upload (production endpoint)
- `TemplateController` — get template metadata

### Phase 3.1d — Remove Cross-Project Dependencies

**Current problem**: `FormController` uses `IExcelCaptureService` directly. After the split, the Runtime API cannot reference the Designer Engine.

**Solution**: The `POST /api/form/from-excel` endpoint in `FormController` calls `IExcelCaptureService`. In the future architecture, this endpoint belongs to the **Designer API**, not the Runtime API. The Runtime API should receive a pre-built JSON payload.

For the transition period:
1. `PaperLess.Designer.Api` handles Excel upload + capture
2. `PaperLess.Runtime.Api` handles save, render, runtime retrieval
3. The frontend calls Designer API first (upload Excel → get JSON), then calls Runtime API (POST JSON → save + render)

| Current Endpoint | Future Home | Reason |
|------------------|-------------|--------|
| `POST /api/excel/upload` | **Designer API** | Requires COM |
| `POST /api/publish/publish` | **Designer API** | Requires COM |
| `POST /api/publish/publish-and-verify` | **Designer API** | Requires COM |
| `POST /api/publish/verify` | **Designer API** | Requires DB + COM comparison |
| `POST /api/publish/regression` | **Designer API** | Requires COM for publish step |
| `GET /api/publish/transform` | **Designer API** | Legacy engine config |
| `POST /api/form/save` | **Runtime API** | No COM |
| `POST /api/form/from-excel` | → **Designer API** (capture) → **Runtime API** (save) | Two-step in production |
| `POST /api/form/output-excel` | **Designer API** | Requires COM for generation |
| `GET /api/form/runtime/{id}` | **Runtime API** | No COM |
| `POST /api/form/upload-preview` | **Designer API** | Requires COM/Excel |
| `POST /api/form/upload-excel` | **Runtime API** | OpenXML only, no COM |
| `POST /api/runtime/upload` | → **Designer API** (capture) → **Runtime API** (persist) | Two-step in production |
| `GET /api/template/{id}` | **Runtime API** | Read metadata, no COM |
| `POST /api/legacy-runtime/upload` | **Designer API** | Requires COM |

---

## 4. JSON Contract (Stable, Versioned)

The `RuntimeForm` model must be the stable contract between the Designer and the Runtime server. It must be pure JSON — no COM, no Excel, no binary dependencies.

### RuntimeForm (already exists in `Runtime/`)

```json
{
  "workbookName": "CustomerForm.xlsx",
  "title": "Customer Information Form",
  "version": "1.0",
  "sheets": [
    {
      "sheetName": "Page 1",
      "pageWidthPx": 2550,
      "pageHeightPx": 3300,
      "dpi": 300,
      "backgroundImage": "/forms/bg_abc123.png",
      "printArea": { "leftPt": 70, "topPt": 70, "widthPt": 476, "heightPt": 652 },
      "fields": [
        {
          "id": "f1",
          "name": "customer_name",
          "cellReference": "B2",
          "leftRatio": 0.12,
          "topRatio": 0.08,
          "widthRatio": 0.35,
          "heightRatio": 0.04,
          "type": "text",
          "required": true,
          "readOnly": false,
          "maxLength": 100,
          "tabIndex": 1,
          "font": { "name": "Calibri", "size": 11, "bold": false },
          "border": { "style": "thin", "color": "#000000" },
          "alignment": { "horizontal": "left", "vertical": "center" },
          "validationPattern": null,
          "validationMessage": null,
          "placeholder": "Enter customer name"
        }
      ],
      "images": [],
      "shapes": []
    }
  ]
}
```

**Key design decisions**:
- **All coordinates are ratios** (0.0–1.0) — resolution-independent
- **No Excel references** — cellReference is informational only
- **backgroundImage is a URL** — the Runtime server serves static files
- **Font, border, alignment are fully specified** — no need to reparse OOXML
- **Validation rules are explicit** — no inference needed

This JSON is the **single artifact** that moves from the Designer API to the Runtime API. The Runtime API never needs to see the original .xlsx file.

---

## 5. Service Refactoring

### Services That Stay in Designer.Engine (COM required)

| Service | Reason | Refactoring |
|---------|--------|-------------|
| `ExcelCaptureService` | COM automation | Split interface into `Contracts`, keep implementation |
| `WorkbookGenerator` | COM `Workbook.SaveAs` | Same |
| `PdfGenerator` | COM `ExportAsFixedFormat` | Same |
| `PreviewGenerator` | COM + PDFtoImage | Same |
| `PublishEngine` | COM workbook loading | Same |
| All LegacyEngine COM loaders | COM Interop | Same |

### Services That Move to Runtime.Engine (no COM)

| Service | Reason | Refactoring |
|---------|--------|-------------|
| `FormSaveService` | Orchestrates generators | Extract `IWorkbookReader` interface, inject via constructor |
| `WorkbookReaderService` | OpenXML-only reader | Pure OOXML parsing, no changes needed |
| `PythonRenderService` | HTTP client to external service | No changes needed |
| `CoordinateTransformer` | Math formula | Pure calculation |
| `RuntimeCoordinateGenerator` | JSON persistence | File I/O only |

### Services That Become Implementation-Specific

| Current | Split Into |
|---------|-----------|
| `DatabaseGenerator` | Interface in `Contracts`, impl in `Storage` |
| `XmlGenerator` | Interface in `Contracts`, impl in `Runtime.Engine` |

---

## 6. Interface Contracts

```csharp
// PaperLess.Contracts/Interfaces/IExcelCaptureService.cs
public interface IExcelCaptureService
{
    Task<CaptureResult> CapturePrintAreaAsync(
        string filePath, string? templateId,
        CancellationToken ct = default);
}

// PaperLess.Contracts/Interfaces/IFormRuntimeService.cs
public interface IFormRuntimeService
{
    RuntimeForm BuildRuntime(string templateId, string? xlsxPath = null);
}

// PaperLess.Contracts/Interfaces/IStorageService.cs
public interface IStorageService
{
    Task SaveFormDefinitionAsync(FormDefinition form);
    Task<FormDefinition?> LoadFormDefinitionAsync(string templateId);
}
```

---

## 7. Runtime Rendering Architecture (No COM, No Excel)

The frontend renders forms purely from JSON:

```
Frontend (Browser)
    │
    ├─ 1. GET /api/form/runtime/{id}
    │      ↓
    │      RuntimeForm JSON ← Lightweight API call
    │
    ├─ 2. Load background image from URL
    │      ↓
    │      PNG image (served as static file)
    │
    └─ 3. Overlay fields on background using ratio coordinates
           ↓
           Yellow editable overlays (CSS)
```

The Runtime server never:
- Opens Excel
- Uses COM
- Generates PDFs
- Reads OOXML at request time

The RuntimeForm JSON is pre-computed once (when the template is uploaded) and served from cache or storage on every subsequent request.

---

## 8. Extensibility Points

| Future Feature | Where It Goes | Why |
|---------------|---------------|-----|
| **OCR** | `PaperLess.Designer.Engine` | OCR runs on upload, extracts text from scanned forms |
| **PDF import** | `PaperLess.Designer.Engine` | PDF→Excel conversion requires processing |
| **Multi-page** | Already supported in `RuntimeForm.Sheets[]` | Just add more sheets |
| **Versioning** | `PaperLess.Storage` | Version tracking in database schema |
| **Collaboration** | `PaperLess.Runtime.Api` | Real-time sync is a Runtime concern |
| **Validation rules** | `PaperLess.Runtime.Engine` | Rule engine is pure logic, no COM |
| **Template publishing** | `PaperLess.Designer.Api` | Publishing workflow controls access |
| **Template library** | `PaperLess.Runtime.Api` | CRUD operations, template search |

---

## 9. Incremental Migration Plan

### Step 1: Create Solution + Projects (Day 1)

```
dotnet new sln -n PaperLess
dotnet new classlib -n PaperLess.Contracts
dotnet new classlib -n PaperLess.Designer.Engine
dotnet new classlib -n PaperLess.Runtime.Engine
dotnet new webapi -n PaperLess.Designer.Api
dotnet new webapi -n PaperLess.Runtime.Api
dotnet new classlib -n PaperLess.Storage
dotnet new classlib -n PaperLess.Shared

dotnet sln add **/*.csproj
```

### Step 2: Extract Contracts (Day 1)

Move DTOs and interfaces to `PaperLess.Contracts`. Update all `using` statements.

### Step 3: Move Files to Target Projects (Day 2)

Move source files without changing logic. Update namespaces. Fix `Program.cs` DI registrations — split into each API project.

### Step 4: Split Controllers (Day 2–3)

- `PaperLess.Designer.Api` gets: `ExcelController`, `PublishController`, `LegacyRuntimeController`
- `PaperLess.Runtime.Api` gets: `FormController` (minus `from-excel` and `output-excel`), `RuntimeController`, `TemplateController`

### Step 5: Remove Cross-Project COM References (Day 3)

The `FormController` currently calls `IExcelCaptureService` directly. After the split:
- The frontend calls Designer API first (upload + capture) → receives RuntimeForm JSON
- Then calls Runtime API (save + render) → receives saved form ID + URLs

This two-step flow eliminates the COM dependency from the Runtime API.

### Step 6: Update Frontend (Day 3–4)

Update the Next.js frontend to call two APIs:
1. `POST /api/excel/upload` (Designer API) for Excel processing
2. `POST /api/form/save` (Runtime API) for persisting

### Step 7: Remove Dead Code (Day 4)

Once the migration is complete and verified:
- Remove `IExcelCaptureService` dependency from Runtime API
- Remove `DocumentFormat.OpenXml` from Designer API (if not used directly)
- Clean up duplicate `appsettings.json` entries

---

## 10. Summary: What Changes and What Stays

| Aspect | Current | After Refactoring |
|--------|---------|-------------------|
| **Projects** | 1 (ExcelAPI) | 7 |
| **Controllers** | Mixed in one project | Split by concern (Designer vs Runtime) |
| **COM dependencies** | Everywhere | Isolated to `Designer.Engine` |
| **NuGet dependencies** | Single project has everything | Each project has only what it needs |
| **OpenXml/SkiaSharp** | Same project as COM | `Runtime.Engine` only |
| **Npgsql** | Same project | `Storage` only |
| **Code changes** | — | Namespace updates + DI split (no logic changes) |
| **API endpoints** | Same URLs | Same URLs (can be deployed as one or two servers) |
| **JSON contract** | Same `RuntimeForm` | Same — unchanged |
| **Rendering pipeline** | Same 34 files | Same — moved to `Runtime.Engine` |
| **Legacy engine** | Same 45 files | Same — moved to `Designer.Engine` |
| **Frontend code** | Calls one API | Calls two APIs (or one, depending on deployment) |

### Deployment Options After Refactoring

**Single-server** (development):
```
PaperLess.Designer.Api + PaperLess.Runtime.Api → single deployment
```
Just deploy both API projects on the same machine with Excel installed. Controllers are merged at the routing level.

**Two-server** (production):
```
Server A (Designer): PaperLess.Designer.Api + PaperLess.Designer.Engine
    Has Excel, COM, legacy engine
    Handles: file upload, capture, publish, output-excel

Server B (Runtime): PaperLess.Runtime.Api + PaperLess.Runtime.Engine + PaperLess.Storage
    No Excel, no COM
    Handles: form save, runtime queries, rendering, storage
```

The two-server deployment requires zero code changes — only configuration changes (different URLs for the frontend).
