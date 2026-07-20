# Phase 6 — Reopen Existing PaperLess / Legacy ConMas Workbook

**Status: Complete** ✅

---

## Objective

Implement the ability to reopen an exported PaperLess (or legacy ConMas) workbook and completely reconstruct the designer state. The workbook itself is the source of truth — no database, no project file, no external metadata required.

---

## Design

### Detection

When a workbook is uploaded via `POST /api/form/upload-preview`, the backend now checks for embedded designer configuration **before** running the Python preview pipeline.

Detection is lightweight OOXML-based (no COM required):

```csharp
private static bool HasEmbeddedConfiguration(string xlsxPath)
{
    using var archive = ZipFile.OpenRead(xlsxPath);
    var workbookEntry = archive.GetEntry("xl/workbook.xml");
    // Check for ExcelOutputSetting or _Fields sheet names
    return workbookXml.Contains("ExcelOutputSetting")
        || workbookXml.Contains("_Fields");
}
```

### Reconstruction Pipeline

When embedded configuration is detected, the upload flow redirects to:

```
Uploaded Workbook
        │
        ▼
HasEmbeddedConfiguration() → true
        │
        ├──► WorkbookReaderService.Read()
        │        ↓
        │    FormDefinition
        │        ↓
        ├──► WorkbookDefinitionConverter.FromFormDefinition()
        │        ↓
        │    WorkbookDefinition
        │        ↓
        ├──► FormRuntimeBuilder.BuildFromDefinitionDirect()
        │        ↓
        │    RuntimeForm
        │
        ├──► PythonRenderService.UploadPreviewAsync()  (for background image only)
        │
        ├──► SessionWorkbookStore.CreateSession()
        │
        ▼
Return { success, sessionId, runtimeForm, isReconstructed }
```

### Response Format

**Reopened workbook** (config detected):
```json
{
    "success": true,
    "sessionId": "guid...",
    "isReconstructed": true,
    "message": "Existing designer configuration detected and loaded. 12 field(s) restored.",
    "runtimeForm": { ... full RuntimeForm ... },
    "fieldCount": 12,
    "sheetCount": 1
}
```

**New template** (no config — unchanged):
```json
{
    "success": true,
    "sessionId": "guid...",
    "pages": [ ... existing page/field data ... ]
}
```

---

## Files Modified

### 1. `ExcelAPI/ExcelAPI/Controllers/FormController.cs`

**New method: `HasEmbeddedConfiguration(string xlsxPath)`**
- Opens the workbook as a ZIP archive
- Reads `xl/workbook.xml` 
- Checks for `ExcelOutputSetting` or `_Fields` sheet names via `string.Contains`
- Returns `true` if either sheet exists
- No COM Interop, no heavy XML parsing — pure OOXML string check

**Modified: `UploadPreview([FromForm] IFormFile? file)`**
- **Config detected path**: 
  1. Calls `WorkbookReaderService.Read()` to reconstruct FormDefinition via COM
  2. Converts `FormDefinition → WorkbookDefinition` via canonical converter
  3. Builds `RuntimeForm` directly using `FormRuntimeBuilder.BuildFromDefinitionDirect()`
  4. Calls Python renderer for background image only (field data comes from COM reconstruction)
  5. Sets background image URLs and page dimensions from Python result
  6. Creates session in session store
  7. Returns `runtimeForm` directly to frontend
- **No config path**: Unchanged existing flow

### 2. `paperless/app/page.tsx`

**Modified: `handleUpload()`** — detects `result.runtimeForm` property:
- If present: Uses the RuntimeForm directly with safe defaults for all fields
- If absent: Falls through to existing Python-page conversion flow
- Added `console.log` diagnostics for debugging

---

## How It Works

### Scenario 1 — Upload Normal Excel (New Template)
1. User uploads `MyForm.xlsx` (no embedded config)
2. `HasEmbeddedConfiguration()` returns `false`
3. Existing Python preview flow executes
4. Field ratios and background image returned as before
5. Frontend converts pages to RuntimeForm

### Scenario 2 — Upload Exported Workbook (Reopen)
1. User uploads the previously exported workbook
2. `HasEmbeddedConfiguration()` detects `ExcelOutputSetting` or `_Fields` sheets → returns `true`
3. `WorkbookReaderService.Read()` reads the workbook via COM, reconstructing:
   - All sheets (excluding _Fields)
   - All field definitions from _Fields sheet
   - Cell styles, fonts, borders, alignment
   - Page setup, print area, margins
   - Merged cells, row heights, column widths
4. FormDefinition converted to WorkbookDefinition → RuntimeForm
5. Python renderer generates background image (used for visual overlay)
6. RuntimeForm returned directly to frontend
7. Designer loads with all fields, properties, and layout restored

---

## Verification

### Build
- Backend: `dotnet build` — **Build succeeded** (0 errors)
- Frontend: `npx tsc --noEmit` — **0 errors**

### Code Review
- Detection is lightweight OOXML (no COM until confirmed)
- FormRuntimeBuilder uses parameterless constructor (no DI overhead)
- Python renderer still generates background images for reopened workbooks
- Existing Python-only flow completely unchanged
- Frontend has safe defaults for all RuntimeForm fields
- Session store integration preserved

### API Compatibility
- Existing API responses unchanged for new templates
- Reopened workbooks return an additional `runtimeForm` property
- Frontend detects `runtimeForm` vs `pages` to choose the display path

---

## Migration Path

| Phase | Backend | Frontend |
|-------|---------|----------|
| 6 (current) | OOXML detection + WorkbookReaderService reconstruction | Handle `runtimeForm` response |
| Future | Remove round-trip upload endpoints entirely | Remove old Python-page conversion when all users are on new pipeline |

---

## Success Criteria

| Criteria | Status |
|----------|--------|
| ✅ Automatic detection of embedded configuration | Done |
| ✅ Reconstruct designer state from embedded metadata | Done |
| ✅ Return RuntimeForm directly to frontend | Done |
| ✅ Background image still rendered for reopened workbooks | Done |
| ✅ No changes to existing new-template flow | Done |
| ✅ Frontend handles both response formats | Done |
| ✅ Build passes (0 errors) | Done |
| ✅ TypeScript compilation passes (0 errors) | Done |
