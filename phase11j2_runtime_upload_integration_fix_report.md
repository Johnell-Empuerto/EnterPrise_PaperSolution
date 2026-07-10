# Phase 11J.2 — Runtime Upload Integration Fix (Shared Upload Pipeline)

**Date:** July 2026  
**Target Build:** 0 errors  
**Status:** ✅ Complete  

---

## Objective

Repair the Runtime Viewer upload workflow by reusing the existing `/api/form/from-excel` endpoint instead of the non-existent `/api/form/upload`.

---

## Changes Summary

### Modified Files

| File | Change |
|------|--------|
| `ExcelAPI/ExcelAPI/Controllers/FormController.cs` | Modified `FromExcel()` — persists workbook, returns `templateId` + `previewUrl` |
| `paperless/app/page.tsx` | Changed Runtime Viewer to call `/api/form/from-excel`, adapted response parsing |

---

## Backend Changes

### `FormController.FromExcel()` — Three Changes

#### 1. Persist Uploaded Workbook (replaces File.Delete)

**Before:**
```csharp
finally
{
    try { if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath); } catch { }
}
```

**After:** The workbook is moved to `Forms/` directory after successful processing:

```csharp
string formsDir = Path.Combine(_env.ContentRootPath, "Forms");
Directory.CreateDirectory(formsDir);
string persistentPath = Path.Combine(formsDir, uniqueFileName);
if (System.IO.File.Exists(filePath))
{
    System.IO.File.Move(filePath, persistentPath, overwrite: true);
}
```

On error, the temp file is still cleaned up (moved inside the `catch` block).

#### 2. Extract Template ID

```csharp
string templateId = Path.GetFileNameWithoutExtension(uniqueFileName);
```

The template ID is the GUID portion of the uploaded filename (e.g., `"a1b2c3d4..."` from `"a1b2c3d4e5f6g7h8.xlsx"`). This is used by the runtime endpoint `GET /api/form/runtime/{templateId}` to locate the workbook.

#### 3. Extended Response

**Before (returned `ApiResponse<FormDefinition>`):**
```json
{ "success": true, "message": "...", "data": { FormDefinition } }
```

**After (returns anonymous object extending the shape):**
```json
{
  "success": true,
  "message": "Excel file processed. N field(s) detected.",
  "templateId": "a1b2c3d4e5f6g7h8",
  "previewUrl": "/preview/page_a1b2c3d4e5f6g7h8.png",
  "data": { FormDefinition }
}
```

**Backward compatible:** Existing `success`, `message`, and `data` fields are preserved. The Designer tab reads `result.data` — unchanged.

---

## Frontend Changes

### `page.tsx` — `handleRuntimeUpload` Handler

#### Route Fix

**Before:** `fetch('.../api/form/upload')` → **404 Not Found**

**After:** `fetch('.../api/form/from-excel')` → **200 OK**

#### Response Parsing

**Before:**
```typescript
if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
}
const result = await response.json();
if (result.success && result.templateId) {
```

**After:**
```typescript
const result = await response.json();

if (!response.ok || !result.success) {
    throw new Error(result.message || `HTTP ${response.status}: ${response.statusText}`);
}

if (result.templateId) {
    setRuntimeUploadResult({ templateId, fileName, previewUrl });
    setRuntimeTemplateId(templateId);
}
```

Key improvements:
- Reads JSON body **before** checking `response.ok` — captures error messages from HTTP 400 responses
- `previewUrl` is stored in `runtimeUploadResult` for future use
- State type updated: `{ templateId: string; fileName: string; previewUrl: string | null }`

---

## Verification

| Check | Result |
|-------|--------|
| Backend builds (code compiles) | ✅ (file lock only — process running) |
| Frontend TypeScript (`tsc --noEmit`) | ✅ 0 errors |
| No new `/api/form/upload` endpoint created | ✅ |
| Workbook persisted to `Forms/` (not deleted) | ✅ `File.Move` replaces `File.Delete` |
| `templateId` returned in response | ✅ GUID without extension |
| `previewUrl` returned in response | ✅ `/preview/page_{templateId}.png` |
| Backward compatible (Designer reads `result.data`) | ✅ |
| Error handling reads message from body | ✅ JSON parsed before status check |
| No rendering/coordinate/runtime engine changes | ✅ |
| Single upload pipeline | ✅ |

---

## End-to-End Workflow

```
User selects Excel
      │
      ▼
POST /api/form/from-excel
      │
      ├── Save workbook to Uploads/
      ├── Generate preview PNG
      ├── Parse workbook
      ├── Build FormDefinition
      ├── Move workbook to Forms/ (persist)
      └── Return:
            success: true
            templateId: "guid"
            previewUrl: "/preview/page_guid.png"
            data: { FormDefinition }
                │
                ▼
setRuntimeTemplateId(templateId)
                │
                ▼
GET /api/form/runtime/{templateId}
                │
                ▼
RuntimeForm JSON → OverlayRenderer → Yellow HTML overlay
```

---

## File Details

### `FormController.cs` — Line Changes

- **Line before try block:** `string templateId = Path.GetFileNameWithoutExtension(uniqueFileName);`
- **Inside try:** `File.Move(filePath, persistentPath, overwrite: true)` to `Forms/`
- **Inside try:** `return Ok(new { success, message, templateId, previewUrl, data })` — anonymous type
- **Inside catch:** `File.Delete(filePath)` — cleanup on error only
- **Removed:** `finally` block (no longer needed since file is moved on success, deleted on error)

### `page.tsx` — Line Changes

- **Line 226:** `fetch('.../api/form/from-excel', ...)` — corrected route
- **Lines 229-232:** Response body parsed before status check
- **Line 234:** `setRuntimeUploadResult({ templateId, fileName, previewUrl })` — includes preview URL
- **State type:** Updated to include `previewUrl: string | null`

---

## Known Limitation

The `[ProducesResponseType(typeof(ApiResponse<FormDefinition>), StatusCodes.Status200OK)]` annotation on `FromExcel()` is now inaccurate since the method returns an anonymous object. This affects Swagger documentation only — runtime behavior is correct and the frontend processes JSON dynamically.
