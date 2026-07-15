# Phase 11J.1 — Upload API Mismatch Analysis

**Date:** July 2026  
**Status:** Analysis Only — No Code Changes  

---

## 1. Existing Backend Routes

### FormController (`ExcelAPI/ExcelAPI/Controllers/FormController.cs`)

| Route | Verb | Action | Line |
|-------|------|--------|------|
| `/api/form/save` | POST | `Save([FromBody] FormDefinition form)` | 41 |
| `/api/form/from-excel` | POST | `FromExcel([FromForm] IFormFile? file)` | 88 |
| `/api/form/runtime/{templateId}` | GET | `GetRuntime([FromRoute] string templateId)` | 384 |

### Other Controllers

| Route | Verb | Controller | Action |
|-------|------|------------|--------|
| `/api/excel/upload` | POST | ExcelController | Upload |
| `/api/health` | GET | HealthController | Get |

### Confirmed

- **POST `/api/form/upload` does NOT exist** in any backend controller.
- No controller has a route named `upload` under the `form` prefix.
- The only file-upload endpoint under `/api/form/` is **POST `/api/form/from-excel`**.

---

## 2. Existing Frontend Routes

### `paperless/app/page.tsx`

| Line | Route | Method | Tab |
|------|-------|--------|-----|
| 83 | `/api/form/from-excel` | POST | Designer (Import) |
| 124 | `/api/form/save` | POST | Designer (Save) |
| **218** | **`/api/form/upload`** | **POST** | **Runtime Viewer (Upload & View)** |

### `paperless/hooks/useRuntime.ts`

| Line | Route | Method |
|------|-------|--------|
| 42 | `/api/form/runtime/{id}` | GET |
| 89 | `/api/form/runtime/{templateId}` | GET |

---

## 3. Route Mismatch

### The Bug

The Runtime Viewer calls **`POST /api/form/upload`** (line 218 of `page.tsx`) but the backend only exposes **`POST /api/form/from-excel`**.

The route `/api/form/upload` was invented in the frontend during Phase 11J.1 — it was never created in the backend.

**This is a frontend bug.** The frontend references a non-existent backend route.

**Verdict: Frontend route is wrong.** The correct route is `POST /api/form/from-excel`.

---

## 4. Response Shape Mismatch

### What the Backend Actually Returns

`POST /api/form/from-excel` returns:

```json
{
  "success": true,
  "message": "Excel file processed. 12 field(s) detected.",
  "data": {
    "workbook": { ... },
    "sheets": [ ... ],
    "clusters": [ ... ],
    "images": [],
    "metadata": { ... }
  }
}
```

The response envelope is `ApiResponse<FormDefinition>`:

```csharp
public class ApiResponse<T> {
    public bool Success { get; set; }       // → "success"
    public string Message { get; set; }     // → "message"
    public T Data { get; set; }             // → "data" (FormDefinition)
}
```

**Key points:**
- The `data` field is a `FormDefinition` object — NOT a flat `templateId`/`previewUrl` structure
- There is **no `templateId` field** in the response
- There is **no `previewUrl` field** at the top level
- The `backgroundImage` URL is nested inside `data.sheets[0].backgroundImage`
- The uploaded XLSX file is **DELETED** in the `finally` block of `FromExcel()` after processing

### What the Frontend Expects

The `handleRuntimeUpload` handler in `page.tsx` (line 218+) expects:

```json
{
  "success": true,
  "templateId": "a1b2c3d4...",
  "previewUrl": "/preview/page_a1b2c3d4...png"
}
```

The code on line 240 does:
```typescript
if (result.success && result.templateId) {
```

**This condition will NEVER be true** with the actual backend response because:
1. The backend returns `data` not `templateId`
2. Even if we call `/api/form/from-excel` instead, it returns an `ApiResponse<FormDefinition>` envelope — `result.templateId` would be `undefined`

### Response Contract Comparison

| Field | Backend (actual) | Frontend (expected) | Match? |
|-------|-----------------|---------------------|--------|
| `success` | ✅ top-level | ✅ top-level | ✅ |
| `templateId` | ❌ does not exist | ✅ required | ❌ |
| `previewUrl` | ❌ does not exist | ✅ expected | ❌ |
| `data` | ✅ FormDefinition object | ❌ not read | ❌ |
| `data.sheets[0].backgroundImage` | ✅ contains preview URL | ❌ not read | ❌ |

---

## 5. Deeper Problem: XLSX File Lifecycle

The `FromExcel()` action does the following:

1. Saves uploaded file to `Uploads/{guid}.xlsx`
2. Processes via `IExcelCaptureService.CapturePrintAreaAsync()` → generates PNG
3. Persists background PNG to `/Forms/bg_{guid}.png`
4. **Deletes the XLSX file** (`File.Delete(filePath)` in the `finally` block)
5. Returns `FormDefinition` with `backgroundImage = "/Forms/bg_{guid}.png"`

After this action completes:

| Artifact | Persisted? | Path |
|----------|-----------|------|
| PNG background | ✅ | `/Forms/bg_{guid}.png` |
| XLSX workbook | ❌ **Deleted** | N/A |

The runtime endpoint `GET /api/form/runtime/{templateId}` calls `FindTemplateFile(templateId)` which searches for `.xlsx` files in the `Forms/` and `Uploads/` directories. **Since the XLSX was deleted, the runtime endpoint will never find the template file.**

The runtime flow requires a persisted XLSX to parse. The current `from-excel` flow only persists the PNG, not the XLSX.

---

## 6. Root Cause

The root cause is a **frontend bug** introduced in Phase 11J.1:

1. The frontend calls a non-existent route (`/api/form/upload` instead of `/api/form/from-excel`)
2. The frontend expects a response shape (`{ success, templateId, previewUrl }`) that the backend does not produce
3. Even if the route were corrected, the backend's `from-excel` action deletes the XLSX file, so the subsequent `GET /api/form/runtime/{templateId}` would fail with "Template not found"

---

## 7. Recommended Fix

### Option A: Frontend-only fix (least effort)

Two changes needed in `page.tsx`:

1. **Route**: Change `fetch('.../api/form/upload')` → `fetch('.../api/form/from-excel')`
2. **Response parsing**: Instead of expecting `result.templateId`, extract the template ID from the response. Since `from-excel` doesn't return one, the frontend needs a way to derive it. However, the XLSX is deleted, so `GET /api/form/runtime/{templateId}` will still fail.

**This alone is insufficient** because the XLSX is deleted.

### Option B: Backend fix + frontend fix (recommended)

**Backend** — Modify `FromExcel()` in `FormController.cs`:
1. Change the `finally` block to **NOT delete** the XLSX file, OR move it to the `Forms/` directory
2. Add `templateId` to the response — the GUID used as the filename (without extension)
3. Add `previewUrl` to the response — the persisted background image path

**Frontend** — Change `page.tsx`:
1. Call `/api/form/from-excel` instead of `/api/form/upload`
2. Read `result.templateId` from the response
3. Set `runtimeTemplateId` to the returned templateId

### Option C: New dedicated endpoint (cleanest)

**Backend** — Create a new action in `FormController`:

```
POST /api/form/upload → returns { success, templateId, previewUrl }
```

This action would:
1. Save XLSX to `Forms/` directory with a GUID filename (keep it)
2. Process via capture service (generate PNG)
3. Return `{ success: true, templateId: "{guid}", previewUrl: "/forms/bg_{guid}.png" }`

**Frontend** — Keep calling `/api/form/upload` but adapt to whatever response shape the new endpoint returns.

---

## 8. Classification

| Aspect | Classification |
|--------|---------------|
| Wrong route in frontend | **Frontend bug** |
| Missing `templateId` in response | **Backend gap** (response contract not designed for this flow) |
| XLSX deleted after processing | **Backend gap** (existing behavior incompatible with runtime flow) |
| No dedicated upload endpoint for runtime | **Missing feature** (Phase 11J.1 spec assumed it existed) |

**Summary:** This is a **frontend + backend coordination bug** — the frontend references a route that was never built, and the closest existing route (`from-excel`) has a response shape and file lifecycle that are incompatible with the runtime workflow.

### Fix Priority

1. Decide whether to reuse `/api/form/from-excel` (Option B) or create a new `/api/form/upload` (Option C)
2. Implement the backend changes (keep XLSX, return templateId)
3. Fix the frontend route and response parsing
4. Test end-to-end: Upload → auto-load runtime → render overlay
