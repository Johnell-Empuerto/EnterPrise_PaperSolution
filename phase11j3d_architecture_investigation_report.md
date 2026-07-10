# Phase 11J.3D — Runtime Preview Pipeline & UI Architecture Investigation

**Date:** July 2026  
**Status:** Investigation Complete — No Code Changes  

---

## Part 1 — Runtime Preview Pipeline

*(Findings from Phase 11J.3C, confirmed and expanded)*

### Complete File Generation Trace

```
POST /api/form/from-excel
    │
    ├── (1) File saved: Uploads/{uploadGuid}.xlsx
    │     templateId = uploadGuid
    │
    ├── (2) CapturePrintAreaAsync() — ExcelCaptureService.cs:505-512
    │     ├── fileId = Guid.NewGuid().ToString("N")       ← NEW RANDOM GUID
    │     ├── previewFileName = $"page_{fileId}.png"
    │     └── PNG saved to:  Preview/page_{fileId}.png    ✓ FILE EXISTS
    │                        Return: ImageUrl = "/preview/page_{fileId}.png"
    │
    ├── (3) FromExcel() constructs previewUrl — FormController.cs:146
    │     previewUrl = $"/preview/page_{templateId}.png"   ✗ WRONG GUID
    │     → Uses upload Guid, NOT the capture Guid
    │
    ├── (4) ConvertCaptureToForm() copies PNG — FormController.cs:270-283
    │     ├── bgFileName = $"bg_{Guid.NewGuid():N}.png"   ← THIRD GUID
    │     └── Copy saved:  Forms/bg_{copyGuid}.png        ✓ FILE EXISTS
    │                       backgroundImage = "/forms/bg_{copyGuid}.png"
    │
    ├── (5) Response returned with TWO different image URLs
    │     ├── previewUrl:   "/preview/page_{uploadGuid}.png"   ✗ 404
    │     └── data.sheets[0].backgroundImage: "/forms/bg_{copyGuid}.png"  ✓
    │
    └── (6) Frontend FormPage.tsx:18-19 reconstructs URL
          img.src = `${API_BASE_URL}/preview/page_${templateId}.png`  ✗ 404
          → Ignores both backend-provided previewUrl and backgroundImage
```

### Root Cause

| File | Line | Issue |
|------|------|-------|
| `ExcelCaptureService.cs` | 507 | Uses its own `Guid.NewGuid()` for PNG filename — no access to `templateId` |
| `FormController.cs` | 146 | Guesses PNG path from `templateId`, but actual file uses capture GUID |
| `FormController.cs` | 271 | Copies PNG to `Forms/bg_{anotherGuid}.png` — third GUID unnecessarily |
| `FormPage.tsx` | 18 | Reconstructs URL from `templateId` instead of consuming backend-provided URL |

### The Three GUID Problem

| Scope | Variable | File | Exists? |
|-------|----------|------|---------|
| templateId (upload) | `uniqueFileName` GUID | `Forms/{templateId}.xlsx` | ✅ |
| capture fileId | `fileId` = `Guid.NewGuid()` | `Preview/page_{fileId}.png` | ✅ |
| copy bgGuid | `Guid.NewGuid()` | `Forms/bg_{bgGuid}.png` | ✅ |
| previewUrl | `$"/preview/page_{templateId}.png"` | — | ❌ **404** |

### Verification

- `Preview/` directory contains `page_*.png` files — all with fileId GUIDs
- No `page_{templateId}.png` exists because the two GUIDs are independently generated
- `Forms/bg_*.png` files exist as copies made by `ConvertCaptureToForm()`

---

## Part 2 — UI Architecture Review

### 2.1 Current Frontend Structure

```
paperless/app/page.tsx  ← SINGLE PAGE, NO ROUTER
         │
         ├── Tab: "Designer"
         │     ├── Upload Excel  → POST /api/form/from-excel
         │     ├── Save          → POST /api/form/save
         │     ├── New Form      → reset
         │     ├── Background PNG display
         │     ├── Cluster overlays (colored rectangles, clickable, selectable)
         │     ├── Guide lines on selection
         │     └── Legend bar
         │
         └── Tab: "Runtime Viewer"
               ├── Upload Excel  → POST /api/form/from-excel  (same endpoint!)
               ├── Advanced: manual template ID input
               ├── Background PNG display
               ├── Yellow HTML overlay fields (TextField, NumberField, etc.)
               └── Focused field info bar

paperless/components/
    └── FormViewer/  ← Runtime Viewer components only
          ├── FormPage.tsx
          ├── OverlayRenderer.tsx
          ├── OverlayField.tsx
          └── fields/
                ├── TextField.tsx, NumberField.tsx, DateField.tsx
                ├── CheckboxField.tsx, DropdownField.tsx, SignatureField.tsx
```

### 2.2 DesignerController Capability Analysis

**Confirmed: The DesignerController has methods that are NEVER called from the UI.**

| DesignerController method | Called from UI? | Evidence |
|--------------------------|----------------|----------|
| `loadForm()` | ✅ Yes — after upload | `page.tsx` handleUpload |
| `save()` | ✅ Yes — Save button | `page.tsx` handleSave |
| `resetForm()` | ✅ Yes — New Form button | `page.tsx` handleReset |
| `selectCluster()` | ✅ Yes — clicking clusters | `page.tsx` handleClusterClick |
| `undo()` | ✅ Yes — Undo button | `page.tsx` |
| `addCluster()` | ❌ **No** | No UI for manual cluster creation |
| `removeCluster()` | ❌ **No** | No UI for removing clusters |
| `updateCluster()` | ❌ **No** | No UI for editing cluster properties |
| `addSheet()` | ❌ **No** | No multi-sheet UI |
| `setSheetName()` | ❌ **No** | No sheet rename UI |
| `setPrintArea()` | ❌ **No** | No print area editing UI |
| `setPageSettings()` | ❌ **No** | No page settings editing UI |
| `setRowHeight()` | ❌ **No** | No row height editing UI |
| `setColumnWidth()` | ❌ **No** | No column width editing UI |
| `addMergedCell()` | ❌ **No** | No merge editing UI |
| `removeMergedCell()` | ❌ **No** | No merge editing UI |
| `setCellStyle()` | ❌ **No** | No cell style editing UI |
| `setCellValue()` | ❌ **No** | No cell value editing UI |
| `setFreezePane()` | ❌ **No** | No freeze pane UI |
| `addImage()` | ❌ **No** | No image addition UI |
| `removeImage()` | ❌ **No** | No image removal UI |
| `getClustersForSheet()` | ❌ **No** | Not called anywhere |
| `getFormSnapshot()` | ❌ **No** | Internal only |
| `redo()` | ❌ **No** | No Redo button |
| `beginBatch()` / `endBatch()` | ❌ **No** | No batch operations |

**Approximately 22 of 26 methods (85%) are dead code.**

### 2.3 Backend Endpoint Usage Analysis

| Endpoint | Used by Designer? | Used by Runtime? | Still needed if Designer removed? |
|----------|-------------------|-------------------|-----------------------------------|
| `POST /api/form/from-excel` | ✅ Yes | ✅ Yes | ✅ Yes — necessary for upload |
| `POST /api/form/save` | ✅ Yes | ❌ No | **Probably not** — templates are persisted during upload |
| `GET /api/form/runtime/{id}` | ❌ No | ✅ Yes | ✅ Yes — necessary for runtime |
| `POST /api/excel/upload` | ❌ No | ❌ No | ❌ No — unused (use from-excel instead) |
| `GET /api/health` | ❌ No | ❌ No | ⚠️ Optional — monitoring |

### 2.4 Feature Comparison

| Feature | Designer | Runtime | Duplicated? | Keep? |
|---------|----------|---------|-------------|-------|
| Upload Excel | ✅ Yes | ✅ Yes | **YES** | ✅ One upload flow needed |
| Display background PNG | ✅ Yes | ✅ Yes | **YES** | ✅ One display needed |
| Zoom controls | ✅ Yes | ✅ Yes | **YES** | ✅ One zoom needed |
| Parse workbook | ✅ (backend) | ✅ (backend) | No (same endpoint) | ✅ Backend only |
| Detect fields | ✅ (clusters) | ✅ (runtime fields) | **Partially** | ✅ Unified to runtime fields |
| Interactive form fields | ❌ No | ✅ Yes (yellow overlays) | No | ✅ Runtime only |
| Cluster overlays | ✅ Colored rectangles | ❌ No | No | ❌ Replace with runtime fields |
| Field data entry | ❌ No | ✅ Yes (text, checkbox, etc.) | No | ✅ Runtime only |
| Save form definition | ✅ POST /api/form/save | ❌ No | No | ❌ See analysis below |
| Manual cluster editing | ❌ No (no UI exists) | ❌ No | No | ❌ Not needed |
| Manual cell editing | ❌ No (no UI exists) | ❌ No | No | ❌ Not needed |
| Background-less fallback | ✅ Yes | ✅ Yes | **YES** | ✅ One fallback needed |
| Advanced template ID entry | ❌ No | ✅ Yes | No | ✅ Keep in unified view |
| Focused field info | ❌ No | ✅ Yes | No | ✅ Keep |
| Undo button | ✅ Yes | ❌ No | No | ❌ Not needed (no manual edits) |
| Legend bar | ✅ Yes | ❌ No | No | ❌ Not needed |
| Guide lines on selection | ✅ Yes | ❌ No | No | ❌ Not needed |

### 2.5 Duplicate Functionality

The following are **directly duplicated** between the two tabs:

1. **Excel upload** — Both tabs call `POST /api/form/from-excel` with identical form data
2. **Background PNG display** — Both tabs show the rendered background
3. **Zoom controls** — Both tabs have independent zoom state
4. **File selection UI** — Both tabs have their own file input

The only **unique** functionality in Designer is:
- `POST /api/form/save` (which persists a FormDefinition that's already persisted during upload)
- Cluster overlays (colored rectangles showing detected fields — a preview of what the backend found)

The only **unique** functionality in Runtime Viewer is:
- Yellow HTML overlay fields (interactive form inputs)
- Auto-load of runtime JSON after upload
- Focused field info bar
- Advanced template ID input

### 2.6 Navigation Analysis

Current navigation uses a client-side tab switcher (`activeTab` state in `page.tsx`):

```
<button> Designer </button>
<button> Runtime Viewer </button>
```

There is NO router (no Next.js `/page` routing). Both views are in the same file, toggled by state. This is simpler than real routing but means both feature sets share the same URL.

### 2.7 Save Flow Analysis

**Current flow:**
1. Upload → Backend processes → Returns `FormDefinition` with `backgroundImage`
2. User can click **Save** → `POST /api/form/save` → Backend generates XML/XLSX/database objects
3. But templates are ALREADY persisted during upload (workbook moved to Forms/)

**If Designer is removed:**
- The `POST /api/form/save` endpoint becomes unnecessary
- Templates are created and persisted during the upload step
- The upload IS the save operation — the workbook is in Forms/, the PNG is in Preview/
- Runtime JSON is generated on-demand from the persisted workbook

### 2.8 Component Inventory

**Components that would become obsolete if Designer is removed:**

| Component/Feature | File | Reason |
|------------------|------|--------|
| `DesignerController` (entire class) | `paperless/lib/designerController.ts` | 85% dead code; cluster management not used |
| `FormDefinition` type imports in page.tsx | `paperless/app/page.tsx` | Only needed for Designer |
| Designer upload form | `page.tsx` lines 636-728 | Duplicated by Runtime upload |
| Save button | `page.tsx` lines 727-770 | Save not needed (upload persists template) |
| Cluster overlay rendering | `page.tsx` lines 980-1064 | Replaced by Runtime overlay fields |
| Guide lines | `page.tsx` lines 1042-1075 | Only needed for cluster positioning |
| Legend bar | `page.tsx` lines 1082-1120 | Only needed for browsing clusters |
| Undo button | `page.tsx` lines 90-103 | No manual edits to undo |
| `handleClusterClick` | Named export | No clusters to click |
| `ApiResponse<FormDefinition>` interface | Local type | Only used by Designer upload |

**Components that remain if Designer is removed:**

| Component | File | Purpose |
|-----------|------|---------|
| `useRuntime` hook | `paperless/hooks/useRuntime.ts` | Runtime JSON loading |
| `useFieldState` hook | `paperless/hooks/useFieldState.ts` | Field value management |
| `FormPage` | `paperless/components/FormViewer/FormPage.tsx` | Background + overlay |
| `OverlayRenderer` | `paperless/components/FormViewer/OverlayRenderer.tsx` | Field positioning |
| `OverlayField` | `paperless/components/FormViewer/OverlayField.tsx` | Field type dispatch |
| All field components | `paperless/components/FormViewer/fields/*` | Interactive inputs |
| `RuntimeForm` types | `paperless/types/runtime.ts` | Type definitions |
| `formDefinition.ts` | Dead after DesignerController removed | Only used by Designer |

---

## Part 3 — Recommended Target Architecture

### Recommendation: Option C — Remove Designer Completely

**Rationale:**
1. The backend already does ALL the design work (parsing, field detection, coordinate calculation, PNG generation)
2. 85% of DesignerController methods are dead code — they never execute
3. Both tabs call the identical upload endpoint
4. The Designer has no actual editing capability (no canvas, no toolbar, no property editor)
5. The Runtime Viewer already has the upload flow — it subsumes everything Designer does
6. The `POST /api/form/save` endpoint is unnecessary when upload already persists templates

### Target Architecture

```
Excel Workbook
    │
    ▼
Upload  ──→  POST /api/form/from-excel
    │
    ├── Save  Forms/{templateId}.xlsx
    ├── Generate  Preview/page_{templateId}.png
    ├── Parse workbook via OpenXmlParser
    ├── Detect fields via FieldDetector
    ├── Build RuntimeForm via FormRuntimeBuilder
    └── Return  { templateId, previewUrl, data }
                    │
                    ▼
Runtime Viewer
    │
    ├── Load background: /preview/page_{templateId}.png
    ├── Load runtime:   GET /api/form/runtime/{templateId}
    ├── Render fields:  Yellow HTML overlay
    └── User fills form → no intermediate design phase
```

### Single Source of Truth

```
templateId
    │
    ├── Forms/{templateId}.xlsx       ← Workbook
    ├── Preview/page_{templateId}.png ← Preview image
    └── GET /api/form/runtime/{id}    ← Runtime JSON (generated on-demand)
```

**All three artifacts use the same identifier.** No duplicated GUIDs.

### Required Changes Summary

| Area | Change |
|------|--------|
| `ExcelCaptureService.cs` | Accept `templateId` as parameter for PNG filename (or return `fileId` so caller can use it) |
| `FormController.cs` | Use `captureResult.ImageUrl` for `previewUrl` instead of constructing from `templateId` |
| `FormController.cs` | Remove `POST /api/form/save` endpoint (or keep for backward compat but mark deprecated) |
| `FormPage.tsx` | Consume backend-provided `previewUrl` instead of reconstructing from `templateId` |
| `page.tsx` | Remove Designer tab, remove DesignerController, merge tabs into single Runtime view |
| `designerController.ts` | Remove file entirely (all dead code) |
| `formDefinition.ts` | Remove file (no longer needed) |

### Backend Endpoint Changes

- **KEEP:** `POST /api/form/from-excel` — upload + parse + persist
- **KEEP:** `GET /api/form/runtime/{id}` — runtime JSON generation
- **REMOVE or DEPRECATE:** `POST /api/form/save` — no longer needed

### Frontend Component Changes

- **REMOVE:**
  - `designerController.ts` (entire file — 85% dead code)
  - `formDefinition.ts` (entire file — Designer-only types)
  - All Designer-related UI in `page.tsx`
  - Save button, Undo button, Cluster overlays, Guide lines, Legend bar
- **CONSOLIDATE:**
  - Single upload flow (keep Runtime's upload, remove Designer's)
  - Single background PNG display
  - Single zoom control
- **FIX:**
  - `previewUrl` construction in `FormController.cs`
  - `previewUrl` consumption in `FormPage.tsx`

---

## Part 4 — Facts vs. Recommendations

### Confirmed Facts

| Fact | Evidence |
|------|----------|
| Designer and Runtime both call `POST /api/form/from-excel` | `page.tsx` line 83 and line 226 |
| DesignerController has 22 unused methods out of 26 | Code analysis of `designerController.ts` |
| No separate Designer component files exist | `ls paperless/components` shows only `FormViewer/` |
| `previewUrl` uses wrong GUID | `FormController.cs` line 146 vs `ExcelCaptureService.cs` line 507 |
| Three different GUIDs for one image | Code trace in Part 1 |
| Templates are persisted during upload | `FormController.cs` `FromExcel()` — `File.Move()` to Forms/ |
| Runtime fields work (6 detected in test) | Verified from Phase 11J.3B testing |

### Recommendations

| Recommendation | Basis |
|---------------|-------|
| Remove Designer tab | No unique editing functionality; 85% dead code |
| Remove DesignerController | All methods dead or subsumed by Runtime |
| Remove `POST /api/form/save` | Upload already persists the template |
| Fix `previewUrl` to use actual generated path | `captureResult.ImageUrl` already contains correct path |
| Accept `templateId` in `CapturePrintAreaAsync` | Ensures single GUID across all artifacts |
| Remove `formDefinition.ts` | Designer-only types, unused by Runtime |
