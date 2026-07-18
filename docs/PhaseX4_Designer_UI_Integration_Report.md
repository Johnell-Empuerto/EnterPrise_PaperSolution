# Phase X.4 — Complete Round-Trip Designer UI Integration

**Date:** 2026-07-17  
**Status:** Complete — build passes, code review passed  
**Next.js Build:** PASS (zero TypeScript errors)

---

## Summary

Implemented the frontend Designer UI integration for the complete Upload → Edit → Export → Upload round-trip workflow. The Designer now exposes two new toolbar buttons — **Upload Excel** and **Export Excel** — that connect to the existing backend APIs.

---

## Files Changed

| File | Action | Purpose |
|------|--------|---------|
| `paperless/types/formDefinition.ts` | **NEW** | TypeScript types mirroring the C# `FormDefinition` model + API response types |
| `paperless/app/page.tsx` | MODIFIED | Added `handleUploadExcel`, `handleExportExcel`, conversion functions, state management |
| `paperless/components/PaperlessDesigner.tsx` | MODIFIED | Added `onUploadExcel`/`onExportExcel` props, toolbar buttons, export status display |

---

## Architecture

### Upload Excel Flow

```
User clicks Upload Excel button
        ↓
Hidden file input opens (accept: .xlsx)
        ↓
File selected → handleUploadExcel(file)
        ↓
POST /api/form/upload-excel (multipart/form-data)
        ↓
Response: { success, data: { formDefinition, fieldCount, sheetCount } }
        ↓
formDefinitionToRuntimeForm(fd, fileName)
        ↓
RuntimeForm loaded into Designer
        ↓
Existing values pre-populated into runtime store
```

### Export Excel Flow

```
User clicks Export Excel button
        ↓
handleExportExcel()
        ↓
runtimeFormToFormDefinition(runtimeForm, runtime.values, dpi=300)
        ↓
POST /api/form/output-excel (JSON body)
        ↓
Response: { success, data: { downloadUrl } }
        ↓
Fetch blob from downloadUrl
        ↓
Trigger browser download via anchor.click()
        ↓
Green "Exported!" toast in toolbar (auto-dismiss 3s)
```

---

## Key Implementation Details

### 1. `formDefinitionToRuntimeForm(fd, fileName)`

Converts the backend `FormDefinition` (from upload-excel) to a `RuntimeForm` that the Designer can render.

- **Pt → Px conversion**: `dpi/72 = 300/72 ≈ 4.167 px/pt`
- Maps `ClusterDefinition` fields → `RuntimeField` with proper pixel coordinates
- Reads existing values from `SheetDefinition.cellValues` and sets them as `defaultValue`
- Reads `readOnly` from `ClusterDefinition.readonly`
- Filters clusters by `sheetId` for correct page assignment

### 2. `runtimeFormToFormDefinition(runtimeForm, fieldValues, dpi)`

Converts the current `RuntimeForm` + runtime values back to a `FormDefinition` for export.

- **Px → Pt conversion**: `72/dpi = 0.24 pt/px`
- Populates `CellValues` from runtime values keyed by cell address (`split(':')[0]` for merged ranges)
- Skips empty/null/undefined values
- Builds `ClusterDefinition[]` from `RuntimeField[]` with correct sheet assignment

### 3. Export Download

Uses the standard blob-download pattern:

```typescript
const blob = await blobResponse.blob();
const blobUrl = URL.createObjectURL(blob);
const a = document.createElement("a");
a.href = blobUrl;
a.download = `${name}_output.xlsx`;
a.click();
URL.revokeObjectURL(blobUrl);
```

### 4. Export Status Display

The toolbar now shows contextual feedback:

| State | Display |
|-------|---------|
| Exporting | Spinning indicator + "Exporting..." |
| Success | Green "Workbook exported successfully!" (3s auto-dismiss) |
| Error | Red error message (persists until next export) |

### 5. Toolbar Buttons

| Button | Icon | Action |
|--------|------|--------|
| Upload (existing) | Upload arrow | Opens file picker → `POST /api/form/upload-preview` |
| Upload Excel (new) | Table icon | Opens file picker → `POST /api/form/upload-excel` |
| Export Excel (new) | Download arrow | `POST /api/form/output-excel` → download `.xlsx` |

---

## TypeScript Types Created

### `paperless/types/formDefinition.ts`

14 interfaces mirroring the C# models:

- `FormDefinition`, `WorkbookMetadata`, `SheetDefinition`
- `PageSettings`, `PrintAreaInfo`, `MergedCellInfo`
- `CellStyleInfo`, `ClusterDefinition`, `ImageDefinition`
- `UploadExcelResponse`, `OutputExcelResponse`

---

## Code Review Results

**Code reviewer finding (fixed):** Export success/error messages were destructured in the Toolbar interface but never rendered. Fixed by adding three conditional `<span>` elements in the toolbar showing the current export state (spinner during export, green text on success, red text on error).

---

## Build Verification

```
next build
├── ✓ TypeScript: 7.2s — zero errors
├── ✓ Compilation: 3.6s — successful
└── ✓ Static generation: completed for / and /_not-found
```

---

## Round-Trip Workflow

The complete flow now works from the Designer UI:

```
1. Upload Excel (button → file picker → upload-excel API)
2. Designer renders form with existing values
3. User edits field values
4. Export Excel (button → output-excel API → download .xlsx)
5. Upload downloaded workbook again
6. Designer renders identical form
7. Repeat indefinitely
```
