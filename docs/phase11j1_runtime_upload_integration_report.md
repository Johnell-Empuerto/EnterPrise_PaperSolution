# Phase 11J.1 — Runtime Upload Integration (Testing Workflow)

**Date:** July 2026  
**Target Build:** 0 TypeScript errors  
**Status:** ✅ Complete  

---

## Objective

Extend the existing Runtime Viewer so it can upload an Excel workbook directly for testing — no manual Template ID entry required.

---

## Changes Summary

### Modified Files

| File | Change |
|------|--------|
| `hooks/useRuntime.ts` | Added `loadByTemplateId(templateId)` programmatic loading method alongside reactive `templateId` prop. Includes `doFetch` internal helper and `loadedTemplateId` tracking. |
| `app/page.tsx` | Replaced the Template ID textbox with an Upload Excel workflow. Template ID textbox moved to an "Advanced — Manual Template ID" collapsible section. |

---

## Architecture

### New Upload Flow

```
Runtime Viewer
  │
  ├── Choose Excel File
  │
  ├── POST /api/form/upload (multipart/form-data)
  │     │
  │     └── Returns { success, templateId, previewUrl }
  │
  ├── setRuntimeTemplateId(templateId)
  │     │
  │     └── useRuntime useEffect auto-fetches GET /api/form/runtime/{id}
  │
  ├── Load Preview PNG (/preview/page_{templateId}.png)
  │
  └── Render Yellow Overlay (via FormPage + OverlayRenderer)
```

### No Manual Entry Required

After upload, the runtime loads automatically. The Template ID textbox is hidden inside an `<details>` collapsible section labeled "Advanced — Manual Template ID" for debugging purposes.

---

## State Management

### Upload States

| State | Variable | Description |
|-------|----------|-------------|
| Idle | `runtimeUploadFile = null` | No file selected |
| File Selected | `runtimeUploadFile = File` | File chosen, Upload button enabled |
| Uploading | `runtimeUploading = true` | POST in progress, spinner shown |
| Upload Complete | `runtimeUploadResult = { templateId, fileName }` | Success — auto-loads runtime |
| Upload Failed | `runtimeUploadError = string` | Error message displayed |

### Upload Error Handling

| Scenario | Handling |
|----------|----------|
| No file selected | `setRuntimeUploadError("Please select an Excel file first.")` |
| Invalid extension (not .xlsx/.xls) | `setRuntimeUploadError("Invalid file extension...")` |
| HTTP error during upload | Catches fetch error, displays message |
| Upload returns no templateId | `throw new Error("Upload failed — no template ID returned.")` |
| Runtime fetch failure | Existing `useRuntime` error state (shown separately) |
| Preview image not found | `FormPage` `onError` handler shows fallback |
| Template not found | `useRuntime` returns error, displayed in red banner |

---

## UI Components (Runtime Viewer Tab)

### Upload Card
- File input (`.xlsx, .xls` accept)
- File name + size indicator after selection
- **"Upload & View"** button — gradient green, loading spinner during upload
- **"New Form"** reset button — shown after successful upload
- **"Advanced — Manual Template ID"** collapsible section with textbox + Load button

### Upload Info Banner
- Shown after successful upload
- Displays: file name, sheet count, field count, DPI, template ID (monospace)

### Error Banners
- **Upload error** (red) — file validation or HTTP failure
- **Runtime error** (red) — runtime fetch failure (hidden during upload)

### Form Viewer (unchanged from Phase 11J)
- Info bar: workbook name, sheets, fields, DPI
- Zoom controls (-, %, +)
- Background PNG + yellow overlay fields
- Focused field info bar

---

## Constraints Enforced

| Rule | Status |
|------|--------|
| No Rendering Engine changes | ✅ |
| No Runtime Engine changes | ✅ |
| No coordinate recalculations | ✅ |
| Coordinates from RuntimeField.LeftPx/TopPx/WidthPx/HeightPx | ✅ |
| Existing Designer unchanged | ✅ |
| Existing Runtime API unchanged | ✅ |
| No new backend endpoints | ✅ (uses existing `/api/form/upload`) |

---

## Validation Results

| Check | Result |
|-------|--------|
| TypeScript compilation (`tsc --noEmit`) | ✅ 0 errors |
| File extension validation | ✅ |
| Upload flow works without manual template ID | ✅ |
| Template ID accessible via Advanced section | ✅ |
| Existing Designer tab unchanged | ✅ |

---

## File Details

### `hooks/useRuntime.ts` — Key Additions

```typescript
interface UseRuntimeResult {
  // ... existing properties ...
  loadByTemplateId: (templateId: string) => Promise<void>;
}
```

- `doFetch`: Internal memoized fetch function (avoids code duplication)
- `loadByTemplateId`: Sets `loadedTemplateId` state + calls `doFetch`
- `loadedTemplateId`: Tracks programmatic loads separately from reactive `templateId` prop
- Effect prioritizes parent's `templateId` prop over programmatic loads

### `app/page.tsx` — Key State

```typescript
const [runtimeUploadFile, setRuntimeUploadFile] = useState<File | null>(null);
const [runtimeUploading, setRuntimeUploading] = useState(false);
const [runtimeUploadError, setRuntimeUploadError] = useState<string | null>(null);
const [runtimeUploadResult, setRuntimeUploadResult] = useState<{
  templateId: string;
  fileName: string;
} | null>(null);
```

### `app/page.tsx` — Key Handlers

- `handleRuntimeUpload`: Validates → POSTs → sets templateId → effect auto-fetches runtime
- `handleRuntimeReset`: Clears all upload/runtime state, resets file input

---

## Known Limitations

- `loadByTemplateId` in the hook is not used by the upload flow (the reactive `templateId` prop drives fetching via `useEffect`). It's available for future programmatic use cases.
- The `loadedTemplateId` internal state adds some complexity that could be simplified if `loadByTemplateId` is only ever used alongside a parent-set `templateId`.

---

## Next Steps

- **Phase 11K** — Designer Enhancements
- **Phase 11L** — Production Release
