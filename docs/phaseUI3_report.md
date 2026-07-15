# Phase UI.3 — Multi-Worksheet Pagination Implementation Report

## Objective

Enable true multi-worksheet pagination: every printable worksheet in the Excel workbook becomes an independent page in the viewer, with its own background image, its own set of fields, and independent navigation.

## Problem

The existing pipeline only exported the ActiveSheet as the sanitized PDF, returning a flat `fields[]` list. All fields from all worksheets were merged into a single sheet in the frontend, making multi-page navigation impossible.

## Changes

### 1. Backend — `render_service/upload_coordinate_generator.py`

**`render_pdf_to_image()`**
- Added `page_index` parameter (default `0` for backward compatibility)
- Validates `page_index < len(doc)` to prevent out-of-range errors

**`generate_coordinates_and_preview()`** — Major rewrite:

| Aspect | Before | After |
|--------|--------|-------|
| **Comment scan** | Flat `cluster_meta` without sheet tracking | Each entry includes `sheetName` field |
| **Address storage** | Single `cluster_addrs` set | `cluster_addrs_by_sheet` dict (sheetName → set of addresses) |
| **Sanitization** | Applied all addresses to all worksheets | Only applies addresses to the correct worksheet |
| **PDF export** | `ws_active.ExportAsFixedFormat` (ActiveSheet only) | `wb.ExportAsFixedFormat` (ALL visible worksheets) |
| **Non-COM processing** | Single page 0 render + scan | Loop over each visible sheet: render PDF page i → scan → normalize → background PNG |
| **Return value** | `{backgroundImage, page, fields}` | `{pages: [{sheetName, backgroundImage, page, fields}]}` |

**Stale COM proxy fix (critical):**
- `visible_sheet_names` is captured as plain Python strings during the COM phase
- The NC phase iterates by string name instead of COM proxy objects
- Prevents crash after `excel.Quit()` / `pythoncom.CoUninitialize()`

### 2. Backend — `render_service/app.py`

- `/upload/preview` endpoint now returns `{"success": true, "pages": [...]}`
- Added per-page logging for diagnostics

### 3. ASP.NET — `PythonRenderService.cs`

- `PythonPreviewResponse` now has `Pages: List<PythonPreviewPageResult>` instead of flat `BackgroundImage`/`Page`/`Fields`
- Added `PythonPreviewPageResult` class with `SheetName`, `BackgroundImage`, `Page`, `Fields`

### 4. ASP.NET — `FormController.cs`

- `UploadPreview` maps `pythonResult.Pages` to per-page anonymous objects
- Each page includes `sheetName`, `backgroundImage`, `page` (dimensions), and `fields`

### 5. Frontend — `paperless/app/page.tsx`

- Parses `result.pages[]` into `RuntimeForm.sheets[]`
- Each page creates a `RuntimeSheet` with its own:
  - `name` — from `p.sheetName` (falls back to `Page ${i+1}`)
  - `pageWidthPx` / `pageHeightPx`
  - `backgroundImage`
  - `fields[]` — field IDs scoped per-sheet for React key uniqueness

## Architecture

### Data Flow

```
Frontend (Next.js)
  │
  │ POST /api/form/upload-preview
  ▼
ASP.NET FormController.UploadPreview
  │
  │ PythonRenderService.UploadPreviewAsync()
  ▼
Python /upload/preview
  │
  │ generate_coordinates_and_preview()
  ▼
1. Open Excel (1 COM session)
2. Collect visible worksheet names (plain strings)
3. Export original PDF (ALL worksheets)
4. Read comments per-sheet (track sheetName)
5. Sanitize per-sheet (only apply addresses to correct sheet)
6. Export sanitized PDF (ALL worksheets, workbook-level export)
7. Quit Excel
8. For each visible worksheet:
   - Render sanitized PDF page i → scan → normalize
   - Render original PDF page i → background PNG
9. Return {pages: [{sheetName, backgroundImage, page, fields}]}
  │
  ▼
ASP.NET maps pages to per-page response
  │
  ▼
Frontend parses pages[] into RuntimeForm.sheets[]
  │
  ▼
PaperlessDesigner shows multi-page viewer
```

### Edge Cases Covered

| Scenario | Behavior |
|----------|----------|
| Single-sheet workbook | 1 page, same as before |
| Multi-sheet with comments on all | Each sheet = separate page with its own fields |
| Multi-sheet, some have comments | Comment sheets have fields; others appear as blank pages with background only |
| Hidden sheets | Excluded from processing entirely |
| All sheets hidden | Returns `{pages: []}` → frontend shows upload screen |
| No comments on any sheet | Early exit → returns `{pages: []}` |

## Verification

- ✅ **TypeScript compilation**: Zero errors (`npx tsc --noEmit`)
- ✅ **Python syntax**: Both modified files pass `ast.parse()`
- ✅ **Code review**: Approved — no remaining stale COM proxy issues
- ✅ **Backward compatibility**: `/upload/coordinates` endpoint unchanged (uses old `generate_coordinates()`)

## Files Modified

| File | Change |
|------|--------|
| `render_service/upload_coordinate_generator.py` | `render_pdf_to_image(page_index)` + `generate_coordinates_and_preview()` rewrite |
| `render_service/app.py` | `/upload/preview` returns `pages[]` |
| `ExcelAPI/ExcelAPI/Services/PythonRenderService.cs` | `PythonPreviewResponse` with `Pages` list |
| `ExcelAPI/ExcelAPI/Controllers/FormController.cs` | `UploadPreview` maps multi-page response |
| `paperless/app/page.tsx` | Parses `pages[]` into multi-sheet `RuntimeForm` |
