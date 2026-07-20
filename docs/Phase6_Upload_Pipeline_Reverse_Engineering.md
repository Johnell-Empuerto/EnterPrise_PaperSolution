# Phase 6 — Upload Pipeline Reverse Engineering Report

**Date:** July 20, 2026  
**Status:** Investigation Complete (Do Not Modify Code Yet)

---

## 1. Complete Upload Pipeline Diagram

```
User Uploads .xlsx
       │
       ▼
┌──────────────────────────────────────────────────────────────────────┐
│ FormController.UploadPreview()                    [FormController.cs] │
│   1. Saves uploaded file to disk                                      │
│   2. Calls HasEmbeddedConfiguration() (OOXML check)                   │
│   3. Routes to one of two paths:                                      │
└──────────────────────────────────────────────────────────────────────┘
       │
       ├── CONFIG DETECTED ──────────────────────────────────────┐
       │                                                        │
       ▼                                                        ▼
┌─────────────────────────────┐                  ┌─────────────────────────────┐
│ Python Render (background   │                  │ Python Render (full         │
│ image only)                 │                  │ preview + fields)           │
│ POST /upload/preview        │                  │ POST /upload/preview        │
└─────────────────────────────┘                  └─────────────────────────────┘
       │                                                        │
       ▼                                                        ▼
┌─────────────────────────────┐                  ┌─────────────────────────────┐
│ C# WorkbookReaderService    │                  │ Python:                     │
│   .Read()                   │                  │ generate_coordinates_and_   │
│   (COM-based, reconstructs  │                  │   preview()                 │
│    FormDefinition from      │                  │   [upload_coordinate_       │
│    _Fields + comments)      │                  │    generator.py]            │
└─────────────────────────────┘                  └─────────────────────────────┘
       │                                                        │
       ▼                                                        ▼
┌─────────────────────────────┐                  ┌─────────────────────────────┐
│ FormRuntimeBuilder          │                  │ Returns {pages[]} with:    │
│   .BuildFromDefinitionDirect│                  │   sheetName                │
│   → RuntimeForm             │                  │   backgroundImage           │
└─────────────────────────────┘                  │   page dimensions           │
       │                                         │   fields (ratios)           │
       ▼                                         └─────────────────────────────┘
┌─────────────────────────────┐                              │
│ { runtimeForm,              │                              │
│   isReconstructed: true }   │                              ▼
└─────────────────────────────┘                  ┌─────────────────────────────┐
       │                                         │ { pages[],                  │
       ▼                                         │   sessionId }               │
┌───────────────────────────────────────────────────────────────────────────────┐
│                              Frontend                                          │
│                                                                                │
│   if (result.runtimeForm) → use directly (reopened)                           │
│   else → convert result.pages → RuntimeForm (new template)                     │
│                                                                                │
│   PaperlessDesigner → RuntimeFormViewer → RuntimeCanvas → RuntimeField         │
└───────────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Every Worksheet Enumeration Point

### Python — `render_service/upload_coordinate_generator.py`

| Line | Function | Filter | Includes ExcelOutputSetting? |
|------|----------|--------|------------------------------|
| **1077** | `generate_coordinates_and_preview()` — "Collect Visible Sheets" | `ws.Visible != -1` (skips non-visible) | **YES** — ExcelOutputSetting is visible |
| **510** | `_delete_metadata_sheets(wb)` — deletes before PDF export | Only `_Fields`, `_RawData` | **YES** — NOT in the delete list |
| **95** | `_read_comments_native()` — iterates `wb.Worksheets` | `Comments.Count == 0` skip | **YES** — iterated but skipped if no comments |
| **168** | `_read_fields_sheet_data()` — finds `_Fields` by name | Name == "_Fields" | **NO** — only looks for _Fields |

### Python — `render_service/excel_cluster_reader.py`

| Line | Function | Filter | Includes ExcelOutputSetting? |
|------|----------|--------|------------------------------|
| **371** | Cluster detection | `ws.Visible != -1` (skips hidden) | **YES** — if visible, included |

### C# — `ExcelAPI/ExcelAPI/Designer/Analysis/WorkbookReaderService.cs`

| Line | Function | Filter | Includes ExcelOutputSetting? |
|------|----------|--------|------------------------------|
| **115** | `Read()` — sheet classification | `_Fields` captured separately; all others → `contentSheets` | **WAS YES before fix, NOW NO** — ConfigurationSheetNames filter added |
| **509** | `ReadSheetPageSetupFromOoxml()` — reads by sheet name | None | Only reads specific requested sheet |

### C# — `ExcelAPI/ExcelAPI/Controllers/FormController.cs`

| Line | Function | Filter | Includes ExcelOutputSetting? |
|------|----------|--------|------------------------------|
| **264** | `HasEmbeddedConfiguration()` — OOXML detection | `Contains("ExcelOutputSetting")` | Detection only, not enumeration |

---

## 3. Root Cause Analysis: Where ExcelOutputSetting Leaks

### The problem is in TWO places:

### Location A: Python `_delete_metadata_sheets()` — **The Primary Leak**

```python
# upload_coordinate_generator.py, line 510
def _delete_metadata_sheets(wb):
    """Delete metadata sheets (_Fields, _RawData) from an opened workbook."""
    for i in range(wb.Sheets.Count, 0, -1):
        name = ""
        try:
            name = wb.Sheets(i).Name
        except Exception:
            pass
        if name in ("_Fields", "_RawData"):
            wb.Sheets(i).Delete()
```

**Bug:** Only `_Fields` and `_RawData` are deleted. `ExcelOutputSetting` is NOT deleted.

This function is called in TWO places:
1. **Line 1090** — Export Sanitized PDF: `_delete_metadata_sheets(wb)` runs before `ExportAsFixedFormat`. Since `ExcelOutputSetting` is not deleted, it becomes an extra page in the sanitized PDF.
2. **Line 508** — `export_sanitized_pdf()`: Same issue — `ExcelOutputSetting` survives.

### Location B: Python "Collect Visible Sheets" — **The Secondary Leak**

```python
# upload_coordinate_generator.py, line 1077
for ws in wb.Worksheets:
    try:
        if ws.Visible != -1:  # -1 = xlSheetVisible
            continue
    except Exception:
        pass
    visible_sheets.append(ws)  # ExcelOutputSetting IS added here
    visible_sheet_names.append(ws.Name)
```

**Bug:** The only filter is `ws.Visible != -1`. Since ExcelOutputSetting in a ConMas-generated workbook IS visible (not hidden), it passes this check and is treated as a printable worksheet.

### Why this causes the preview to show 2 pages

```
Workbook:
  Sheet1              ← visible (-1), has fields → page 1 ✅
  ExcelOutputSetting  ← visible (-1), no fields  → page 2 ❌
  _Fields             ← hidden (0 or 2)          → not collected ✅

Result: generate_coordinates_and_preview returns pages[2]:
  pages[0] = Sheet1             (has fields)
  pages[1] = ExcelOutputSetting (empty fields)
```

### Why the C# fix alone wasn't enough

The C# `WorkbookReaderService` fix correctly filters `ExcelOutputSetting` from `contentSheets`. But the C# path is ONLY used when `HasEmbeddedConfiguration()` returns true (reopened path). For new templates, the C# path is never reached — the Python preview runs independently and returns `pages[].` with ExcelOutputSetting included.

**The C# fix prevents config sheets from being designer pages in the reopen path. But the Python preview still sends ExcelOutputSetting as a separate page, which means:**
- Reopened path: 1 sheet in RuntimeForm ✅, BUT Python preview returns 2 pages → background images for ExcelOutputSetting are generated unnecessarily
- New template path: Python returns 2 pages → frontend creates 2 RuntimeForm pages ❌

---

## 4. Where the Filter SHOULD Be Applied

### Recommendation: Filter in THREE places

| Priority | Location | Why |
|----------|----------|-----|
| **1 (CRITICAL)** | Python `_delete_metadata_sheets()` | Add `ExcelOutputSetting` to the delete list. This fixes the sanitized PDF and prevents the extra page from appearing in pixel-scan results. |
| **2 (CRITICAL)** | Python "Collect Visible Sheets" loop (line 1077) | Add a name-based filter alongside the `Visible` check. This prevents `ExcelOutputSetting` from being treated as a printable page at the earliest point. |
| **3 (DEFENSIVE)** | C# `WorkbookReaderService` (ALREADY DONE) | The `ConfigurationSheetNames` filter already prevents config sheets from becoming designer pages. |

### Why two places in Python?

The "Collect Visible Sheets" filter and `_delete_metadata_sheets` serve different purposes:

- **Collect Visible Sheets** determines which sheets become pages in the preview response. Filtering here prevents `ExcelOutputSetting` from appearing in the `pages[]` array.
- **_delete_metadata_sheets** removes sheets before PDF export. Filtering here prevents `ExcelOutputSetting` from being rendered as a PDF page (waste of processing) and from appearing in the sanitized PDF (which could cause pixel-scan misalignment).

### The filter must be name-based, not visibility-based

ExcelOutputSetting in ConMas workbooks is VISIBLE (not hidden). So `ws.Visible != -1` does NOT filter it out. A name-based check is required:

```python
CONFIG_SHEET_NAMES = {"_Fields", "_RawData", "ExcelOutputSetting"}
```

---

## 5. How Legacy ConMas Most Likely Filters

The original ConMas Designer is a C# application installed at:
```
C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer
C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator
```

Based on the reverse engineering evidence:

1. **ConMas enumerates worksheets by iterating `Globals.ThisWorkbook.Worksheets`** but filters out any worksheet that either:
   - Is hidden (Visible = xlSheetHidden or xlSheetVeryHidden)
   - Has a known reserved name (`_Fields`, `ExcelOutputSetting`)

2. **The `_delete_metadata_sheets` function** in our Python code was derived from reverse engineering the ConMas behavior. The original ConMas deletes `_Fields` and `ExcelOutputSetting` before PDF export.

3. **The fact that `ExcelOutputSetting` was never added to `_delete_metadata_sheets`** suggests this was an oversight in our implementation — the original ConMas definitely deletes both.

---

## 6. Summary of Required Changes

| File | Change | Risk |
|------|--------|------|
| `render_service/upload_coordinate_generator.py` line 510 | Add `"ExcelOutputSetting"` to `_delete_metadata_sheets()` delete set | Low — only affects sanitized PDF export |
| `render_service/upload_coordinate_generator.py` line 1077 | Add name-based filter alongside `ws.Visible != -1` to exclude known config sheet names | Low — prevents ExcelOutputSetting from being a page |
| `ExcelAPI/ExcelAPI/Designer/Analysis/WorkbookReaderService.cs` | Already done — `ConfigurationSheetNames` filter in place | ✅ Done |

### Expected result after fixes:

```
Workbook:
  Sheet1              → page 1 ✅ (has fields)
  ExcelOutputSetting  → skipped ❌ (config sheet)
  _Fields             → skipped ✅ (config sheet)

Preview pages = 1 (Sheet1 only)
RuntimeForm sheets = 1 (Sheet1 only)
```
