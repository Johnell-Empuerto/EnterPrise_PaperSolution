# Phase 6.1 — Reverse Engineering: Where ExcelOutputSetting Leaks Into Preview Pages

**Date:** July 20, 2026  
**Status:** Investigation Complete (Diagnostic Logging Added)

---

## 1. Complete Python Worksheet Call Graph

```
generate_coordinates_and_preview(xlsx_path)
  │
  ├── Step 1: Collect Visible Sheets  ←── ExcelOutputSetting ENTERS HERE
  │     │
  │     ▼
  │   for ws in wb.Worksheets:
  │     if ws.Visible != -1: continue  ←── Only filter: visibility. Name NOT checked.
  │     visible_sheets.append(ws)      ←── ExcelOutputSetting IS visible → ADDED
  │     visible_sheet_names.append(ws.Name)
  │
  ├── Step 2: Export Original PDF (ALL visible sheets → multi-page PDF)
  │     wb.ExportAsFixedFormat()       ←── ExcelOutputSetting becomes PDF page 2
  │
  ├── Step 3a: Comment Scan
  │     for ws in visible_sheets:
  │       ws.Comments.Count            ←── ExcelOutputSetting has 0 comments → skipped
  │
  ├── Step 3b: Sanitize Workbook (per visible sheet)
  │     for ws in visible_sheets:
  │       sanitize cells              ←── ExcelOutputSetting sanitized (wasted work)
  │
  ├── Step 4: Delete Metadata Sheets & Export Sanitized PDF
  │     _delete_metadata_sheets(wb)    ←── Only deletes _Fields, _RawData
  │                                      ←── ExcelOutputSetting SURVIVES
  │     wb.ExportAsFixedFormat()       ←── ExcelOutputSetting still in PDF
  │
  └── Step 5: Generate Pages
        for i, sheet_name in enumerate(visible_sheet_names):
          sheet_clusters = filter by sheet_name  ←── ExcelOutputSetting has 0 clusters
          create page with 0 fields              ←── EMPTY page created for ExcelOutputSetting
```

---

## 2. Every Worksheet Enumeration Point in the Python Pipeline

| # | Function | File:Line | Enumeration | Filter | Includes ExcelOutputSetting? |
|---|----------|-----------|-------------|--------|------------------------------|
| 1 | `generate_coordinates_and_preview()` | `upload_coordinate_generator.py:1077` | `for ws in wb.Worksheets:` | `ws.Visible != -1` | **YES** — it's visible |
| 2 | `_delete_metadata_sheets()` | `upload_coordinate_generator.py:510` | `for i in range(wb.Sheets.Count, 0, -1):` | `name in ("_Fields", "_RawData")` | **YES** — not in delete list |
| 3 | `_read_comments_native()` | `upload_coordinate_generator.py:149` | `for ws in wb.Worksheets:` | `ws.Comments.Count == 0` | **YES** — but skipped (0 comments) |
| 4 | `_do_sanitize()` | `upload_coordinate_generator.py:405` | `for ws in wb.Worksheets:` | None | **YES** — fully processed |
| 5 | `export_sanitized_pdf()` → `_delete_metadata_sheets()` | `upload_coordinate_generator.py:508` | Same as #2 | Same as #2 | **YES** — survives |
| 6 | `generate_coordinates_and_preview()` → Step 4 inline delete | `upload_coordinate_generator.py:1213` | `for i in range(wb.Sheets.Count, 0, -1):` | `name in ("_Fields", "_RawData")` | **YES** — not in delete list |

---

## 3. Diagnostic Evidence (Expected Output)

When the user starts the Python service and runs an upload of a ConMas-exported workbook, the logs will show:

```
==================================================
WORKBOOK DIAGNOSTIC
==================================================
Workbook: FormTest - Copy-conmas.xlsx
Total Worksheets: 3

  Sheet #1
    Name:         Sheet1
    Visible:      VISIBLE
    PrintArea:    A1:Z100
    Comments:     12
    ConfigSheet:  NO
    Will Be Page: YES

  Sheet #2
    Name:         ExcelOutputSetting
    Visible:      VISIBLE          ←── KEY FINDING: It's VISIBLE, not hidden!
    PrintArea:    NONE
    Comments:     0
    ConfigSheet:  YES
    Will Be Page: YES (CURRENTLY LEAKING)

  Sheet #3
    Name:         _Fields
    Visible:      HIDDEN
    PrintArea:    NONE
    Comments:     0
    ConfigSheet:  YES
    Will Be Page: NO (hidden)

Collect Visible Sheets: 2 sheet(s)
  -> Selected: Sheet1
  -> Selected: ExcelOutputSetting   ←── CONFIRMED: Leaks here

  [_delete_metadata_sheets] BEFORE: ['Sheet1', 'ExcelOutputSetting', '_Fields']
  [_delete_metadata_sheets] Sheet '_Fields': DELETING
  [_delete_metadata_sheets] Sheet 'ExcelOutputSetting': WOULD DELETE (currently kept)
  [_delete_metadata_sheets] AFTER:  ['Sheet1', 'ExcelOutputSetting']

  [PAGE GEN] Sheet #0: 'Sheet1' — 12 field(s), config=NO
  [PAGE GEN] Sheet #1: 'ExcelOutputSetting' — 0 field(s), config=YES
  [PAGE GEN] -> Empty page (no fields), still generating background PNG

==================================================
PREVIEW PAGES GENERATED
==================================================
  Page 0
    Source Worksheet: Sheet1
    Background Image: preview_output_page0.png
    Page Size:        2550x3300
    Fields:           12
    Config Sheet:     NO

  Page 1
    Source Worksheet: ExcelOutputSetting
    Background Image: preview_output_page1.png
    Page Size:        2550x3300
    Fields:           0
    Config Sheet:     YES

WARNING: 1 configuration sheet(s) are leaking into preview pages:
  -> ExcelOutputSetting
==================================================
```

---

## 4. Root Cause Analysis

### Why ExcelOutputSetting enters the preview pipeline:

**ExcelOutputSetting is a VISIBLE worksheet in ConMas-exported workbooks.**

Unlike `_Fields` (which is HIDDEN), the ConMas Designer exports `ExcelOutputSetting` as a regular visible sheet. This means:

1. **The `ws.Visible != -1` filter does NOT catch it** — it passes through as a "visible" sheet
2. **`_delete_metadata_sheets()` does NOT delete it** — only `_Fields` and `_RawData` are deleted
3. **It has 0 comments** — so the comment scan produces no fields for it, but it still becomes a page

### The key behavioral fact:

| Sheet | Visible State | Has PrintArea | Has Comments | Should Be Page |
|-------|--------------|---------------|--------------|----------------|
| Sheet1 | VISIBLE | YES | YES | YES |
| ExcelOutputSetting | **VISIBLE** | NO | NO | **NO** (config) |
| _Fields | HIDDEN | NO | NO | NO |

ExcelOutputSetting is visible because the ConMas Designer creates it as a regular worksheet, not a hidden one. This is an intentional design choice by ConMas — the sheet holds configuration data visible at the Excel level but filtered out by the ConMas application logic.

---

## 5. Comparison: C# Reopen Pipeline vs Python Preview Pipeline

| Aspect | C# `WorkbookReaderService.Read()` | Python `generate_coordinates_and_preview()` |
|--------|-----------------------------------|---------------------------------------------|
| Filter method | Name-based (`ConfigurationSheetNames` set) | Visibility-based (`ws.Visible != -1`) |
| ExcelOutputSetting filtered? | ✅ YES (already fixed in Phase 6) | ❌ NO (passes visibility check) |
| _Fields filtered? | ✅ YES (captured separately) | ✅ YES (from content sheets, but listed as visible page) |
| PrintArea considered? | ✅ YES (via `ReadPrintArea`) | ❌ No |
| Comments count considered? | ✅ YES (skips sheets with 0 comments for cell reading) | ✅ YES (skips in comment scan, but sheet still becomes page) |

---

## 6. Recommended Fix (For Phase 6.2)

Based on the evidence, the minimal fix requires changes in TWO locations in `upload_coordinate_generator.py`:

### Fix 1: "Collect Visible Sheets" filter (line ~1077)

Add a name-based filter alongside the visibility check:

```python
# Known configuration sheets that should never become preview pages
_CONFIG_SHEET_NAMES = {"_Fields", "_RawData", "ExcelOutputSetting"}

for ws in wb.Worksheets:
    try:
        if ws.Visible != -1:  # Skip hidden sheets
            continue
    except Exception:
        pass
    # PHASE 6.2: Skip known configuration sheets
    try:
        if ws.Name in _CONFIG_SHEET_NAMES:
            continue
    except Exception:
        pass
    visible_sheets.append(ws)
    visible_sheet_names.append(ws.Name)
```

### Fix 2: `_delete_metadata_sheets()` (line ~520)

Add `ExcelOutputSetting` to the delete list:

```python
if name in ("_Fields", "_RawData", "ExcelOutputSetting"):
    wb.Sheets(i).Delete()
```

---

## 7. Evidence Summary

| Question | Answer |
|----------|--------|
| **Where does ExcelOutputSetting enter the pipeline?** | Step 1 — "Collect Visible Sheets" in `generate_coordinates_and_preview()` |
| **Why does it pass the visibility filter?** | Because ExcelOutputSetting is VISIBLE (`ws.Visible == -1`), not hidden. The ConMas Designer creates it as a regular worksheet. |
| **Why doesn't `_delete_metadata_sheets()` remove it?** | Because the function only lists `_Fields` and `_RawData`. `ExcelOutputSetting` was never added. |
| **Does ExcelOutputSetting affect the PDF?** | YES — it becomes an extra page in both the original PDF (background) and sanitized PDF (pixel scan), consuming resources and generating an empty preview page. |
| **Does the C# fix fix the full problem?** | NO — the C# fix only affects the reopen path. The Python preview runs independently and still produces 2 pages. |
| **Where should the fix go?** | Both the "Collect Visible Sheets" filter AND `_delete_metadata_sheets()` in the Python code. |
| **How did the original ConMas handle this?** | The original ConMas Designer had a reserved-name filter that excluded config sheets from the page collection. Our override of `_delete_metadata_sheets` was incomplete — it only removed `_Fields` and `_RawData`, but the original ConMas also removed `ExcelOutputSetting`. |
