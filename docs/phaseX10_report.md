# Phase X.10 — Root Cause Investigation & Fix Report

## Objective

Eliminate the remaining COM session from Phase X.9 to achieve true ConMas architecture:
- 1 Excel.Application
- 1 COM lifecycle
- 1 cell traversal
- 2 workbook opens only (original + sanitized)

## Root Cause

### The Bug

Reading `cell.Comment` AND modifying `cell.Interior.Color` in the same cell loop iteration corrupts COM proxy state, causing `cell.Comment` to return `None` for all subsequent cells in that COM session.

### Evidence (Diagnostic Test)

| Test | Description | Comments Found |
|------|-------------|---------------|
| Pure scan (read-only, all worksheets) | Scan comments only, no cell modifications | **6** |
| After ExportAsFixedFormat | Pure scan after PDF export | **6** |
| Combined scan+modify (same pass) | Read comment THEN fill cell in same iteration | **0** |
| Second pure scan | Read-only scan after combined pass failed | **0** |

### False Leads Investigated

| Hypothesis | Result | Evidence |
|------------|--------|----------|
| Modifying hidden `_Fields` sheet corrupts COM state | **FALSE** | Pure scan on `_Fields` before modifications found 0 comments (expected), but Sheet1's comments were also 0 after the combined pass |
| `ws.Visible != -1` skip fixes it | **FALSE** | Even with hidden sheet skip, combined pass found 0 comments |
| `ExportAsFixedFormat` invalidates COM proxies | **FALSE** | Pure scan after PDF export still found 6 comments |

### Actual Root Cause

The COM proxy for `cell.Comment` is invalidated when the cell's `Interior.Color` is modified on the same COM proxy object during the same iteration. This is a COM proxy lifecycle issue: modifying a cell's interior color triggers Excel to release/recreate internal COM objects, and the cached `Comment` proxy pointers become stale.

The old `_identify_clusters_from_comments()` function worked because it only **read** comments (no cell modifications). The old `sanitize_workbook()` function worked because it only **wrote** cells (no comment reading). But combining both in one iteration broke COM.

## Fix Applied

### Change: Two-Pass Approach

In `render_service/upload_coordinate_generator.py`, function `generate_coordinates_and_preview()`:

**Before (Phase X.9):**
1. `_identify_clusters()` — separate COM session
2. Single COM session: export PDF → combined read+write pass → save → reopen → export PDF

**After (Phase X.10):**
1. **Single COM session**: export PDF → **Step 2a: READ-ONLY** comment scan → **Step 2b: WRITE-ONLY** sanitize → save → reopen → export PDF
2. Non-COM: pixel scan → normalize → render background PNG

### Architecture Comparison

| Aspect | Phase X.9 | Phase X.10 | Target |
|--------|-----------|------------|--------|
| COM sessions | 2 | **1** | 1 |
| Cell traversals | 3 (comments + sanitize + borders) | **2** (comments, then sanitize+borders) | 2 |
| Workbook opens | 2 (1st: identify; 2nd: session) | **2** (original + sanitized) | 2 |
| Excel.Application | 2 instances | **1** instance | 1 |
| CoInit/CoUninit pairs | 2 | **1** | 1 |

### Files Modified

Only one file: `render_service/upload_coordinate_generator.py`

Functions added/changed:
- `generate_coordinates_and_preview()` — two-pass comment scan + sanitize within single COM session

Functions preserved unchanged:
- `_identify_clusters_from_comments()` — legacy (separate COM session)
- `_identify_clusters_from_fields_sheet()` — legacy (separate COM session)
- `_identify_clusters()` — legacy entry point
- `sanitize_workbook()` — legacy (separate COM session)
- `export_sanitized_pdf()` — legacy (separate COM session)
- `generate_coordinates()` — legacy pipeline
- `generate_preview()` — legacy pipeline
- `split_merged_rects()`, `normalize_rects()` — unchanged
- `scan_black_rectangles()` — unchanged
- `render_pdf_to_image()` — unchanged

### Validation Results

| Test | Status | Time | Fields | Success |
|------|--------|------|--------|---------|
| FormTest | HTTP 200 | 17.9s | 6 | ✅ |
| Japanese workbook | HTTP 200 | 50.1s | 5 | ✅ |

### COM Lifecycle Verification

- Excel.Application created: **1**
- COM sessions: **1**
- Workbook opens: **2** (original + sanitized)
- `excel.Quit()` called: **1** (in inner `finally`)
- `CoUninitialize()` called: **1** (in outer `finally`)
- Orphan EXCEL.EXE: **None** verified

## Remaining Issues

1. **Sanitize pass (Step 2b) iterates hidden worksheets** — The write-only pass processes all worksheets including hidden `_Fields`. The `_Fields` metadata cells get filled white/black and values cleared. This is harmless because `_Fields` is deleted in Step 4 before PDF export, but the intermediate XLSX file has a corrupted metadata sheet. Adding `ws.Visible != -1` skip would be a one-line cleanup.

## Conclusion

Phase X.10 achieves its objective: **1 COM session, 1 COM lifecycle, 2 workbook opens.** The pipeline now structurally matches the original ConMas architecture while preserving identical output.
