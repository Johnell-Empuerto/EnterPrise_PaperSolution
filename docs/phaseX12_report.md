# Phase X.12 — Performance Instrumentation Report

## Objective

Add a comprehensive performance profiler to the upload pipeline so every stage can be accurately measured during development, without changing any rendering logic, coordinate generation, COM behavior, or output.

## Changes Made

**File:** `render_service/upload_coordinate_generator.py`

### 1. Performance Flag

Added a single debug flag at the top of the file:

```python
ENABLE_PERFORMANCE_LOGS = True
```

Set to `False` to disable all timing output.

### 2. `_print_perf_report()` Helper

A new function that prints a formatted performance report:

```
==========================================================
PaperLess Upload Performance
==========================================================

  COM: CoInitialize              :       0 ms
  Excel Launch                   :     236 ms
  Workbook Open                  :     401 ms
  Export Original PDF            :     522 ms
  Comment Scan                   :     318 ms
  Workbook Sanitization          :   1,904 ms
  Workbook SaveAs                :      63 ms
  Reopen Sanitized Workbook      :     210 ms
  Export Sanitized PDF           :     471 ms
  Excel Quit                     :       2 ms
  COM: CoUninitialize            :       0 ms
  Render PDF                     :      84 ms
  Pixel Scan                     :     192 ms
  Split Rectangles               :      21 ms
  Normalize Rectangles           :      18 ms
  Render Background PNG          :     104 ms
  Cleanup                        :      17 ms
  ------------------------------------------
  TOTAL                          :   4,421 ms
  TOTAL                          :   4.421 sec
==========================================================

Slow Stage Detected
  Workbook Sanitization          : 1904 ms

Top 1 Slowest Operations
  1. Workbook Sanitization         1904 ms
```

### 3. `_stage_mark()` Calls

Added 17 stage markers at every operation boundary within `generate_coordinates_and_preview()`:

| # | Stage Label | What It Measures |
|---|-------------|-----------------|
| 1 | COM: CoInitialize | `pythoncom.CoInitialize()` |
| 2 | Excel Launch | `win32com.client.Dispatch("Excel.Application")` |
| 3 | Workbook Open | `excel.Workbooks.Open(xlsx_path)` |
| 4 | Export Original PDF | `wb.ExportAsFixedFormat()` for background |
| 5 | Comment Scan | Read-only cell comment traversal (Step 2a) |
| 6 | Workbook Sanitization | Cell fill, shape deletion, border clearing (Step 2b) |
| 7 | Workbook SaveAs | `wb.SaveAs(sanitized_path)` |
| 8 | Reopen Sanitized Workbook | `excel.Workbooks.Open(sanitized_path)` |
| 9 | Export Sanitized PDF | `ws.ExportAsFixedFormat()` for coordinate scan |
| 10 | Excel Quit | `excel.Quit()` |
| 11 | COM: CoUninitialize | `pythoncom.CoUninitialize()` |
| 12 | Render PDF | `render_pdf_to_image()` — fitz PDF → numpy array |
| 13 | Pixel Scan | `scan_black_rectangles()` — NumPy GetAddress |
| 14 | Split Rectangles | `split_merged_rects()` — merge splitting |
| 15 | Normalize Rectangles | `normalize_rects()` — pixel → ratio |
| 16 | Render Background PNG | `pdf_page_to_png()` + `get_page_dimensions()` |
| 17 | Cleanup | `shutil.rmtree()` for temp directories |

### 4. Slow Stage Detection

- **Warning** (>1,000ms): Prints "Slow Stage Detected" with stage name and time
- **Critical** (>5,000ms): Prints "CRITICAL BOTTLENECK" with stage name and time
- **Top-N Ranking**: All slow stages sorted from slowest to fastest

### 5. Report Formatting

- Main report shows stages in **execution order** (not time-sorted)
- Top-N ranking section is **time-sorted** (slowest first)
- Print uses `time.perf_counter()` for high-resolution timing
- All durations printed in milliseconds

## Stages NOT Modified

- `scan_black_rectangles()` — untouched
- `normalize_rects()` — untouched  
- `split_merged_rects()` — untouched
- `generate_coordinates()` — untouched
- `generate_preview()` — untouched
- All XML/JSON output — untouched
- All COM workflow — untouched

## Validation

Both workbooks pass with identical output:

| Workbook | Status | Fields | HTTP |
|----------|--------|--------|------|
| FormTest | Success | 6 | 200 |
| Japanese | Success | 5 | 200 |

## Usage

The performance report prints to **stdout** (the uvicorn server console), not the HTTP response. To see it:

1. Start the server: `python -m uvicorn render_service.app:app --host 127.0.0.1 --port 5091`
2. Upload a workbook via `/upload/preview`
3. Check the server console for the formatted report
