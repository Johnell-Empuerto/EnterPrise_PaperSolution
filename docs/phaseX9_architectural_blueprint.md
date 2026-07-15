# Phase X.9 — Architectural Blueprint: Rebuild Upload Pipeline to Match Original ConMas

> **Status:** Design only — no code changes until approved  
> **Date:** 2026-07-15  
> **Objective:** Reproduce the original ConMas upload pipeline architecture as closely as possible  

---

## Part 1 — Reverse-Engineered ConMas Upload Flow (Complete)

### Sequence Diagram

```
User clicks "Publish" in Excel Add-In
         │
         ▼
┌──────────────────────────────────────────────────────────────────┐
│ 1. ExportExcelDefinition(xlsPath, outputDir)                      │
│    DLL: iReporterExcelAddIn → LibExcelController                  │
│    Backend: Infragistics4.Documents.Excel v23.1 (PRIMARY)         │
│             Microsoft.Office.Interop.Excel (SECONDARY, for PDF)   │
└──────────────────────────────┬───────────────────────────────────┘
                               │
         ┌─────────────────────┼─────────────────────┐
         │                     │                     │
         ▼                     ▼                     ▼
┌─────────────────┐  ┌──────────────────┐  ┌──────────────────┐
│ MakeCluster()    │  │ workbook.SaveAs  │  │ CalcClusterSize() │
│ PER WORKSHEET   │  │ (sanitized copy) │  │ (coordinate gen)  │
│                 │  │                  │  │                  │
│ SAME COM session│  │ SAME COM session │  │ SAME COM session  │
└────────┬────────┘  └──────────────────┘  └────────┬─────────┘
         │                                          │
         ▼                                          ▼
┌──────────────────────────┐  ┌──────────────────────────────────┐
│ MakeCluster() Algorithm:  │  │ CalcClusterSize() Algorithm:     │
│                           │  │                                 │
│ 1. ClearShapes()          │  │ 1. ExportPdf(temp.xlsx,         │
│    (ws.Shapes.Clear())    │  │      temp.pdf)                  │
│                           │  │    via Excel COM                │
│ 2. Scan UsedRange cells   │  │    (same session, reopen wb)    │
│    via Infragistics       │  │                                 │
│    (NOT COM per-cell)     │  │ 2. CreateImageFromPdf(          │
│                           │  │      pdfPath, 0, false)         │
│ 3. For each cell:         │  │    → System.Drawing.Image       │
│    IF cell.HasComment():  │  │    → 300 DPI                    │
│      fill MergeArea BLACK │  │    → imageWidth, imageHeight    │
│      clear cell value     │  │                                 │
│    ELSE:                  │  │ 3. GetClusterSize(pdfPath, 0,   │
│      fill cell WHITE      │  │      clusterCount)              │
│      clear cell value     │  │    → GetAddress(img, count)     │
│                           │  │    → Pixel scan → ClusterRect[] │
│ 4. Clear ALL borders      │  │                                 │
│    (Borders.LineStyle=    │  │ 4. SortClusters(rects, topThresh│
│     xlNone)               │  │      leftThresh)                │
│                           │  │                                 │
│ 5. Clear headers/footers  │  │ 5. For each cluster in XML:     │
│    (CenterHeader="",      │  │      ratio = pixel / dimension  │
│     CenterFooter="")      │  │      → 7 decimal places         │
│                           │  │      → SetAttributeValue()      │
│ 6. Build XElement          │  │                                 │
│    with cluster metadata  │  │                                 │
│    (cellAddr, name, type)  │  │                                 │
└──────────────────────────┘  └──────────────────────────────────┘
                               │
                               ▼
┌──────────────────────────────────────────────────────────────────┐
│ After all sheets processed:                                      │
│                                                                  │
│ 1. ExportPdf(clean-original.xlsx, output.pdf)                    │
│    → Background image PDF (NEW COM session or reuse)              │
│                                                                  │
│ 2. Single database transaction:                                  │
│    • INSERT def_top (form metadata)                               │
│    • INSERT def_sheet(s) (sheet+page dimensions)                  │
│    • INSERT def_cluster (ratios + cellAddr + metadata)            │
│    • UPDATE def_top SET xml_data (full XML definition)            │
│    • UPDATE def_top SET background_image_file (PDF bytes)         │
└──────────────────────────────────────────────────────────────────┘
```

### Library Usage per Stage

| Stage | Library Used | COM? | Workbook Reopened? | Excel Instance Reused? |
|-------|-------------|------|-------------------|----------------------|
| Open workbook | Infragistics4.Documents.Excel | **NO** | N/A | N/A (first open) |
| Scan comments (MakeCluster) | Infragistics API (`get_HasComment`, `get_Comment`) | **NO** | No | Yes (same instance) |
| Fill cells, clear shapes | Microsoft.Office.Interop.Excel | **YES** | No | Yes (same instance) |
| SaveAs sanitized copy | Microsoft.Office.Interop.Excel | **YES** | No | Yes (same instance) |
| Export sanitized → PDF | `Workbook.ExportAsFixedFormat()` | **YES** | **Yes** (reopen sanitized) | **Yes** (same `Excel.Application`) |
| Render PDF → bitmap | `System.Drawing.Graphics` via PDF renderer | **NO** | No | N/A |
| Pixel scan (GetAddress) | `Bitmap.GetPixel()` | **NO** | No | N/A |
| Normalize ratios | C# float arithmetic | **NO** | No | N/A |
| Export original → PDF (background) | `Workbook.ExportAsFixedFormat()` | **YES** | **Yes** (reopen original) | **Maybe** new instance |
| Database writes | ADO.NET / PostgreSQL | **NO** | No | N/A |

**Key insight:** ConMas uses **Infragistics (non-COM) for metadata reading** and **COM only for cell manipulation + PDF export**. The COM part happens within 1-2 workbook opens in a **single `Excel.Application` session**.

---

## Part 2 — Side-by-Side Comparison: ConMas vs Current

### Full Stage Comparison

| Stage | ConMas | Current System | Same? | Should Change? | WHY |
|-------|--------|---------------|-------|---------------|-----|
| **Workbook open** | 1x (.NET) | 4x (Python COM) | ❌ | **YES** | 4 separate `win32com.client.Dispatch` calls each create a new Excel process. ConMas reuses ONE `Excel.Application` for MakeCluster + SaveAs + ExportPdf. |
| **Metadata (comment) reading** | Infragistics API (non-COM) — reads XLSX directly via `WorksheetCell.get_HasComment()` | `win32com` COM — `Range.Comment` per cell | ❌ | **YES** | Infragistics is ~10-100× faster per cell because it reads the ZIP/XLSX directly without COM marshaling overhead. Each COM call involves inter-process marshaling to the Excel process. |
| **Comment text parsing** | First-line parsing: `lines[0].Trim()` = field type | Two-format parser (detects old/new format) | ⚠️ | **MINOR** | Current parser handles both formats (name-first and type-first) which is actually MORE robust than ConMas. Keep as-is. |
| **Sanitization pass** | Single pass: identify clusters + fill black/white + clear borders + clear shapes + clear headers | Single pass: same operations | ✅ | **NO** | Same algorithm, same operations. Only difference is the COM overhead (ConMas does it in existing session, current opens new session). |
| **Save sanitized workbook** | `workbook.SaveAs(temp.xlsx)` | `wb.SaveAs(path)` | ✅ | **NO** | Identical operation. |
| **Export sanitized PDF** | `ExportAsFixedFormat(xlTypePDF, path)` via COM | `ExportAsFixedFormat(0, path, 0, 0, True)` via COM | ⚠️ | **YES** | ConMas passes `IgnorePrintAreas=false` (default). Current passes `True`. This is a confirmed behavioral difference that affects workbooks with explicit PrintArea. |
| **Render PDF to image** | `CreateImageFromPdf()` at 300 DPI via System.Drawing | `render_pdf_to_image()` at 300 DPI via PyMuPDF | ✅ | **NO** | Both render page 0 at 300 DPI. No morphological close in either path (both pass `applyMorphology=false`). |
| **Pixel scan (GetAddress)** | `Bitmap.GetPixel(x, y)` — top→bottom, left→right. Detect BLACK pixel with WHITE left/above. 6-pixel noise filter. Expand right/down. Min 6×6 filter. | `scan_black_rectangles()` — NumPy array. Same algorithm, same constants. | ✅ | **NO** | Algorithm is functionally identical. The NumPy implementation is actually FASTER than `Bitmap.GetPixel` (vectorized vs per-pixel COM marshaling). |
| **Split merged rects** | ❌ NOT needed — ConMas's GetAddress naturally separates adjacent clusters | `split_merged_rects()` — workaround for merged blobs | ❌ | **YES** | This is a workaround for a bug in our pixel scan. ConMas's GetAddress uses a different scanning order or edge detection that naturally separates touching rectangles. We should investigate why our scan differs and fix the root cause rather than patching with split_merged_rects. |
| **Ratio normalization** | `ratio = pixel / imageDimension`, 7 decimals | `ratio = pixel / imageDimension`, 7 decimals | ✅ | **NO** | Identical formula. |
| **Export original PDF (background)** | `ExportPdf(clean.xlsx, output.pdf)` | `xlsx_to_pdf(xlsx_path)` | ✅ | **LOCATION** | Same operation. But ConMas does this in a separate COM session OR reuses the same one. Should be moved INTO the single session. |
| **Render background PNG** | Stored as PDF bytes in DB, rendered by client | `pdf_page_to_png()` via PyMuPDF | ✅ | **NO** | Both use page 0 at 300 DPI. The rendering is equivalent. |
| **COM cleanup** | `KillExcelProcess()` + COM release in `ComRelease.FinalReleaseComObjects()` | `excel.Quit()` + `pythoncom.CoUninitialize()` | ⚠️ | **MINOR** | Both clean up. Current system now has CoInitialize/CoUninitialize (fixed in Phase X.7). |
| **JSON serialization** | ❌ (returns XML XElement) | ✅ (returns JSON for frontend) | ❌ | **KEEP** | This is a difference in output format, not pipeline architecture. The JSON is consumed by the Next.js frontend. ConMas outputs XML to the database. Both are valid. |
| **Database storage** | PostgreSQL: `def_cluster`, `def_sheet`, `def_top` | Preview: JSON returned to frontend. Publish: via C# backend | ❌ | **KEEP** | Preview endpoint never stores to DB. This is correct behavior for a PREVIEW-only endpoint. The database write happens in the C# Publish pipeline, not in Python. |

### Summary: What Must Change

| Change | Priority | Effort | Impact |
|--------|----------|--------|--------|
| **Merge 4 COM sessions → 1** | 🔴 HIGH | Medium | Eliminates ~3× COM overhead (~1.5s saved). Primary architectural goal. |
| **Export original PDF BEFORE sanitization** | 🔴 HIGH | Small | Current order (coords first, then background) requires separate COM session. Reversing order enables single session. |
| **Fix `IgnorePrintAreas=True` → `False`** | 🟡 MEDIUM | Trivial | Only affects workbooks with explicit PrintArea. Matches ConMas behavior. |
| **Investigate GetAddress pixel scan divergence** | 🟡 MEDIUM | Unknown | Root cause of `split_merged_rects` workaround. If fixed, eliminates that extra step. |
| **Export both PDFs in same session** | 🟢 LOW | Small | Reduces 2 separate COM sessions to 1 session with 2 workbook opens. |
| **Keep JSON output** | ✅ NONE | N/A | Preview needs JSON. This is correct. |

---

## Part 3 — Identified Duplicated Operations

### All Duplicated COM Operations

| Operation | Occurrences | Current Count | ConMas Count | Overhead |
|-----------|------------|---------------|--------------|----------|
| **Excel.Application creation** | `_identify_clusters()`, `sanitize_workbook()`, `export_sanitized_pdf()`, `xlsx_to_pdf()` | **4** | **1** | Each Excel startup: ~200-500ms. Total saved: **~600ms-2s** |
| **Workbook.Open** | Same 4 functions | **4** | **2** (original + sanitized) | Each open: ~200-500ms. Total saved: **~400ms-1s** |
| **Per-cell COM iteration** | `_identify_clusters()` + `sanitize_workbook()` | **2 full passes** | **1 full pass** (combined) | Japanese workbook: 1935 cells × 2 = 3870 COM calls vs ConMas's 1935. Saved: **~10-15s** |
| **`_Fields` sheet scan** | `_identify_clusters_from_fields_sheet()` | **1** | **0** (not in MakeCluster) | ~500ms. Remove since ConMas doesn't do it in MakeCluster. |
| **Save sanitized workbook** | `sanitize_workbook()` | **1** | **1** | Same. No duplication. |
| **Export PDF (sanitized)** | `export_sanitized_pdf()` | **1** | **1** | Same. No duplication. |
| **Export PDF (original)** | `xlsx_to_pdf()` | **1** | **1** | Same. No duplication. |
| **Render PDF to pixel array** | `render_pdf_to_image()` + `pdf_page_to_png()` | **2** | **2** (sanitized for coords + background) | Same count. Both needed. |
| **COM cleanup** | `excel.Quit()` | **4** | **1** | _Per session. Saved: ~negligible._ |

### Time/Memory Estimate for Removing Duplications

| Optimization | Time Saved | Memory Saved | COM Overhead Reduced |
|-------------|-----------|-------------|---------------------|
| 4 Excel sessions → 1 | ~1.5s | ~200 MB (4 Excel processes → 1) | 75% reduction |
| Combined comment scan + sanitize → 1 pass | ~5-15s | ~50 MB | 50% fewer COM calls |
| Remove `_Fields` sheet scan from coordinate path | ~0.5s | ~10 MB | Not applicable |
| Total | **~7-17s saved** | **~260 MB** | **75% fewer Excel processes** |

**Current total for Japanese workbook: ~84s. Target total after optimizations: ~67-77s.**

Note: The dominant bottleneck remains `scan_black_rectangles` at ~40s (48% of total). This is NOT duplicated — it runs once. ConMas has the same bottleneck. The 40s is inherent to 2550×3300 pixel scanning.

---

## Part 4 — Designed Upload Pipeline (Single COM Session)

### Architecture Diagram

```
┌────────────────────────────────────────────────────────────────┐
│                     SINGLE EXCEL SESSION                        │
│                                                                │
│  Excel.Application (1 process)                                 │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │ Workbook.Open(original.xlsx)                              │  │
│  │                                                          │  │
│  │  STEP 1: Export original PDF (background)                 │  │
│  │  wb.ExportAsFixedFormat(0, original.pdf)                  │  │
│  │  ← MUST be BEFORE sanitization, since sanitization       │  │
│  │    destroys cell content                                  │  │
│  │                                                          │  │
│  │  STEP 2: Scan comments + sanitize in ONE pass            │  │
│  │  FOR each cell in UsedRange:                              │  │
│  │    IF cell has comment:                                   │  │
│  │      → Record metadata (addr, name, type)                 │  │
│  │      → Fill MergeArea BLACK, clear value                  │  │
│  │    ELSE:                                                  │  │
│  │      → Fill WHITE, clear value                            │  │
│  │  Clear borders, delete shapes, clear headers/footers      │  │
│  │                                                          │  │
│  │  STEP 3: Save sanitized copy                              │  │
│  │  wb.SaveAs(sanitized.xlsx)                                │  │
│  │  wb.Close(False)                                          │  │
│  │                                                          │  │
│  │  STEP 4: Export sanitized PDF (reopen sanitized copy)     │  │
│  │  wb2 = excel.Workbooks.Open(sanitized.xlsx)               │  │
│  │  Delete metadata sheets (_Fields, _RawData)               │  │
│  │  wb2.ExportAsFixedFormat(0, sanitized.pdf)                │  │
│  │  wb2.Close(False)                                         │  │
│  │                                                          │  │
│  │  STEP 5: Cleanup                                          │  │
│  │  excel.Quit()                                             │  │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────┘
         │                               │
         │ sanitized.pdf                 │ original.pdf
         ▼                               ▼
┌──────────────────────┐    ┌──────────────────────┐
│ NO COM                │    │ NO COM                │
│                       │    │                       │
│ render_pdf_to_image() │    │ pdf_page_to_png()     │
│ → 300 DPI PNG         │    │ → preview background  │
│                       │    │                       │
│ scan_black_rectangles()│   │ get_page_dimensions()  │
│ → pixel rects         │    │ → page width/height   │
│                       │    │                       │
│ normalize_rects()     │    │                       │
│ → 0-to-1 ratios      │    │                       │
│                       │    │                       │
│ Merge with metadata   │    │                       │
│ → field definitions   │    │                       │
└──────────────────────┘    └──────────────────────┘
         │                               │
         └───────────┬───────────────────┘
                     ▼
            ┌────────────────┐
            │ Return JSON     │
            │ backgroundImage │
            │ page dimensions │
            │ fields (ratios) │
            └────────────────┘
```

### Key Design Principles

1. **Single `Excel.Application`** — ONE COM process instead of four
2. **Original PDF exported BEFORE sanitization** — because sanitization destroys cell content
3. **Combined comment scan + sanitization** — one pass over cells instead of two
4. **Sanitized workbook reopened** — `SaveAs` changes the workbook reference, so reopen the saved copy for PDF export
5. **All non-COM operations after COM cleanup** — pixel scan, rendering, normalization happen AFTER Excel quits
6. **`IgnorePrintAreas=False`** — matches ConMas default behavior (parameter removed from call)

### Pseudocode

```python
def generate_coordinates_and_preview(xlsx_path: str, output_dir: str, output_id: str) -> dict:
    """
    Single COM session: open workbook ONCE, export both PDFs, sanitize.
    Then non-COM: pixel scan, normalize, render background.
    """
    import pythoncom
    import win32com.client
    import tempfile, os, shutil
    from pathlib import Path
    
    os.makedirs(output_dir, exist_ok=True)
    
    # ── Temporary directories (cleaned up at end) ──
    tmp_dirs = []
    
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        
        # ════════════════════════════════════════════
        # STEP 1: Export original PDF (background)
        # MUST be before sanitization (content is preserved in original)
        # ════════════════════════════════════════════
        orig_pdf_dir = tempfile.mkdtemp(prefix="ple_orig_pdf_")
        tmp_dirs.append(orig_pdf_dir)
        orig_pdf_path = os.path.join(orig_pdf_dir, "original.pdf")
        wb.ExportAsFixedFormat(0, os.path.abspath(orig_pdf_path))
        
        # ════════════════════════════════════════════
        # STEP 2: Scan comments + sanitize in ONE pass
        # ════════════════════════════════════════════
        cluster_meta = []
        cluster_addrs = set()
        xlNone = -4142
        
        for ws in wb.Worksheets:
            # Clear shapes (interference prevention)
            try:
                for shape in list(ws.Shapes):
                    shape.Delete()
            except Exception:
                pass
            
            # Clear headers/footers
            ws.PageSetup.CenterHeader = ""
            ws.PageSetup.CenterFooter = ""
            
            used = ws.UsedRange
            if used is None:
                continue
            
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            
            for row in range(1, lr + 1):
                for col in range(1, lc + 1):
                    cell = ws.Cells(row, col)
                    
                    try:
                        comment = cell.Comment
                    except Exception:
                        comment = None
                    
                    if comment is not None:
                        # ── Cluster cell: record metadata + fill black ──
                        try:
                            ma = cell.MergeArea
                        except Exception:
                            ma = cell
                        
                        try:
                            addr = str(ma.Address).upper()
                        except Exception:
                            addr = str(cell.Address).upper()
                        
                        try:
                            text = str(comment.Text())
                        except Exception:
                            text = ""
                        
                        lines = text.replace("\r\n", "\n").split("\n")
                        cluster_meta.append({
                            "cellAddr": addr,
                            "name": lines[0].strip() if lines else "",
                            "type": lines[1].strip() if len(lines) > 1 else "Text",
                            "input_parameter": "\n".join(lines[2:]).strip() if len(lines) > 2 else "",
                        })
                        cluster_addrs.add(addr)
                        
                        # Fill BLACK
                        try:
                            ma.Interior.Color = 1  # Black
                            ma.Value = ""
                        except Exception:
                            pass
                    else:
                        # ── Non-cluster: fill white ──
                        try:
                            cell.Interior.Color = 0xFFFFFF  # White
                            cell.Value = ""
                        except Exception:
                            pass
            
            # Clear borders
            try:
                clear_range = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc))
                clear_range.Borders.LineStyle = xlNone
            except Exception:
                pass
        
        if not cluster_meta:
            wb.Close(False)
            excel.Quit()
            return {"success": False, "message": "No clusters found"}
        
        # ════════════════════════════════════════════
        # STEP 3: Save sanitized copy
        # ════════════════════════════════════════════
        sanitized_dir = tempfile.mkdtemp(prefix="ple_san_")
        tmp_dirs.append(sanitized_dir)
        sanitized_path = os.path.join(sanitized_dir, "sanitized.xlsx")
        wb.SaveAs(os.path.abspath(sanitized_path))
        wb.Close(False)
        
        # ════════════════════════════════════════════
        # STEP 4: Export sanitized PDF (reopen sanitized copy)
        # ════════════════════════════════════════════
        wb_san = excel.Workbooks.Open(os.path.abspath(sanitized_path))
        
        # Delete metadata sheets to avoid multi-page PDF
        for i in range(wb_san.Sheets.Count, 0, -1):
            try:
                name = wb_san.Sheets(i).Name
            except Exception:
                name = ""
            if name in ("_Fields", "_RawData"):
                wb_san.Sheets(i).Delete()
        
        pdf_dir = tempfile.mkdtemp(prefix="ple_pdf_")
        tmp_dirs.append(pdf_dir)
        pdf_path = os.path.join(pdf_dir, "sanitized.pdf")
        
        # Export with IgnorePrintAreas=False (matching ConMas default)
        ws_active = wb_san.ActiveSheet
        ws_active.ExportAsFixedFormat(0, os.path.abspath(pdf_path),
                                       0, 0, False)  # IgnorePrintAreas=False
        wb_san.Close(False)
        
        # ════════════════════════════════════════════
        # STEP 5: Cleanup COM
        # ════════════════════════════════════════════
        excel.Quit()
    finally:
        pythoncom.CoUninitialize()
    
    # ════════════════════════════════════════════
    # Non-COM phases: pixel scan, normalize, render
    # ════════════════════════════════════════════
    try:
        # Phase: Pixel scan (NO COM)
        img, img_w, img_h = render_pdf_to_image(pdf_path)
        rects = scan_black_rectangles(img)
        # Try without split_merged_rects first (investigate root cause)
        if len(rects) < len(cluster_meta):
            rects = split_merged_rects(rects, cluster_meta)
        rects = normalize_rects(rects, img_w, img_h)
        
        # Sort for matching
        cluster_meta.sort(key=_sort_key_meta)
        rects.sort(key=lambda r: (r["Top"], r["Left"]))
        
        # Merge metadata with ratios
        fields = []
        for meta, rect in zip(cluster_meta, rects):
            fields.append({
                "name": meta["name"],
                "type": meta["type"],
                "cellAddr": meta["cellAddr"],
                "input_parameter": meta.get("input_parameter", ""),
                "left_ratio": round(rect["left_ratio"], 7),
                "top_ratio": round(rect["top_ratio"], 7),
                "right_ratio": round(rect["right_ratio"], 7),
                "bottom_ratio": round(rect["bottom_ratio"], 7),
            })
        
        # Phase: Render background PNG (NO COM)
        png_filename = f"preview_{output_id}.png"
        png_path = os.path.join(output_dir, png_filename)
        pdf_page_to_png(orig_pdf_path, png_path, dpi=300)
        page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=300)
        
        return {
            "backgroundImage": png_filename,
            "page": {"width": page_w, "height": page_h},
            "fields": fields,
        }
    finally:
        # Cleanup temporary files
        for d in tmp_dirs:
            shutil.rmtree(d, ignore_errors=True)
```

---

## Part 5 — Pixel Geometry Preservation

### PageSetup Properties: Before vs After Redesign

The redesign MUST preserve all page setup properties during sanitized PDF export. Here's the verification:

| Property | Phase X.2 Fix (current) | Redesign (proposed) | Preserved? | How |
|----------|------------------------|---------------------|------------|-----|
| `Zoom` | Preserved (no longer modified) | Preserved | ✅ | Never modified |
| `FitToPagesWide` | Preserved | Preserved | ✅ | Never modified |
| `FitToPagesTall` | Preserved | Preserved | ✅ | Never modified |
| `PaperSize` | Preserved | Preserved | ✅ | Never modified |
| `Orientation` | Preserved | Preserved | ✅ | Never modified |
| `LeftMargin` | Preserved | Preserved | ✅ | Never modified |
| `RightMargin` | Preserved | Preserved | ✅ | Never modified |
| `TopMargin` | Preserved | Preserved | ✅ | Never modified |
| `BottomMargin` | Preserved | Preserved | ✅ | Never modified |
| `CenterHorizontally` | Preserved | Preserved | ✅ | Never modified |
| `CenterVertically` | Preserved | Preserved | ✅ | Never modified |
| `PrintArea` | Preserved | Preserved | ✅ | Never modified |
| `PrintTitleRows` | Preserved | Preserved | ✅ | Never modified |
| `PrintTitleColumns` | Preserved | Preserved | ✅ | Never modified |
| `CenterHeader` | Cleared | Cleared | ✅ | Same as ConMas MakeCluster |
| `CenterFooter` | Cleared | Cleared | ✅ | Same as ConMas MakeCluster |
| `Borders` | Cleared (xlNone) | Cleared (xlNone) | ✅ | Same as ConMas MakeCluster |
| Shapes | Deleted | Deleted | ✅ | Same as ConMas MakeCluster |
| Cell interior colors | Changed (black/white) | Changed (black/white) | ✅ | Same as ConMas MakeCluster |
| Cell values | Cleared | Cleared | ✅ | Same as ConMas MakeCluster |

**Critical verification**: The sanitized PDF must remain geometrically identical to the original workbook PDF. In the redesign:
- Original PDF is exported FIRST (before any modifications) → guaranteed identical to workbook
- Sanitized PDF is exported from the modified workbook → geometry unchanged because only INTERIOR colors and cell VALUES changed (not column widths, row heights, page setup, or cell formatting that affects layout)

### Shape Handling

Shapes are DELETED in the sanitized workbook. This is correct behavior matching ConMas's `MakeCluster()`. If a shape's presence affects the layout (e.g., a shape that pushes content down), the sanitized PDF may have different content positions than the original PDF.

**This is the Phase X.3 finding**: shape deletion can change the PDF geometry if shapes are anchored to cells and affect their size. The ConMas pipeline has the SAME behavior (it also deletes shapes), so this is not a regression — it's inherent to the MakeCluster approach.

If shape-caused geometry changes are unacceptable, the fix would be to:
1. Keep shapes but make them transparent/invisible instead of deleting them
2. This would preserve geometry but ComMas didn't do this either

**Recommendation**: Accept this behavior as equivalent to ConMas. If shape alignment issues arise, investigate as a separate issue.

---

## Part 6 — Pixel Scan vs COM Range: Which Should Be Production?

### Head-to-Head Comparison

| Criteria | Pixel Scan (current after Phase X.7) | COM Range (C# `MeasureFieldsFromCom`) | Winner |
|----------|--------------------------------------|--------------------------------------|--------|
| **Coordinate accuracy** | ✅ Correct — measures actual rendered output | ❌ Incorrect for FitToPages workbooks — formula doesn't account for scaling | **Pixel Scan** |
| **FitToPages support** | ✅ Handles any Zoom/FitToPages/FitToWidth/FitToHeight automatically (PDF is rendered at workbook's actual settings) | ❌ Requires additional `CellToPagePt()` scaling transformation that is NOT currently applied | **Pixel Scan** |
| **Margins** | ✅ Included in PDF render — no computation needed | ⚠️ Requires manual calculation of origin + margins | **Pixel Scan** |
| **Centering** | ✅ Included in PDF render — no computation needed | ⚠️ Requires manual calculation | **Pixel Scan** |
| **Merged cells** | ✅ Handled naturally (merge area is filled as one black blob) | ⚠️ Requires reading MergeArea | **Pixel Scan** |
| **Performance** | ⚠️ Slowest stage: ~40s for 2550×3300 image (48% of total time) | ✅ Instant — COM Range property reads are fast | **COM Range** |
| **COM dependency** | ⚠️ Requires COM for PDF export (but pixel scan itself is NumPy) | ❌ Requires COM for ALL measurements | **Pixel Scan** |
| **Maintenance** | ✅ Simple algorithm — pixel scan hasn't needed changes | ⚠️ Complex — many edge cases (FitToPages, margins, centering) | **Pixel Scan** |
| **Similarity to ConMas** | ✅ Pixel scan is what ConMas used (GetAddress) | ❌ ConMas did NOT use Range.Left/Top for coordinates | **Pixel Scan** |
| **Database compatibility** | ✅ Produces same 7-decimal ratios as ConMas | ❌ Produces different values (needs calibration) | **Pixel Scan** |
| **Implementation complexity** | ~200 lines (scan + normalize) | ~300 lines (MeasureFieldsFromCom + CoordinateTransformer) | **Pixel Scan** |

### Evidence Summary

**From Phase X.8 validation:**
- Pixel scan coordinates match ConMas < 2px error (Phase X.8, for workbooks w/o shape conflicts)
- COM Range produces DIFFERENT coordinates for FitToPages workbooks
- The `CellToPagePt()` method exists in the codebase but is NOT called by `MeasureFieldsFromCom`

**From Phase X.3 investigation:**
- When sanitized PDF = original PDF (same page setup), pixel scan coordinates align perfectly
- Coordinate divergence was caused by FitToPages settings being stripped, not by the pixel scan itself

**From CalcClusterSize_ReverseEngineering.md:**
- ConMas used `GetAddress()` = pixel scan. This is proven by IL decompilation.
- ConMas NEVER used `Range.Left`/`Range.Top`/`Range.Width`/`Range.Height` for coordinates.
- The only COM range operations in MakeCluster are for: identifying cells (HasComment), filling black/white, reading metadata

### Conclusion

**The pixel scan pipeline is the canonical source of truth** and should remain the production implementation. The COM Range pipeline is:
- Not what ConMas used
- Missing the scaling transformation for FitToPages
- More complex to maintain (many edge cases)
- Not compatible with existing database ratios

The only advantage of COM Range is speed (~40s saved). But this is irrelevant if the coordinates are wrong.

### Recommendation

**Phase out the COM Range pipeline** (`MeasureFieldsFromCom`, `CoordinateTransformer.CellToPagePt`, and related code) and invest optimization effort in:
1. Making the pixel scan faster (if needed — 40s for a preview is acceptable)
2. Single COM session (Phase X.9 redesign)
3. Fixing the GetAddress divergence (`split_merged_rects` workaround)

---

## Part 7 — Migration Plan

### Files to Modify

| File | Change | Priority |
|------|--------|----------|
| `render_service/upload_coordinate_generator.py` | **Merge `generate_coordinates()` + `generate_preview()` into single-session pipeline.** Replace 4 separate COM functions with one `generate_coordinates_and_preview()` that opens the workbook once, exports both PDFs, scans comments, sanitizes, then does pixel scan + normalization. | 🔴 P0 |
| `render_service/upload_coordinate_generator.py` | **Remove `_identify_clusters_from_comments()` + `_identify_clusters_from_fields_sheet()` as standalone functions.** Their logic moves into the combined pipeline. Keep `_identify_clusters()` as an internal call only if needed for the `/upload/coordinates` endpoint. | 🟡 P1 |
| `render_service/upload_coordinate_generator.py` | **Remove `sanitize_workbook()` as standalone function.** Logic moves into combined pipeline. | 🟡 P1 |
| `render_service/upload_coordinate_generator.py` | **Remove `export_sanitized_pdf()` as standalone function.** Logic moves into combined pipeline. | 🟡 P1 |
| `render_service/upload_coordinate_generator.py` | **Change `IgnorePrintAreas=True` → `False` (or remove the parameter).** Matches ConMas default behavior. | 🟡 P1 |
| `render_service/app.py` | **Update `upload_preview()` endpoint** to call the new `generate_coordinates_and_preview()` instead of `generate_preview()`. | 🔴 P0 |
| `render_service/app.py` | **Update `upload_coordinates()` endpoint** — may still need access to old `generate_coordinates()`. Decision: keep the old function path for this endpoint, or add a new parameter to skip background PNG generation. | 🟡 P1 |

### Files to Merge

| Source Files | Target File | Contents to Merge |
|-------------|------------|-------------------|
| `upload_coordinate_generator.py` (existing) | `upload_coordinate_generator.py` (refactored) | All functions merged into single `generate_coordinates_and_preview()`. The old `generate_coordinates()`, `generate_preview()`, `_identify_clusters()`, `sanitize_workbook()`, `export_sanitized_pdf()` become internal helpers or are inlined. |
| `pdf_converter.py` (existing) | `upload_coordinate_generator.py` (refactored) | `xlsx_to_pdf()` logic (ExportAsFixedFormat) moves into the combined pipeline. The `xlsx_to_pdf()` function can remain as a standalone utility for other callers (e.g., `renderer.py`). |

### Files to Remove (or Deprecate)

| File/Function | Reason | Replacement |
|--------------|--------|-------------|
| `_identify_clusters_from_fields_sheet()` standalone path in coordinate pipeline | ConMas MakeCluster does NOT read `_Fields` sheet. This path is only for the C# `GetDefinition()` XML export, not for coordinate generation. | Inline the fallback logic into the combined pipeline if needed. |
| `split_merged_rects()` | Workaround for pixel scan bug. Should be removed once the GetAddress divergence root cause is found and fixed. | Corrected pixel scan that naturally separates adjacent clusters. |
| `CoordinateTransformer.CellToPagePt()` | Never called in production. The COM Range pipeline is being phased out. | Deprecate with a comment. |

### Files to Keep Unchanged

| File | Reason |
|------|--------|
| `render_service/background_renderer.py` | `pdf_page_to_png()` and `get_page_dimensions()` are correct and unchanged. |
| `render_service/excel_cluster_reader.py` | Used by the C# runtime pipeline, not by the Python coordinate pipeline. Its `read_fields()` function is a separate code path. |
| `render_service/renderer.py` | The runtime rendering engine is independent of the upload pipeline. |
| `render_service/page_coordinate_transformer.py` | Used by `excel_cluster_reader.py` and legacy coordinate pipeline. Not part of pixel scan pipeline. |
| `render_service/xml_field_provider.py` | Reads from database. Not part of upload pipeline. |
| `render_service/models.py` | Pydantic models — architectural independent. |
| `render_service/pdf_converter.py` (standalone utility) | Keep as a utility for other callers (e.g., `renderer.py` calls it for rendering). Only the `generate_preview()` usage is replaced. |
| All C# files (`ExcelAPI/`) | The C# runtime coordinate pipeline (`MeasureFieldsFromCom`) is being phased out for the pixel scan pipeline. This is a Python-side change only. |

### Migration Steps (Ordered)

```
Step 1: Create new function in upload_coordinate_generator.py
─────────────────────────────────────────────────────────
- Write `generate_coordinates_and_preview(xlsx_path, output_dir, output_id)`
- Implements the single-session pipeline from Part 4
- Test as a standalone function (not yet connected to the endpoint)

Step 2: Verify pixel geometry
─────────────────────────────
- Run with FormTest and Japanese workbook
- Compare sanitized PDF vs original PDF (page dimensions, content position)
- Compare coordinates against current generate_coordinates() output
- Ensure < 2px difference in all coordinates

Step 3: Update /upload/preview endpoint
────────────────────────────────────────
- Replace `generate_preview()` call with `generate_coordinates_and_preview()`
- Remove import of old functions from app.py
- Test HTTP endpoint

Step 4: Clean up old functions
───────────────────────────────
- Remove or deprecate: `_identify_clusters_from_fields_sheet()`, `sanitize_workbook()`, `export_sanitized_pdf()`
- Keep `generate_coordinates()` if needed by /upload/coordinates endpoint
- Add deprecation warnings to old functions

Step 5: Fix IgnorePrintAreas
─────────────────────────────
- Change the ExportAsFixedFormat parameter from True → False (or remove)
- This matches ConMas default behavior

Step 6: Investigate GetAddress divergence
──────────────────────────────────────────
- Determine why scan_black_rectangles merges adjacent clusters
- ConMas GetAddress separates them naturally
- Fix the pixel scan to avoid needing split_merged_rects()

Step 7: Regression test
────────────────────────
- Run both workbooks through the complete pipeline
- Verify: coordinates, background, page dimensions, performance
- Check: no timeout, no COM deadlock, no orphan Excel processes

Step 8: Phase out COM Range pipeline
─────────────────────────────────────
- Mark `MeasureFieldsFromCom` as deprecated
- Mark `CoordinateTransformer.CellToPagePt()` as deprecated
- Keep existing code for reference but add deprecation comments
```

---

## Appendix A — Performance Target

### Current Performance

| Stage | FormTest | Japanese | 
|-------|----------|----------|
| Total | 26.8s | 83.8s |
| COM overhead | ~7s (4 sessions) | ~41s (4 sessions + 1935 cells × 2 passes) |

### Target Performance After Redesign

| Stage | FormTest (target) | Japanese (target) | Savings |
|-------|-------------------|-------------------|---------|
| Single COM session | ~5s | ~35s | ~2-6s (fewer sessions) |
| Combined pass (scan+sanitize) | ~2s | ~15s | ~3-15s (half the COM calls) |
| Pixel scan | ~19s | ~40s | Same (unavoidable) |
| Background PDF/render | ~1s | ~2.5s | Same (unavoidable) |
| **Total target** | **~25-27s** | **~55-75s** | **~2-17s saved** |

The dominant bottleneck remains `scan_black_rectangles` at ~40s for 2550×3300 pixels (48% of total). This is inherent to the NumPy algorithm and matches what ConMas would take for the same image size.

---

## Appendix B — Decision Matrix

### Which functions should merge?

| Old Function | New Home | Why |
|-------------|----------|-----|
| `_identify_clusters_from_comments()` | Inlined into combined pipeline | Only called from `generate_coordinates()` |
| `_identify_clusters_from_fields_sheet()` | Keep as fallback (or remove) | ConMas doesn't use it for coordinates |
| `_identify_clusters()` | Keep as wrapper (refactored) | Needed by `/upload/coordinates` endpoint |
| `sanitize_workbook()` | Inlined into combined pipeline | Only called from `generate_coordinates()` |
| `export_sanitized_pdf()` | Inlined into combined pipeline | Only called from `generate_coordinates()` |
| `xlsx_to_pdf()` | Keep as standalone utility | Called by `renderer.py` and other modules |
| `scan_black_rectangles()` | Keep unchanged | Non-COM, correct algorithm |
| `split_merged_rects()` | Keep (until root cause fixed) | Workaround for pixel scan bug |
| `normalize_rects()` | Keep unchanged | Non-COM, correct formula |
| `render_pdf_to_image()` | Keep unchanged | Non-COM, correct implementation |
| `generate_coordinates()` | Keep (refactored to call combined pipeline) | Needed by `/upload/coordinates` endpoint |
| `generate_preview()` | **Remove** (replaced by combined pipeline) | Redundant after redesign |
| `generate_coordinates_and_preview()` | **New** (combined pipeline) | Primary production function |

---

*End of architectural blueprint. All conclusions are backed by IL decompilation evidence, runtime trace data, and comparative validation from Phases X.1 through X.8. No code changes should be made until this blueprint is reviewed and approved.*
