"""
Phase X.4 — Performance Regression Timing Analysis

Measures execution time for every stage inside generate_preview().
Collects worksheet statistics to identify the bottleneck.

Do NOT modify the pipeline. Timing only.
"""

import os
import sys
import time
import json
import tempfile
import shutil
import traceback
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

JAPANESE_WORKBOOK = r"C:\Users\MCF-JO~1\Documents\[V3.1_Sample]アンケート用紙.xlsx"
FORMTEST_WORKBOOK = r"C:\Users\MCF-JO~1\Documents\FormTest - Copy.xlsx"

OUTPUT_DIR = os.path.join(_THIS_DIR, "timing_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

DPI = 300
PT_TO_PX = DPI / 72.0


# ──────────────────────────────────────────────────────────────────────────────
# Stage 0 — Worksheet Statistics
# ──────────────────────────────────────────────────────────────────────────────

def collect_worksheet_stats(xlsx_path: str) -> dict:
    """Collect statistics about the worksheet contents -- shapes, pictures, etc."""
    import win32com.client
    stats = {}
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        stats["sheet_count"] = wb.Worksheets.Count
        total_shapes = 0
        total_pictures = 0
        total_textboxes = 0
        total_charts = 0
        total_ole = 0
        total_merged = 0
        max_used_rows = 0
        max_used_cols = 0
        sheet_names = []

        for ws in wb.Worksheets:
            name = ws.Name
            sheet_names.append(name)
            try:
                shapes = ws.Shapes
                n_shapes = shapes.Count
                total_shapes += n_shapes
                # Categorize shapes
                for i in range(1, n_shapes + 1):
                    try:
                        sh = shapes.Item(i)
                        sh_type = sh.Type
                        # msoPicture=13, msoTextBox=17, msoChart=3, msoOLEControlObject=12
                        if sh_type == 13:
                            total_pictures += 1
                        elif sh_type == 17:
                            total_textboxes += 1
                        elif sh_type == 3:
                            total_charts += 1
                        elif sh_type == 12:
                            total_ole += 1
                    except Exception:
                        pass
            except Exception:
                pass

            try:
                used = ws.UsedRange
                if used:
                    ur = used.Row + used.Rows.Count - 1
                    uc = used.Column + used.Columns.Count - 1
                    max_used_rows = max(max_used_rows, ur)
                    max_used_cols = max(max_used_cols, uc)
            except Exception:
                pass

            try:
                total_merged += ws.UsedRange.MergeAreas.Count if ws.UsedRange else 0
            except Exception:
                pass

        stats["sheet_names"] = sheet_names
        stats["total_shapes"] = total_shapes
        stats["total_pictures"] = total_pictures
        stats["total_textboxes"] = total_textboxes
        stats["total_charts"] = total_charts
        stats["total_ole_objects"] = total_ole
        stats["max_used_rows"] = max_used_rows
        stats["max_used_cols"] = max_used_cols
        stats["total_merged_areas"] = total_merged

        # PrintArea (first sheet)
        try:
            pa = wb.ActiveSheet.PageSetup.PrintArea
            stats["print_area"] = str(pa) if pa else "None"
        except Exception:
            stats["print_area"] = "Error"

        # Zoom / FitToPages
        try:
            ws1 = wb.Worksheets(1)
            stats["zoom"] = str(ws1.PageSetup.Zoom)
            stats["fit_to_pages_wide"] = str(ws1.PageSetup.FitToPagesWide)
            stats["fit_to_pages_tall"] = str(ws1.PageSetup.FitToPagesTall)
            stats["orientation"] = str(ws1.PageSetup.Orientation)
            stats["paper_size"] = str(ws1.PageSetup.PaperSize)
        except Exception:
            pass

        # File size
        stats["file_size_bytes"] = os.path.getsize(xlsx_path)

        wb.Close(False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    return stats


# ──────────────────────────────────────────────────────────────────────────────
# Timed Pipeline Runner
# ──────────────────────────────────────────────────────────────────────────────

def timed_stage(name: str, func, *args, **kwargs) -> tuple:
    """Run a function with timing. Returns (result, elapsed_ms)."""
    t0 = time.perf_counter()
    try:
        result = func(*args, **kwargs)
        elapsed = (time.perf_counter() - t0) * 1000
        print(f"  [{name}] {elapsed:.0f} ms")
        return result, elapsed
    except Exception as e:
        elapsed = (time.perf_counter() - t0) * 1000
        print(f"  [{name}] FAILED after {elapsed:.0f} ms: {e}")
        raise


def run_timed_pipeline(xlsx_path: str, workbook_label: str) -> dict:
    """Run the full generate_preview pipeline with per-stage timing."""
    print(f"\n{'='*70}")
    print(f"  TIMING ANALYSIS: {workbook_label}")
    print(f"  File: {xlsx_path}")
    print(f"{'='*70}")

    timings = {}
    fields = None
    background_info = None

    tmp_dir = tempfile.mkdtemp(prefix="ple_timing_")

    try:
        # ── Stage 1: Identify clusters ─────────────────────────────────
        from render_service.upload_coordinate_generator import (
            _identify_clusters, sanitize_workbook, export_sanitized_pdf,
            render_pdf_to_image, scan_black_rectangles, split_merged_rects,
            normalize_rects, _sort_key_meta
        )

        cluster_meta, t1 = timed_stage("Identify Clusters",
            _identify_clusters, xlsx_path)
        timings["identify_clusters"] = t1
        print(f"         → {len(cluster_meta)} clusters found")

        # ── Stage 2: Sanitize workbook ─────────────────────────────────
        sanitized_path, t2 = timed_stage("Sanitize Workbook",
            sanitize_workbook, xlsx_path, cluster_meta)
        timings["sanitize_workbook"] = t2
        sanitized_size = os.path.getsize(sanitized_path) if os.path.exists(sanitized_path) else 0
        print(f"         → {os.path.basename(sanitized_path)} ({sanitized_size} bytes)")

        # ── Stage 3: Export sanitized PDF ───────────────────────────────
        t3_start = time.perf_counter()
        pdf_path, t3 = timed_stage("Export Sanitized PDF",
            export_sanitized_pdf, sanitized_path)
        timings["export_sanitized_pdf"] = t3
        pdf_size = os.path.getsize(pdf_path) if os.path.exists(pdf_path) else 0
        print(f"         → {os.path.basename(pdf_path)} ({pdf_size} bytes)")

        # Count pages in sanitized PDF
        try:
            import fitz
            doc = fitz.open(pdf_path)
            sanitized_page_count = len(doc)
            doc.close()
            print(f"         → {sanitized_page_count} page(s)")
        except Exception:
            sanitized_page_count = 0

        # ── Stage 4: Render sanitized PDF to image ──────────────────────
        render_result, t4 = timed_stage("Render Sanitized PDF",
            render_pdf_to_image, pdf_path)
        img, img_w, img_h = render_result
        timings["render_sanitized_pdf"] = t4
        print(f"         → {img_w}x{img_h} px")

        # ── Stage 5: Scan black rectangles ──────────────────────────────
        rects, t5 = timed_stage("Scan Black Rectangles",
            scan_black_rectangles, img)
        timings["scan_black_rectangles"] = t5
        print(f"         → {len(rects)} rectangles found")

        # ── Stage 6: Split merged rects ─────────────────────────────────
        split_rects, t6 = timed_stage("Split Merged Rects",
            split_merged_rects, rects, cluster_meta)
        timings["split_merged_rects"] = t6
        print(f"         → {len(split_rects)} after split")

        # ── Stage 7: Normalize ──────────────────────────────────────────
        normalized, t7 = timed_stage("Normalize Rects",
            normalize_rects, split_rects, img_w, img_h)
        timings["normalize_rects"] = t7

        # Sort and merge
        cluster_meta.sort(key=_sort_key_meta)
        normalized.sort(key=lambda r: (r["Top"], r["Left"]))
        fields = []
        for meta, rect in zip(cluster_meta, normalized):
            fields.append({
                "name": meta["name"],
                "type": meta["type"],
                "cellAddr": meta["cellAddr"],
                "left_ratio": round(rect["left_ratio"], 7),
                "top_ratio": round(rect["top_ratio"], 7),
                "right_ratio": round(rect["right_ratio"], 7),
                "bottom_ratio": round(rect["bottom_ratio"], 7),
            })
        print(f"         → {len(fields)} fields generated")

        # ── Stage 8: Export original workbook PDF ───────────────────────
        from render_service.pdf_converter import xlsx_to_pdf
        orig_pdf_path, t8 = timed_stage("Export Original PDF",
            xlsx_to_pdf, xlsx_path)
        timings["export_original_pdf"] = t8
        orig_pdf_size = os.path.getsize(orig_pdf_path) if os.path.exists(orig_pdf_path) else 0
        print(f"         → {os.path.basename(orig_pdf_path)} ({orig_pdf_size} bytes)")

        # Count pages in original PDF
        try:
            doc = fitz.open(orig_pdf_path)
            orig_page_count = len(doc)
            doc.close()
            print(f"         → {orig_page_count} page(s)")
        except Exception:
            orig_page_count = 0

        # ── Stage 9: Render original PDF to PNG ─────────────────────────
        from render_service.background_renderer import pdf_page_to_png, get_page_dimensions
        png_filename = "timing_background.png"
        png_path = os.path.join(OUTPUT_DIR, png_filename)
        t9_start = time.perf_counter()
        pdf_page_to_png(orig_pdf_path, png_path, dpi=DPI)
        t9 = (time.perf_counter() - t9_start) * 1000
        timings["render_original_png"] = t9
        print(f"  [Render Original PNG] {t9:.0f} ms")

        page_dim_result, t_pd = timed_stage("Get Page Dimensions",
            get_page_dimensions, orig_pdf_path, dpi=DPI)
        page_w, page_h = page_dim_result
        timings["get_page_dimensions"] = t_pd
        print(f"         → {page_w}x{page_h} px")

        # Cleanup original PDF
        try:
            os.unlink(orig_pdf_path)
        except Exception:
            pass
        pdf_dir = os.path.dirname(orig_pdf_path)
        if pdf_dir and pdf_dir != tmp_dir:
            shutil.rmtree(pdf_dir, ignore_errors=True)

    finally:
        # Cleanup temp
        shutil.rmtree(tmp_dir, ignore_errors=True)

    # ── Build totals ────────────────────────────────────────────────────
    total = sum(timings.values())
    result = {
        "workbook": workbook_label,
        "timings_ms": timings,
        "total_ms": total,
        "total_seconds": total / 1000,
        "fields_count": len(fields) if fields else 0,
        "sanitized_pdf_size": pdf_size,
        "sanitized_page_count": sanitized_page_count,
        "orig_pdf_size": orig_pdf_size,
        "orig_page_count": orig_page_count,
    }

    return result


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def main():
    # Redirect stdout to file to avoid Unicode encoding issues on console
    report_path = os.path.join(OUTPUT_DIR, "timing_report.txt")
    sys.stdout = open(report_path, "w", encoding="utf-8")
    
    print("=" * 70)
    print("  Phase X.4 - Performance Regression Timing Analysis")
    print("=" * 70)

    for label, path in [("FormTest", FORMTEST_WORKBOOK), ("Japanese", JAPANESE_WORKBOOK)]:
        print(f"\n  Workbook: {label}")
        try:
            t0 = time.perf_counter()
            stats = collect_worksheet_stats(path)
            elapsed = (time.perf_counter() - t0) * 1000
            print(f"  Stats collected in {elapsed:.0f} ms:")
            for k, v in stats.items():
                print(f"    {k}: {v}")
        except Exception as e:
            print(f"  FAILED: {e}")
            traceback.print_exc()

    # ── Step 2: Run timed pipeline on FormTest ──────────────────────────
    print()
    print("-" * 70)
    print("  STEP 2: Timed Pipeline - FormTest (baseline)")
    print("-" * 70)
    try:
        formtest_result = run_timed_pipeline(FORMTEST_WORKBOOK, "FormTest")
    except Exception as e:
        print(f"  FAILED: {e}")
        traceback.print_exc()
        formtest_result = None

    # ── Step 3: Run timed pipeline on Japanese workbook ─────────────────
    print()
    print("-" * 70)
    print("  STEP 3: Timed Pipeline - Japanese Workbook")
    print("-" * 70)
    try:
        japan_result = run_timed_pipeline(JAPANESE_WORKBOOK, "Japanese")
    except Exception as e:
        print(f"  FAILED: {e}")
        traceback.print_exc()
        japan_result = None

    # ── Step 4: Summary Table ───────────────────────────────────────────
    print("\n" + "=" * 70)
    print("  SUMMARY: Stage-by-Stage Timing Comparison")
    print("=" * 70)

    all_stages = [
        "identify_clusters", "sanitize_workbook", "export_sanitized_pdf",
        "render_sanitized_pdf", "scan_black_rectangles", "split_merged_rects",
        "normalize_rects", "export_original_pdf", "render_original_png",
        "get_page_dimensions"
    ]

    headers = ["Stage", "FormTest (ms)", "%", "Japanese (ms)", "%"]
    col_w = [32, 14, 8, 16, 8]
    def fmt_row(cols):
        return "  ".join(c.ljust(w) for c, w in zip(cols, col_w))

    print()
    print(fmt_row(headers))
    print("  " + "-" * (sum(col_w) + len(col_w) - 1))

    f_res = formtest_result or {"timings_ms": {}, "total_ms": 0}
    j_res = japan_result or {"timings_ms": {}, "total_ms": 0}

    for stage in all_stages:
        ft = f_res["timings_ms"].get(stage, 0)
        jt = j_res["timings_ms"].get(stage, 0)
        if ft == 0 and jt == 0:
            continue
        ft_pct = ft / f_res["total_ms"] * 100 if f_res["total_ms"] > 0 else 0
        jt_pct = jt / j_res["total_ms"] * 100 if j_res["total_ms"] > 0 else 0
        print(fmt_row([
            stage.replace("_", " ").title(),
            f"{ft:.0f}",
            f"{ft_pct:.0f}%",
            f"{jt:.0f}",
            f"{jt_pct:.0f}%",
        ]))

    print("  " + "-" * (sum(col_w) + len(col_w) - 1))
    print(fmt_row([
        "TOTAL",
        f"{f_res['total_ms']:.0f}",
        "100%",
        f"{j_res['total_ms']:.0f}",
        "100%",
    ]))
    print()

    # Additional metrics
    if formtest_result:
        print(f"  FormTest: {formtest_result['fields_count']} fields, "
              f"sanitized PDF: {formtest_result.get('sanitized_page_count', '?')} page(s), "
              f"original PDF: {formtest_result.get('orig_page_count', '?')} page(s)")
    if japan_result:
        print(f"  Japanese: {japan_result['fields_count']} fields, "
              f"sanitized PDF: {japan_result.get('sanitized_page_count', '?')} page(s), "
              f"original PDF: {japan_result.get('orig_page_count', '?')} page(s)")

    # ── Save results to JSON ────────────────────────────────────────────
    report = {
        "formtest": formtest_result,
        "japanese": japan_result,
        "formtest_stats": None,
        "japanese_stats": None,
    }
    report_path = os.path.join(OUTPUT_DIR, "timing_report.json")
    with open(report_path, "w") as f:
        json.dump(report, f, indent=2, default=str)
    print(f"\n  Report saved to: {report_path}")

    # Determine if timeout is likely
    print()
    print("-" * 70)
    print("  TIMEOUT ANALYSIS")
    print("-" * 70)
    for label, result in [("FormTest", formtest_result), ("Japanese", japan_result)]:
        if result:
            total_s = result["total_seconds"]
            print(f"  {label}: {total_s:.1f}s total")
            if total_s > 100:
                print(f"    ⚠ EXCEEDS 100s timeout!")
            elif total_s > 90:
                print(f"    ⚠ Close to 100s timeout")
            else:
                print(f"    ✅ Under 100s timeout")

            # Find slowest stage
            if result["timings_ms"]:
                slowest = max(result["timings_ms"], key=result["timings_ms"].get)
                slowest_ms = result["timings_ms"][slowest]
                print(f"    Slowest stage: {slowest} ({slowest_ms:.0f} ms = {slowest_ms/1000:.1f}s)")

    print("\nDone.")


if __name__ == "__main__":
    main()
