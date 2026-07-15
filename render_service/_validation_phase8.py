"""
Phase X.8 — Production Validation Script

Collects all evidence needed for Parts 1, 3, 4, and 6.
Runs both workbooks through the pipeline and captures:
  - Per-stage timing
  - Pixel scan coordinates vs COM Range coordinates
  - Coordinate accuracy (average, max, RMS error)
  - COM lifecycle logging
"""

import os
import sys
import time
import json
import tempfile
import shutil
import subprocess
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

JAPANESE = r"C:\Users\MCF-JO~1\Documents\[V3.1_Sample]アンケート用紙.xlsx"
FORMTEST = r"C:\Users\MCF-JO~1\Documents\FormTest - Copy.xlsx"
OUTPUT_DIR = os.path.join(_THIS_DIR, "timing_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
DPI = 300
PT_TO_PX = DPI / 72.0


def check_orphan_excel():
    """Check for orphaned EXCEL.EXE processes before/after."""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=10
        )
        lines = [l for l in result.stdout.split('\n') if 'EXCEL' in l.upper()]
        return len(lines)
    except Exception:
        return -1


def get_com_range_coords(xlsx_path):
    """Read COM Range.Left/Top/Width/Height for every cluster cell."""
    import pythoncom
    import win32com.client
    from render_service.upload_coordinate_generator import _identify_clusters

    pythoncom.CoInitialize()
    try:
        clusters = _identify_clusters(xlsx_path)
        if not clusters:
            return [], {}

        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

            # Find active data sheet (not _Fields)
            data_sheet = None
            for i in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets(i)
                if ws.Visible == -1 and ws.Name != "_Fields":
                    data_sheet = ws
                    break
            if data_sheet is None:
                data_sheet = wb.ActiveSheet

            # Page setup info
            try:
                ps = data_sheet.PageSetup
                page_w_pt = 8.5 * 72  # Letter default
                page_h_pt = 11 * 72
                left_margin = ps.LeftMargin
                right_margin = ps.RightMargin
                top_margin = ps.TopMargin
                bottom_margin = ps.BottomMargin
                center_h = ps.CenterHorizontally
                center_v = ps.CenterVertically
                zoom = ps.Zoom
                fit_w = ps.FitToPagesWide
                fit_h = ps.FitToPagesTall
            except Exception:
                page_w_pt = 8.5 * 72
                page_h_pt = 11 * 72
                left_margin = 0.7 * 72
                right_margin = 0.7 * 72
                top_margin = 0.75 * 72
                bottom_margin = 0.75 * 72
                center_h = False
                center_v = False
                zoom = 100
                fit_w = None
                fit_h = None

            printable_w = page_w_pt - left_margin - right_margin
            printable_h = page_h_pt - top_margin - bottom_margin

            # Get PrintArea range or used range
            try:
                pa_str = ps.PrintArea
                if pa_str:
                    pa_range = data_sheet.Range(pa_str)
                    content_w = float(pa_range.Width)
                    content_h = float(pa_range.Height)
                else:
                    used = data_sheet.UsedRange
                    content_w = float(used.Width)
                    content_h = float(used.Height)
            except Exception:
                used = data_sheet.UsedRange
                content_w = float(used.Width)
                content_h = float(used.Height)

            # Compute origin
            origin_x = left_margin
            origin_y = top_margin
            if center_h:
                origin_x = left_margin + (printable_w - content_w) / 2
            if center_v:
                origin_y = top_margin + (printable_h - content_h) / 2

            results = []
            for c in clusters:
                addr = c["cellAddr"]
                try:
                    rng = data_sheet.Range(addr)
                    left_pt = float(rng.Left)
                    top_pt = float(rng.Top)
                    width_pt = float(rng.Width)
                    height_pt = float(rng.Height)

                    page_left_pt = origin_x + left_pt
                    page_top_pt = origin_y + top_pt
                    page_right_pt = page_left_pt + width_pt
                    page_bottom_pt = page_top_pt + height_pt

                    left_px = page_left_pt * PT_TO_PX
                    top_px = page_top_pt * PT_TO_PX
                    right_px = page_right_pt * PT_TO_PX
                    bottom_px = page_bottom_pt * PT_TO_PX

                    results.append({
                        "cellAddr": addr,
                        "name": c["name"],
                        "range_left_pt": left_pt,
                        "range_top_pt": top_pt,
                        "range_width_pt": width_pt,
                        "range_height_pt": height_pt,
                        "page_left_pt": page_left_pt,
                        "page_top_pt": page_top_pt,
                        "page_width_pt": width_pt,
                        "page_height_pt": height_pt,
                        "left_px": left_px,
                        "top_px": top_px,
                        "right_px": right_px,
                        "bottom_px": bottom_px,
                        "width_px": right_px - left_px,
                        "height_px": bottom_px - top_px,
                    })
                except Exception as e:
                    print(f"  COM Range failed for {addr}: {e}")

            wb.Close(False)

            page_info = {
                "page_w_pt": page_w_pt,
                "page_h_pt": page_h_pt,
                "left_margin": left_margin,
                "right_margin": right_margin,
                "top_margin": top_margin,
                "bottom_margin": bottom_margin,
                "center_h": center_h,
                "center_v": center_v,
                "zoom": zoom,
                "fit_w": fit_w,
                "fit_h": fit_h,
                "origin_x_pt": origin_x,
                "origin_y_pt": origin_y,
            }
            return results, page_info
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()


def run_timed_pipeline(xlsx_path, label):
    """Run the full generate_preview pipeline with per-stage timing."""
    from render_service.upload_coordinate_generator import (
        _identify_clusters, sanitize_workbook, export_sanitized_pdf,
        render_pdf_to_image, scan_black_rectangles, split_merged_rects,
        normalize_rects, _sort_key_meta
    )
    from render_service.pdf_converter import xlsx_to_pdf
    from render_service.background_renderer import pdf_page_to_png, get_page_dimensions

    timings = {}
    tmp_dir = tempfile.mkdtemp(prefix="ple_val_")

    # Stage 1: Identify clusters
    t0 = time.perf_counter()
    cluster_meta = _identify_clusters(xlsx_path)
    timings["identify_clusters"] = (time.perf_counter() - t0) * 1000
    print(f"  [identify_clusters] {timings['identify_clusters']:.0f}ms -> {len(cluster_meta)} clusters")

    # Stage 2: Sanitize
    t0 = time.perf_counter()
    sanitized_path = sanitize_workbook(xlsx_path, cluster_meta)
    timings["sanitize_workbook"] = (time.perf_counter() - t0) * 1000
    print(f"  [sanitize_workbook] {timings['sanitize_workbook']:.0f}ms")

    # Stage 3: Export sanitized PDF
    t0 = time.perf_counter()
    pdf_path = export_sanitized_pdf(sanitized_path)
    timings["export_sanitized_pdf"] = (time.perf_counter() - t0) * 1000
    print(f"  [export_sanitized_pdf] {timings['export_sanitized_pdf']:.0f}ms")

    # Stage 4: Render PDF
    t0 = time.perf_counter()
    img, img_w, img_h = render_pdf_to_image(pdf_path)
    timings["render_sanitized_pdf"] = (time.perf_counter() - t0) * 1000
    print(f"  [render_sanitized_pdf] {timings['render_sanitized_pdf']:.0f}ms -> {img_w}x{img_h}")

    # Stage 5: Scan rectangles
    t0 = time.perf_counter()
    rects = scan_black_rectangles(img)
    timings["scan_black_rectangles"] = (time.perf_counter() - t0) * 1000
    print(f"  [scan_black_rectangles] {timings['scan_black_rectangles']:.0f}ms -> {len(rects)} rects")

    # Stage 6: Split merged
    t0 = time.perf_counter()
    split_rects = split_merged_rects(rects, cluster_meta)
    timings["split_merged_rects"] = (time.perf_counter() - t0) * 1000
    print(f"  [split_merged_rects] {timings['split_merged_rects']:.0f}ms -> {len(split_rects)}")

    # Stage 7: Normalize
    t0 = time.perf_counter()
    normalized = normalize_rects(split_rects, img_w, img_h)
    timings["normalize_rects"] = (time.perf_counter() - t0) * 1000

    # Build fields
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
            # Pixel coordinates (for comparison)
            "left_px": rect["Left"],
            "top_px": rect["Top"],
            "right_px": rect["Right"],
            "bottom_px": rect["Bottom"],
            "width_px": rect["Right"] - rect["Left"],
            "height_px": rect["Bottom"] - rect["Top"],
        })
    print(f"  [build_fields] -> {len(fields)} fields")

    # Stage 8: Export original PDF
    t0 = time.perf_counter()
    orig_pdf_path = xlsx_to_pdf(xlsx_path)
    timings["export_original_pdf"] = (time.perf_counter() - t0) * 1000
    print(f"  [export_original_pdf] {timings['export_original_pdf']:.0f}ms")

    # Stage 9: Render background PNG
    t0 = time.perf_counter()
    png_path = os.path.join(OUTPUT_DIR, f"val_{label}.png")
    pdf_page_to_png(orig_pdf_path, png_path, dpi=DPI)
    page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=DPI)
    timings["render_background_png"] = (time.perf_counter() - t0) * 1000
    print(f"  [render_background_png] {timings['render_background_png']:.0f}ms -> {page_w}x{page_h}")

    # Cleanup
    try: os.unlink(orig_pdf_path)
    except: pass
    try: shutil.rmtree(os.path.dirname(orig_pdf_path), ignore_errors=True)
    except: pass
    try: shutil.rmtree(tmp_dir, ignore_errors=True)
    except: pass
    # Clean sanitized/pdf temp dirs
    for p in [sanitized_path, pdf_path]:
        try: shutil.rmtree(os.path.dirname(p), ignore_errors=True)
        except: pass

    total = sum(timings.values())
    return {
        "label": label,
        "timings_ms": timings,
        "total_ms": total,
        "total_seconds": total / 1000,
        "fields": fields,
        "page_dimensions": {"width": page_w, "height": page_h},
    }


def compute_coordinate_accuracy(pixel_fields, com_coords, page_dimensions):
    """Compare pixel scan coordinates against COM Range coordinates."""
    page_w, page_h = page_dimensions["width"], page_dimensions["height"]
    comparisons = []

    com_by_addr = {c["cellAddr"]: c for c in com_coords}

    for pf in pixel_fields:
        addr = pf["cellAddr"]
        com = com_by_addr.get(addr)
        if not com:
            continue

        # Convert pixel scan ratios back to pixels
        pixel_left = pf["left_px"]
        pixel_top = pf["top_px"]
        pixel_right = pf["right_px"]
        pixel_bottom = pf["bottom_px"]

        # COM Range in pixels (from page coordinates)
        com_left = com["left_px"]
        com_top = com["top_px"]
        com_right = com["right_px"]
        com_bottom = com["bottom_px"]

        # Differences in pixels
        diff_left = abs(pixel_left - com_left)
        diff_top = abs(pixel_top - com_top)
        diff_right = abs(pixel_right - com_right)
        diff_bottom = abs(pixel_bottom - com_bottom)
        diff_width = abs((pixel_right - pixel_left) - (com_right - com_left))
        diff_height = abs((pixel_bottom - pixel_top) - (com_bottom - com_top))
        max_diff = max(diff_left, diff_top, diff_right, diff_bottom)

        comparisons.append({
            "cellAddr": addr,
            "name": pf["name"],
            "pixel_left_ratio": pf["left_ratio"],
            "pixel_top_ratio": pf["top_ratio"],
            "pixel_left_px": round(pixel_left, 1),
            "pixel_top_px": round(pixel_top, 1),
            "pixel_right_px": round(pixel_right, 1),
            "pixel_bottom_px": round(pixel_bottom, 1),
            "pixel_width_px": round(pixel_right - pixel_left, 1),
            "pixel_height_px": round(pixel_bottom - pixel_top, 1),
            "com_left_px": round(com_left, 1),
            "com_top_px": round(com_top, 1),
            "com_right_px": round(com_right, 1),
            "com_bottom_px": round(com_bottom, 1),
            "com_width_px": round(com_right - com_left, 1),
            "com_height_px": round(com_bottom - com_top, 1),
            "diff_left_px": round(diff_left, 2),
            "diff_top_px": round(diff_top, 2),
            "diff_right_px": round(diff_right, 2),
            "diff_bottom_px": round(diff_bottom, 2),
            "diff_width_px": round(diff_width, 2),
            "diff_height_px": round(diff_height, 2),
            "max_diff_px": round(max_diff, 2),
        })

    if not comparisons:
        return {"error": "No matching fields", "comparisons": []}

    # Aggregate stats
    max_diffs = [c["max_diff_px"] for c in comparisons]
    avg_error = sum(max_diffs) / len(max_diffs)
    max_error = max(max_diffs)
    rms_error = (sum(d**2 for d in max_diffs) / len(max_diffs)) ** 0.5

    return {
        "average_error_px": round(avg_error, 2),
        "max_error_px": round(max_error, 2),
        "rms_error_px": round(rms_error, 2),
        "n_fields": len(comparisons),
        "passes_criteria": avg_error < 2.0,
        "comparisons": comparisons,
    }


def main():
    report_path = os.path.join(OUTPUT_DIR, "phase8_validation_report.txt")
    report_json_path = os.path.join(OUTPUT_DIR, "phase8_validation_data.json")

    with open(report_path, "w", encoding="utf-8") as f:
        old_stdout = sys.stdout
        sys.stdout = f

        print("=" * 70)
        print("  Phase X.8 — Production Validation Report")
        print("=" * 70)

        # Part 3: COM lifecycle — check orphan processes BEFORE
        print("\n--- PART 3: COM Lifecycle ---")
        before = check_orphan_excel()
        print(f"  EXCEL.EXE processes before: {before}")

        all_data = {}

        for label, path in [("FormTest", FORMTEST), ("Japanese", JAPANESE)]:
            print(f"\n{'='*70}")
            print(f"  Workbook: {label}")
            print(f"{'='*70}")

            # Get COM Range coordinates
            print(f"\n  [COM Range Coordinates]")
            try:
                com_coords, page_info = get_com_range_coords(path)
                print(f"  Found {len(com_coords)} COM Range entries")
                print(f"  Page info: {json.dumps({k: v for k, v in page_info.items() if 'margin' not in k and 'origin' not in k})}")
            except Exception as e:
                print(f"  COM Range FAILED: {e}")
                import traceback as tb
                tb.print_exc()
                com_coords = []
                page_info = {}

            # Run timed pipeline
            print(f"\n  [Timed Pipeline]")
            result = run_timed_pipeline(path, label)
            print(f"  Total: {result['total_seconds']:.1f}s")
            print(f"  Fields: {len(result['fields'])}")
            print(f"  Page: {result['page_dimensions']}")

            # Part 1: Coordinate accuracy comparison
            print(f"\n  [Coordinate Accuracy]")
            accuracy = compute_coordinate_accuracy(result["fields"], com_coords, result["page_dimensions"])
            print(f"  Average error: {accuracy.get('average_error_px', 'N/A')} px")
            print(f"  Max error: {accuracy.get('max_error_px', 'N/A')} px")
            print(f"  RMS error: {accuracy.get('rms_error_px', 'N/A')} px")
            print(f"  Passes <2px criteria: {accuracy.get('passes_criteria', 'N/A')}")

            # Per-field comparison table
            if accuracy.get("comparisons"):
                print(f"\n  {'Cell':<20s} {'Pixel L':>10s} {'COM L':>10s} {'Diff':>8s} {'Pixel T':>10s} {'COM T':>10s} {'Diff':>8s}")
                print(f"  {'-'*70}")
                for c in accuracy["comparisons"]:
                    print(f"  {c['cellAddr']:<20s} {c['pixel_left_px']:>8.1f}px {c['com_left_px']:>8.1f}px {c['diff_left_px']:>6.2f}px "
                          f"{c['pixel_top_px']:>8.1f}px {c['com_top_px']:>8.1f}px {c['diff_top_px']:>6.2f}px")

            # Per-stage timing table
            print(f"\n  [Stage Timing]")
            print(f"  {'Stage':<30s} {'Time (ms)':>10s}")
            print(f"  {'-'*42}")
            for stage, t in sorted(result["timings_ms"].items(), key=lambda x: x[1], reverse=True):
                pct = t / result["total_ms"] * 100
                print(f"  {stage:<30s} {t:>8.0f}ms ({pct:>4.0f}%)")
            print(f"  {'-'*42}")
            print(f"  {'TOTAL':<30s} {result['total_ms']:>8.0f}ms")

            all_data[label] = {
                "com_coords": com_coords,
                "page_info": page_info,
                "pipeline_result": {k: v for k, v in result.items() if k != "fields"},
                "fields": result["fields"],
                "accuracy": accuracy,
            }

        # Part 3: Check orphan processes AFTER
        after = check_orphan_excel()
        print(f"\n  EXCEL.EXE processes before: {before}")
        print(f"  EXCEL.EXE processes after:  {after}")
        print(f"  Orphaned: {'NO' if after <= before + 1 else f'YES (+{after - before})'}")

        # Summary
        print(f"\n{'='*70}")
        print("  SUMMARY")
        print(f"{'='*70}")
        for label in ["FormTest", "Japanese"]:
            d = all_data.get(label, {})
            a = d.get("accuracy", {})
            r = d.get("pipeline_result", {})
            print(f"\n  {label}:")
            print(f"    Total time: {r.get('total_seconds', 0):.1f}s")
            print(f"    Fields: {r.get('fields_count', r.get('fields_count', 0))}")
            print(f"    Avg coordinate error: {a.get('average_error_px', 'N/A')}px")
            print(f"    Max error: {a.get('max_error_px', 'N/A')}px")
            print(f"    Passes <2px: {a.get('passes_criteria', 'N/A')}")

        sys.stdout = old_stdout

    # Save JSON data
    with open(report_json_path, "w", encoding="utf-8") as f:
        json.dump(all_data, f, indent=2, default=str, ensure_ascii=False)

    print(f"Report saved to: {report_path}")
    print(f"Data saved to: {report_json_path}")


if __name__ == "__main__":
    main()
