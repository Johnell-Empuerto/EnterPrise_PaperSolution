"""
Phase X.5 — Production vs Benchmark Timing Investigation

Compares the exact same code path (generate_preview) run in-process
(benchmark) vs through the HTTP production endpoint (/upload/preview).

Measures timestamps at every stage via both paths.
Identifies where the extra ~40s comes from.
"""

import os
import sys
import time
import json
import tempfile
import shutil
import io
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

JAPANESE_WORKBOOK = r"C:\Users\MCF-JO~1\Documents\[V3.1_Sample]アンケート用紙.xlsx"
FORMTEST_WORKBOOK = r"C:\Users\MCF-JO~1\Documents\FormTest - Copy.xlsx"
OUTPUT_DIR = os.path.join(_THIS_DIR, "timing_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
DPI = 300


# ──────────────────────────────────────────────────────────────────────────────
# Path 1: In-process benchmark (same as _timing_analysis.py)
# ──────────────────────────────────────────────────────────────────────────────

def timed_stage(name: str, func, *args, **kwargs) -> tuple:
    t0 = time.perf_counter()
    try:
        result = func(*args, **kwargs)
        elapsed = (time.perf_counter() - t0) * 1000
        return result, elapsed, None
    except Exception as e:
        elapsed = (time.perf_counter() - t0) * 1000
        return None, elapsed, str(e)


def run_benchmark(xlsx_path: str, label: str) -> dict:
    """Run the generate_preview pipeline in-process (same as benchmark)."""
    from render_service.upload_coordinate_generator import (
        _identify_clusters, sanitize_workbook, export_sanitized_pdf,
        render_pdf_to_image, scan_black_rectangles, split_merged_rects,
        normalize_rects, _sort_key_meta
    )
    from render_service.pdf_converter import xlsx_to_pdf
    from render_service.background_renderer import pdf_page_to_png, get_page_dimensions

    timings = {}
    tmp_dir = tempfile.mkdtemp(prefix="ple_bench_")

    # Stage 1: Identify clusters
    cluster_meta, t1, err1 = timed_stage("identify_clusters",
        _identify_clusters, xlsx_path)
    timings["identify_clusters"] = {"ms": t1, "error": err1}
    n_clusters = len(cluster_meta) if cluster_meta else 0
    print(f"  [{label}] identify_clusters: {t1:.0f}ms -> {n_clusters} clusters")

    # Stage 2: Sanitize workbook
    sanitized_path, t2, err2 = timed_stage("sanitize_workbook",
        sanitize_workbook, xlsx_path, cluster_meta)
    timings["sanitize_workbook"] = {"ms": t2, "error": err2}
    print(f"  [{label}] sanitize_workbook: {t2:.0f}ms")

    # Stage 3: Export sanitized PDF
    pdf_path, t3, err3 = timed_stage("export_sanitized_pdf",
        export_sanitized_pdf, sanitized_path)
    timings["export_sanitized_pdf"] = {"ms": t3, "error": err3}
    print(f"  [{label}] export_sanitized_pdf: {t3:.0f}ms")

    # Stage 4: Render PDF to image
    render_result, t4, err4 = timed_stage("render_sanitized_pdf",
        render_pdf_to_image, pdf_path)
    img, img_w, img_h = render_result if render_result else (None, 0, 0)
    timings["render_sanitized_pdf"] = {"ms": t4, "error": err4}
    print(f"  [{label}] render_sanitized_pdf: {t4:.0f}ms -> {img_w}x{img_h}")

    # Stage 5: Scan black rectangles
    rects, t5, err5 = timed_stage("scan_black_rectangles",
        scan_black_rectangles, img)
    timings["scan_black_rectangles"] = {"ms": t5, "error": err5}
    print(f"  [{label}] scan_black_rectangles: {t5:.0f}ms -> {len(rects) if rects else 0} rects")

    # Stage 6: Split merged rects
    split_rects, t6, err6 = timed_stage("split_merged_rects",
        split_merged_rects, rects, cluster_meta)
    timings["split_merged_rects"] = {"ms": t6, "error": err6}
    print(f"  [{label}] split_merged_rects: {t6:.0f}ms -> {len(split_rects) if split_rects else 0}")

    # Stage 7: Normalize
    normalized, t7, err7 = timed_stage("normalize_rects",
        normalize_rects, split_rects, img_w, img_h)
    timings["normalize_rects"] = {"ms": t7, "error": err7}

    # Stage 8: Export original PDF
    orig_pdf_path, t8, err8 = timed_stage("export_original_pdf",
        xlsx_to_pdf, xlsx_path)
    timings["export_original_pdf"] = {"ms": t8, "error": err8}

    # Stage 9: Render original PNG
    png_path = os.path.join(OUTPUT_DIR, f"bench_{label}.png")
    t9 = 0
    if orig_pdf_path:
        t9_start = time.perf_counter()
        try:
            pdf_page_to_png(orig_pdf_path, png_path, dpi=DPI)
            t9 = (time.perf_counter() - t9_start) * 1000
        except Exception as e:
            t9 = (time.perf_counter() - t9_start) * 1000
            timings["render_original_png"] = {"ms": t9, "error": str(e)}
    timings["render_original_png"] = {"ms": t9, "error": None}
    print(f"  [{label}] render_original_png: {t9:.0f}ms")

    # Stage 10: Get page dimensions
    page_dim_result, t10, err10 = timed_stage("get_page_dimensions",
        get_page_dimensions, orig_pdf_path, dpi=DPI)
    timings["get_page_dimensions"] = {"ms": t10, "error": err10}

    # Cleanup
    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass
    try:
        if orig_pdf_path:
            os.unlink(orig_pdf_path)
            pdf_dir = os.path.dirname(orig_pdf_path)
            if pdf_dir and pdf_dir != tmp_dir:
                shutil.rmtree(pdf_dir, ignore_errors=True)
    except Exception:
        pass

    total = sum(v["ms"] for v in timings.values() if v["ms"] is not None)
    result = {
        "workbook": label,
        "timings_ms": {k: v["ms"] for k, v in timings.items()},
        "errors": {k: v["error"] for k, v in timings.items() if v["error"]},
        "total_ms": total,
        "total_seconds": total / 1000,
        "n_clusters": n_clusters,
    }
    return result


# ──────────────────────────────────────────────────────────────────────────────
# Path 2: Production HTTP endpoint
# ──────────────────────────────────────────────────────────────────────────────

def run_production_endpoint(xlsx_path: str, label: str) -> dict:
    """Call the actual /upload/preview endpoint via HTTP and measure total time."""
    import requests

    url = "http://127.0.0.1:5091/upload/preview"

    # Read file
    t_read_start = time.perf_counter()
    with open(xlsx_path, "rb") as f:
        file_bytes = f.read()
    t_read = (time.perf_counter() - t_read_start) * 1000

    # Build multipart request
    t_prep_start = time.perf_counter()
    files = {"file": (os.path.basename(xlsx_path), io.BytesIO(file_bytes),
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    params = {"output_dir": OUTPUT_DIR}
    t_prep = (time.perf_counter() - t_prep_start) * 1000

    # Send HTTP request
    t_http_start = time.perf_counter()
    try:
        response = requests.post(url, files=files, params=params, timeout=180)
        t_http = (time.perf_counter() - t_http_start) * 1000
        status = response.status_code
        try:
            data = response.json()
            n_fields = len(data.get("fields", []))
            success = data.get("success", False)
        except Exception:
            data = {}
            n_fields = 0
            success = False
    except requests.Timeout:
        t_http = (time.perf_counter() - t_http_start) * 1000
        print(f"  [{label}] HTTP REQUEST TIMED OUT after {t_http:.0f}ms!")
        response = None
        status = 0
        n_fields = 0
        success = False
    except Exception as e:
        t_http = (time.perf_counter() - t_http_start) * 1000
        response = None
        status = 0
        n_fields = 0
        success = False
        print(f"  [{label}] HTTP REQUEST FAILED: {e}")

    total = t_read + t_prep + t_http

    result = {
        "workbook": label,
        "timings_ms": {
            "read_file": t_read,
            "prepare_request": t_prep,
            "http_roundtrip": t_http,
        },
        "total_ms": total,
        "total_seconds": total / 1000,
        "http_status": status,
        "n_fields": n_fields,
        "success": success,
    }
    return result


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def main():
    out_path = os.path.join(OUTPUT_DIR, "production_vs_benchmark.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        # Redirect console output to file
        old_stdout = sys.stdout
        sys.stdout = f

        print("=" * 70)
        print("  Phase X.5 — Production vs Benchmark Timing Investigation")
        print("=" * 70)

        # ── PART 1: Run benchmark on both workbooks ─────────────────────
        print("\n--- PART 1: In-Process Benchmark ---\n")
        bench_results = {}
        for label, path in [("FormTest", FORMTEST_WORKBOOK), ("Japanese", JAPANESE_WORKBOOK)]:
            print(f"\n  Running benchmark: {label}")
            try:
                result = run_benchmark(path, label)
                bench_results[label] = result
                print(f"  Total: {result['total_seconds']:.1f}s")
            except Exception as e:
                print(f"  FAILED: {e}")
                import traceback
                traceback.print_exc()
                bench_results[label] = {"total_ms": 0, "total_seconds": 0,
                                        "timings_ms": {}, "n_clusters": 0,
                                        "workbook": label,
                                        "errors": {"fatal": str(e)}}

        # ── PART 2: Run production HTTP endpoint ───────────────────────
        print("\n\n--- PART 2: Production HTTP Endpoint ---\n")
        http_results = {}
        for label, path in [("FormTest", FORMTEST_WORKBOOK), ("Japanese", JAPANESE_WORKBOOK)]:
            print(f"\n  Calling production endpoint: {label}")
            try:
                result = run_production_endpoint(path, label)
                http_results[label] = result
                print(f"  Total: {result['total_seconds']:.1f}s, "
                      f"HTTP status: {result.get('http_status')}, "
                      f"fields: {result.get('n_fields')}, "
                      f"success: {result.get('success')}")
            except Exception as e:
                print(f"  FAILED: {e}")
                import traceback
                traceback.print_exc()
                http_results[label] = {"total_ms": 0, "total_seconds": 0,
                                       "timings_ms": {}, "http_status": 0,
                                       "n_fields": 0, "success": False,
                                       "workbook": label}

        # ── PART 3: Comparison Table ───────────────────────────────────
        print("\n\n" + "=" * 70)
        print("  COMPARISON: Benchmark vs Production HTTP")
        print("=" * 70)

        for label in ["FormTest", "Japanese"]:
            bench = bench_results.get(label, {"total_seconds": 0, "timings_ms": {}})
            http = http_results.get(label, {"total_seconds": 0, "timings_ms": {},
                                            "http_status": 0})

            print(f"\n  --- {label} ---")

            # Benchmark internal stages
            print(f"\n  Benchmark (in-process): {bench['total_seconds']:.1f}s")
            print(f"  {'Stage':<30s} {'Time (ms)':>10s}")
            print(f"  {'-'*42}")
            for stage, t in sorted(bench.get("timings_ms", {}).items(),
                                   key=lambda x: x[1], reverse=True):
                pct = t / bench.get("total_ms", 1) * 100 if bench.get("total_ms") else 0
                print(f"  {stage:<30s} {t:>8.0f}ms ({pct:>4.0f}%)")

            # HTTP endpoint stages
            print(f"\n  Production HTTP: {http['total_seconds']:.1f}s")
            if http.get("timings_ms"):
                print(f"  {'Stage':<30s} {'Time (ms)':>10s}")
                print(f"  {'-'*42}")
                for stage, t in sorted(http.get("timings_ms", {}).items(),
                                       key=lambda x: x[1], reverse=True):
                    pct = t / http.get("total_ms", 1) * 100 if http.get("total_ms") else 0
                    print(f"  {stage:<30s} {t:>8.0f}ms ({pct:>4.0f}%)")

            if http.get("http_status"):
                print(f"\n  HTTP Status: {http['http_status']}")
                print(f"  Fields returned: {http.get('n_fields', 0)}")

            # Comparison
            bench_total = bench.get("total_seconds", 0)
            http_total = http.get("total_seconds", 0)
            diff = http_total - bench_total
            print(f"\n  HTTP overhead: {diff:.1f}s ({diff/bench_total*100:.0f}% increase)"

                  if bench_total > 0 else "\n  HTTP overhead: N/A")

        # ── PART 4: Timeout Analysis ───────────────────────────────────
        print("\n\n" + "-" * 70)
        print("  TIMEOUT ANALYSIS")
        print("-" * 70)

        for label in ["FormTest", "Japanese"]:
            bench = bench_results.get(label, {"total_seconds": 0})
            http = http_results.get(label, {"total_seconds": 0})

            print(f"\n  {label}:")
            print(f"    Benchmark (in-process):    {bench['total_seconds']:.1f}s")
            print(f"    Production (via HTTP):     {http['total_seconds']:.1f}s")

            if http.get("http_status") == 0:
                print(f"    HTTP request FAILED or timed out")
            elif http.get("total_seconds", 0) > 100:
                print(f"    !!! EXCEEDS 100s TIMEOUT !!!")
            elif http.get("total_seconds", 0) > 90:
                print(f"    !!! CLOSE TO 100s TIMEOUT !!!")
            else:
                print(f"    OK - under 100s timeout")

            print(f"    HTTP overhead: {http.get('total_seconds', 0) - bench.get('total_seconds', 0):.1f}s")

        # ── PART 5: Root cause identification ──────────────────────────
        print("\n\n" + "-" * 70)
        print("  ROOT CAUSE ANALYSIS")
        print("-" * 70)

        print("""
  The benchmark runs generate_preview() in-process (no HTTP).
  The production endpoint calls generate_preview() via HTTP POST.

  Overhead sources in production path:
    1. ASP.NET saves uploaded file to disk
    2. PythonRenderService reads file from disk
    3. Multipart form data serialization (105KB for Japanese workbook)
    4. HTTP transport (localhost loopback)
    5. FastAPI multipart parsing
    6. FastAPI saves uploaded file to temp dir
    7. JSON serialization/deserialization on both ends
    8. ASP.NET response processing

  The 100s timeout can come from:
    - Default HttpClient.Timeout (100s) if the 5-minute override doesn't apply
    - Kestrel/IIS request body read timeout 
    - Proxy/load balancer timeout
    - Extra overhead from concurrent COM operations
""")

        # Print all results
        print("\n--- RAW DATA ---")
        print(json.dumps({"benchmark": {k: {"total_ms": v["total_ms"],
                                            "total_seconds": v["total_seconds"],
                                            "timings_ms": v["timings_ms"]}
                                         for k, v in bench_results.items()},
                          "production_http": {k: {"total_ms": v["total_ms"],
                                                  "total_seconds": v["total_seconds"],
                                                  "timings_ms": v.get("timings_ms", {}),
                                                  "http_status": v.get("http_status")}
                                               for k, v in http_results.items()}},
                         indent=2))

        sys.stdout = old_stdout

    print(f"Report saved to: {out_path}")


if __name__ == "__main__":
    main()
