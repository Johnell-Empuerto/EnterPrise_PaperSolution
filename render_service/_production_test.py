"""
Phase X.5 — Production vs Benchmark Timing Test

Calls BOTH the ASP.NET production endpoint AND the Python preview endpoint directly.
Measures total round-trip time for each.
Compares against benchmark (in-process) results.
"""

import os
import sys
import time
import io
import json
import requests

JAPANESE = r"C:\Users\MCF-JO~1\Documents\[V3.1_Sample]アンケート用紙.xlsx"
FORMTEST = r"C:\Users\MCF-JO~1\Documents\FormTest - Copy.xlsx"

ASP_NET_URL = "http://127.0.0.1:5090/api/form/upload-preview"
PYTHON_URL = "http://127.0.0.1:5091/upload/preview"

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "timing_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

RESULTS = {}


def test_endpoint(label, url, xlsx_path, is_aspnet=False):
    """Call an endpoint, measure total round-trip time, and report."""
    print(f"\n{'='*60}")
    print(f"  Testing: {label}")
    print(f"  URL: {url}")
    print(f"  File: {os.path.basename(xlsx_path)}")
    print(f"{'='*60}")

    # Read file
    t0 = time.perf_counter()
    with open(xlsx_path, "rb") as f:
        file_bytes = f.read()
    t_read = (time.perf_counter() - t0) * 1000

    # Build request
    t1 = time.perf_counter()
    files = {"file": (os.path.basename(xlsx_path), io.BytesIO(file_bytes),
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    if not is_aspnet:
        params = {"output_dir": OUTPUT_DIR}
    else:
        params = {}
    t_build = (time.perf_counter() - t1) * 1000

    # Send HTTP request with 5-minute timeout
    t2 = time.perf_counter()
    try:
        response = requests.post(url, files=files, params=params, timeout=300)
        t_http = (time.perf_counter() - t2) * 1000
        status = response.status_code
        content_len = len(response.content)
        try:
            data = response.json()
            success = data.get("success", False)
            n_fields = len(data.get("fields", data.get("data", {}).get("fields", data.get("fields", []))))
        except Exception:
            data = {"raw": response.text[:500]}
            success = False
            n_fields = 0
    except requests.Timeout:
        t_http = (time.perf_counter() - t2) * 1000
        status = 0
        content_len = 0
        success = False
        n_fields = 0
        data = {"error": f"TIMEOUT after {t_http:.0f}ms"}
    except Exception as e:
        t_http = (time.perf_counter() - t2) * 1000
        status = -1
        content_len = 0
        success = False
        n_fields = 0
        data = {"error": str(e)}

    total = t_read + t_build + t_http

    print(f"  Read file:     {t_read:.0f}ms")
    print(f"  Build request: {t_build:.0f}ms")
    print(f"  HTTP request:  {t_http:.0f}ms ({t_http/1000:.1f}s)")
    print(f"  Total:         {total:.0f}ms ({total/1000:.1f}s)")
    print(f"  Status:        {status}")
    print(f"  Content:       {content_len} bytes")
    print(f"  Success:       {success}")
    print(f"  Fields:        {n_fields}")

    if t_http > 90000:
        print(f"  *** TOOK > 90s — DANGEROUSLY CLOSE TO 100s TIMEOUT ***")
    if status == 0:
        print(f"  *** HTTP REQUEST TIMED OUT ***")

    return {
        "label": label,
        "url": url,
        "file": os.path.basename(xlsx_path),
        "read_file_ms": t_read,
        "build_request_ms": t_build,
        "http_roundtrip_ms": t_http,
        "total_ms": total,
        "total_seconds": total / 1000,
        "http_status": status,
        "content_bytes": content_len,
        "success": success,
        "n_fields": n_fields,
    }


def main():
    global RESULTS

    report_path = os.path.join(OUTPUT_DIR, "production_test_results.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        old_stdout = sys.stdout
        sys.stdout = f

        print("Phase X.5 — Production vs Benchmark Comparison")
        print("=" * 70)

        # Test all 4 combinations
        combos = [
            ("Python Direct - FormTest", PYTHON_URL, FORMTEST, False),
            ("Python Direct - Japanese", PYTHON_URL, JAPANESE, False),
            ("ASP.NET - FormTest", ASP_NET_URL, FORMTEST, True),
            ("ASP.NET - Japanese", ASP_NET_URL, JAPANESE, True),
        ]

        for label, url, path, is_asp in combos:
            try:
                r = test_endpoint(label, url, path, is_asp)
                RESULTS[label] = r
            except Exception as e:
                print(f"\n  FAILED: {e}")
                RESULTS[label] = {"label": label, "error": str(e), "total_seconds": 0,
                                  "http_roundtrip_ms": 0, "http_status": -1}

        # Summary table
        print("\n\n" + "=" * 70)
        print("  SUMMARY TABLE")
        print("=" * 70)
        print(f"\n  {'Endpoint':<40s} {'Total (s)':>10s} {'HTTP (s)':>10s} {'Status':>8s} {'Fields':>8s}")
        print(f"  {'-'*80}")
        for label in sorted(RESULTS.keys()):
            r = RESULTS[label]
            total = f"{r.get('total_seconds', 0):.1f}"
            http = f"{r.get('http_roundtrip_ms', 0)/1000:.1f}"
            status = str(r.get('http_status', 'ERR'))
            fields = str(r.get('n_fields', 0))
            print(f"  {label:<40s} {total:>10s} {http:>10s} {status:>8s} {fields:>8s}")

        # Timeout analysis
        print("\n\n" + "-" * 70)
        print("  TIMEOUT ANALYSIS")
        print("-" * 70)

        benchmark_data = {
            "Python Direct - FormTest": 34.5,
            "Python Direct - Japanese": 60.4,
        }

        for label, bench_total in benchmark_data.items():
            r = RESULTS.get(label, {})
            http_total = r.get("total_seconds", 0)
            overhead = http_total - bench_total
            print(f"\n  {label}:")
            print(f"    Benchmark:       {bench_total:.1f}s")
            print(f"    With HTTP:       {http_total:.1f}s")
            print(f"    HTTP overhead:   {overhead:.1f}s")

            if http_total > 100:
                print(f"    !!! EXCEEDS 100s TIMEOUT !!!")
            elif http_total > 0:
                print(f"    OK - under 100s")

        # Compare ASP.NET vs Python direct
        print("\n\n" + "-" * 70)
        print("  ASP.NET vs Python Direct (HTTP overhead comparison)")
        print("-" * 70)

        for file_label in ["FormTest", "Japanese"]:
            py_label = f"Python Direct - {file_label}"
            as_label = f"ASP.NET - {file_label}"

            py_r = RESULTS.get(py_label, {})
            as_r = RESULTS.get(as_label, {})

            if py_r.get("total_seconds") and as_r.get("total_seconds"):
                py_total = py_r["total_seconds"]
                as_total = as_r["total_seconds"]
                diff = as_total - py_total
                print(f"\n  {file_label}:")
                print(f"    Python direct:  {py_total:.1f}s")
                print(f"    ASP.NET proxy:  {as_total:.1f}s")
                print(f"    ASP.NET overhead: {diff:.1f}s")
                if diff > 30:
                    print(f"    *** SIGNIFICANT OVERHEAD - investigate ASP.NET timeout setting ***")

        sys.stdout = old_stdout

    # Also print to console
    print(f"\nReport saved to: {report_path}")
    for label in sorted(RESULTS.keys()):
        r = RESULTS[label]
        total = r.get("total_seconds", 0)
        status = r.get("http_status", "ERR")
        print(f"  {label}: {total:.1f}s (HTTP {status})")


if __name__ == "__main__":
    main()
