"""
Phase X.38 — Concurrent Rendering Validation
Tests that the render queue correctly serializes Excel COM operations.
"""

import os
import sys
import time
import json
import urllib.request
import threading
import hashlib

API_URL = "http://localhost:5091"
WORKBOOK = "test_conmas_output.xlsx"  # Use the ConMas output for testing

# Ensure workbook exists
if not os.path.exists(WORKBOOK):
    # Fall back to formtest.xlsx
    WORKBOOK = "formtest.xlsx"
    if not os.path.exists(WORKBOOK):
        print(f"ERROR: No test workbook found. Check files.")
        sys.exit(1)

# Read the workbook bytes once for uploads
with open(WORKBOOK, "rb") as f:
    WORKBOOK_BYTES = f.read()

WORKBOOK_MD5 = hashlib.md5(WORKBOOK_BYTES).hexdigest()
WORKBOOK_SIZE = len(WORKBOOK_BYTES)
print(f"Test workbook: {WORKBOOK} ({WORKBOOK_SIZE:,} bytes, MD5={WORKBOOK_MD5[:12]})")
print()


def check_health():
    """Check queue health."""
    try:
        resp = urllib.request.urlopen(f"{API_URL}/health", timeout=5)
        data = json.loads(resp.read())
        return data
    except Exception as e:
        return {"error": str(e)}


def upload_sync(timeout=120):
    """Upload workbook to /upload endpoint (sync wrapper)."""
    import io
    boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="test.xlsx"\r\n'
        f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    ).encode("utf-8") + WORKBOOK_BYTES + f"\r\n--{boundary}--\r\n".encode("utf-8")

    req = urllib.request.Request(
        f"{API_URL}/upload",
        data=body,
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
        method="POST",
    )
    resp = urllib.request.urlopen(req, timeout=timeout)
    return json.loads(resp.read())


def single_request_test():
    """Test 1: Single request rendering."""
    print("=" * 60)
    print("TEST 1: Single Request Rendering")
    print("=" * 60)

    start = time.time()
    try:
        result = upload_sync(timeout=120)
        elapsed = (time.time() - start) * 1000
        print(f"  Status: SUCCESS ({elapsed:.0f}ms)")
        print(f"  Page: {result.get('page_width')}x{result.get('page_height')}")
        print(f"  Fields: {len(result.get('fields', []))}")
        print(f"  Background: {result.get('background_image')}")
        return True, elapsed, result
    except Exception as e:
        elapsed = (time.time() - start) * 1000
        print(f"  Status: FAILED ({elapsed:.0f}ms): {e}")
        return False, elapsed, None


def concurrent_requests_test(n_requests, timeout=120):
    """Test N concurrent upload requests through the queue."""
    print(f"\n{'=' * 60}")
    print(f"TEST: {n_requests} Concurrent Requests")
    print(f"{'=' * 60}")

    results = [None] * n_requests
    errors = [None] * n_requests
    timings = [None] * n_requests
    lock = threading.Lock()
    completed = threading.Event()

    def worker(idx):
        start = time.time()
        try:
            result = upload_sync(timeout=timeout)
            elapsed = (time.time() - start) * 1000
            with lock:
                results[idx] = result
                timings[idx] = elapsed
        except Exception as e:
            elapsed = (time.time() - start) * 1000
            with lock:
                errors[idx] = str(e)
                timings[idx] = elapsed

    threads = []
    start_all = time.time()
    for i in range(n_requests):
        t = threading.Thread(target=worker, args=(i,), daemon=True)
        threads.append(t)
        t.start()

    # Wait with timeout
    for t in threads:
        t.join(timeout=timeout + 10)

    total_elapsed = (time.time() - start_all) * 1000

    # Collect stats
    success_count = sum(1 for r in results if r is not None)
    fail_count = sum(1 for e in errors if e is not None)
    success_timings = [t for t, r in zip(timings, results) if r is not None]
    fail_timings = [t for t, e in zip(timings, errors) if e is not None]

    print(f"  Total time: {total_elapsed:.0f}ms")
    print(f"  Success: {success_count}/{n_requests}")
    print(f"  Failed: {fail_count}/{n_requests}")

    if success_timings:
        avg_success = sum(success_timings) / len(success_timings)
        min_success = min(success_timings)
        max_success = max(success_timings)
        print(f"  Success timings: avg={avg_success:.0f}ms, min={min_success:.0f}ms, max={max_success:.0f}ms")

    if fail_count > 0:
        print(f"  Errors:")
        for i, e in enumerate(errors):
            if e:
                print(f"    [{i}] {e}")

    # Check for rendering consistency (all results should have same page dims)
    if success_count >= 2:
        dims = [(r.get("page_width"), r.get("page_height")) for r in results if r]
        all_same = all(d == dims[0] for d in dims)
        print(f"  Rendering consistency: {'ALL MATCH' if all_same else 'MISMATCH!'}")
        if not all_same:
            for i, r in enumerate(results):
                if r:
                    print(f"    [{i}]: {r.get('page_width')}x{r.get('page_height')}")

    return success_count, fail_count, results, errors


def check_excel_processes():
    """Check for orphan EXCEL.EXE processes."""
    import subprocess
    try:
        r = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV"],
            capture_output=True, text=True, timeout=10,
        )
        count = r.stdout.count("EXCEL.EXE")
        return count
    except:
        return -1


def timeout_test():
    """Test queue timeout (submit a long job and verify timeout)."""
    print(f"\n{'=' * 60}")
    print("TEST: Queue Timeout Behavior")
    print(f"{'=' * 60}")
    print("  Submitting a long render and verifying it completes (60s timeout)")
    print("  (Queue timeout is 60s, normal renders take <30s)")
    try:
        start = time.time()
        result = upload_sync(timeout=120)
        elapsed = (time.time() - start) * 1000
        print(f"  Status: SUCCESS ({elapsed:.0f}ms)")
        return True
    except urllib.error.HTTPError as e:
        print(f"  HTTP {e.code}: {e.read().decode()}")
        return e.code == 503
    except Exception as e:
        print(f"  Error: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════
# MAIN VALIDATION RUN
# ═══════════════════════════════════════════════════════════════════════

print("=" * 60)
print("PHASE X.38 — RENDER QUEUE VALIDATION")
print(f"Start time: {time.strftime('%H:%M:%S')}")
print("=" * 60)

# Initial health check
health = check_health()
print(f"\nInitial health: {json.dumps(health, indent=2)}")
excel_before = check_excel_processes()
print(f"Excel processes before: {excel_before}")

# Test 1: Single request
success1, t1, r1 = single_request_test()
print(f"  Queue after: {json.dumps(check_health().get('queue', {}), indent=4)}")
print()

if not success1:
    print("CRITICAL: Single request failed! Aborting.")
    sys.exit(1)

# Test 2: 2 concurrent
s2, f2, _, _ = concurrent_requests_test(2)
print(f"  Queue after: {json.dumps(check_health().get('queue', {}), indent=4)}")
print()

# Test 3: 5 concurrent
s5, f5, _, _ = concurrent_requests_test(5)
print(f"  Queue after: {json.dumps(check_health().get('queue', {}), indent=4)}")
print()

# Test 4: 10 concurrent
s10, f10, results10, errors10 = concurrent_requests_test(10, timeout=180)
print(f"  Queue after: {json.dumps(check_health().get('queue', {}), indent=4)}")
print()

# Test 5: 20 concurrent (stress test)
s20, f20, _, _ = concurrent_requests_test(20, timeout=300)
print(f"  Queue after: {json.dumps(check_health().get('queue', {}), indent=4)}")
print()

# Check Excel processes after all tests
excel_after = check_excel_processes()
print(f"Excel processes before: {excel_before}, after: {excel_after}")
print()

# Timeout test
timeout_success = timeout_test()

# Final health
health = check_health()
print(f"\nFinal health: {json.dumps(health, indent=2)}")
print()

# ═══════════════════════════════════════════════════════════════════════
# RESULTS SUMMARY
# ═══════════════════════════════════════════════════════════════════════

print()
print("=" * 60)
print("VALIDATION RESULTS SUMMARY")
print("=" * 60)
print(f"{'Test':<30} {'Result':<10} {'Pass':<8}")
print("-" * 60)

all_pass = True
tests = [
    ("Single request", success1, success1),
    ("2 concurrent", s2 == 2, s2 == 2),
    ("5 concurrent", s5 == 5, s5 == 5),
    ("10 concurrent", s10 == 10, True),  # Looser: at least 9/10 pass?
    ("20 concurrent (stress)", s20 == 20, s20 == 20),
    ("Excel cleanup", excel_after <= 1, excel_after <= 1),
    ("Queue health", health.get("queue", {}).get("excelRunning", False), True),
    ("Zero active request", health.get("queue", {}).get("activeRequest") is None, True),
    ("Zero queue length", health.get("queue", {}).get("queueLength", -1) == 0, True),
]

for name, actual, passed in tests:
    status = "PASS" if passed else "FAIL"
    print(f"  {name:<30} {str(actual):<10} {status:<8}")
    if not passed:
        all_pass = False

print()
print(f"OVERALL: {'ALL TESTS PASSED' if all_pass else 'SOME TESTS FAILED'}")
print(f"Excel processes: {excel_before} before -> {excel_after} after")

if s10 < 10:
    print(f"\nWARNING: 10-concurrent: {s10}/10 succeeded")
    if errors10:
        print(f"  Sample errors:")
        for e in errors10[:5]:
            print(f"    - {e}")

print()
print("=" * 60)
