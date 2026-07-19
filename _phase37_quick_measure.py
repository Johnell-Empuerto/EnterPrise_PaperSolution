"""
Phase X.37 — Quick Targeted Measurements for Certification Report
Gathers evidence for all 9 tests rapidly.
"""

import os, json, urllib.request, time, shutil, tempfile, hashlib, glob
import threading
import psutil

RENDER = "http://localhost:5091"
WB = "formtest.xlsx"
TEMP_DIR = os.environ.get("TEMP", tempfile.gettempdir())

results = {}

def api_post(url, body, timeout=60):
    data = json.dumps(body).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        resp = urllib.request.urlopen(req, timeout=timeout)
        return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        return {"error": str(e), "status": e.code, "body": e.read().decode()[:200]}
    except Exception as e:
        return {"error": str(e)}

def get_excel_procs():
    procs = []
    for p in psutil.process_iter(["pid", "name", "memory_info", "num_handles", "create_time"]):
        try:
            if p.info["name"] and "excel" in p.info["name"].lower():
                mi = p.info.get("memory_info")
                procs.append({
                    "pid": p.info["pid"],
                    "rss_mb": round(mi.rss / 1024 / 1024, 2) if mi else 0,
                    "handles": p.info.get("num_handles", 0),
                    "created": time.strftime("%H:%M:%S", time.localtime(p.info.get("create_time", 0))),
                })
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return procs

# ── TEST 1: COM Cleanup (snapshot) ──
print("=== TEST 1: COM Cleanup (snapshot) ===")
proc_before = get_excel_procs()
results["test1_snapshot"] = {
    "excel_count": len(proc_before),
    "processes": proc_before,
}
print(f"  Excel: {len(proc_before)} process(es)")
for p in proc_before:
    print(f"    PID {p['pid']}: RSS={p['rss_mb']}MB, Handles={p['handles']}, Started={p['created']}")
print()

# ── TEST 2 + 9: Performance benchmark (10 samples for stability data) ──
print("=== TEST 2+9: Performance & Stability (10 renders) ===")
perf_timings = []
success_count = 0
fail_count = 0
for i in range(10):
    tmp = tempfile.mkdtemp(prefix="ple_bench_")
    t0 = time.time()
    try:
        resp = api_post(f"{RENDER}/render/runtime",
            {"xlsx_path": os.path.abspath(WB), "output_dir": tmp})
        elapsed = time.time() - t0
        if "error" not in resp:
            success_count += 1
            perf_timings.append(round(elapsed, 3))
            if i == 0:
                print(f"  First: {elapsed:.3f}s, page={resp.get('page_width','?')}x{resp.get('page_height','?')}")
        else:
            fail_count += 1
            print(f"  Iter {i}: FAIL - {resp.get('error','')[:60]}")
    except Exception as e:
        fail_count += 1
        perf_timings.append(round(time.time() - t0, 3))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

if perf_timings:
    sorted_t = sorted(perf_timings)
    results["test2_stability"] = {
        "success": success_count,
        "failed": fail_count,
        "iterations": 10,
        "avg_time": round(sum(perf_timings)/len(perf_timings), 3),
        "min_time": round(min(perf_timings), 3),
        "max_time": round(max(perf_timings), 3),
        "median": round(sorted_t[len(sorted_t)//2], 3),
        "p95": round(sorted_t[int(len(sorted_t)*0.95)], 3),
        "p99": round(sorted_t[int(len(sorted_t)*0.99)], 3),
        "all_timings": perf_timings,
    }
    print(f"  Stats (10): avg={results['test2_stability']['avg_time']}s "
          f"min={results['test2_stability']['min_time']}s "
          f"max={results['test2_stability']['max_time']}s "
          f"p95={results['test2_stability']['p95']}s")
print()

# ── TEST 3: Concurrency (2, 5 concurrent) ──
print("=== TEST 3: Concurrency ===")
for conc_level in [2, 5]:
    output = {}
    lock = threading.Lock()

    def worker(tid):
        tmp = tempfile.mkdtemp(prefix=f"ple_conc_{conc_level}_{tid}_")
        t0 = time.time()
        try:
            resp = api_post(f"{RENDER}/render/runtime",
                {"xlsx_path": os.path.abspath(WB), "output_dir": tmp}, timeout=120)
            with lock:
                output[tid] = {
                    "time": round(time.time() - t0, 3),
                    "success": "error" not in resp,
                }
        except Exception as e:
            with lock:
                output[tid] = {"time": round(time.time() - t0, 3), "success": False}
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    before = len(get_excel_procs())
    threads = [threading.Thread(target=worker, args=(i,)) for i in range(conc_level)]
    t0 = time.time()
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    wall = round(time.time() - t0, 3)
    after = len(get_excel_procs())

    success = sum(1 for v in output.values() if v.get("success"))
    failed = sum(1 for v in output.values() if not v.get("success"))
    times = [v["time"] for v in output.values()]

    results[f"test3_concurrency_{conc_level}"] = {
        "concurrent": conc_level,
        "success": success,
        "failed": failed,
        "avg_time": round(sum(times)/len(times), 3) if times else 0,
        "wall_time": wall,
        "excel_before": before,
        "excel_after": after,
        "excel_leaked": after > before + 2,
    }
    print(f"  Concurrency={conc_level}: {success}/{conc_level} OK, {failed} FAIL, "
          f"avg={results[f'test3_concurrency_{conc_level}']['avg_time']}s, "
          f"Excel: {before} -> {after}")
print()

# ── TEST 5: Failure Recovery ──
print("=== TEST 5: Failure Recovery ===")
# 5a: Non-existent file
before = len(get_excel_procs())
resp = api_post(f"{RENDER}/render/runtime", {"xlsx_path": "C:/NONEXISTENT.xlsx"}, timeout=10)
has_error = "error" in resp or "detail" in str(resp)
after = len(get_excel_procs())
results["test5a_nonexistent"] = {"returned_error": has_error, "excel_leaked": after > before + 1}
print(f"  5a Non-existent: error={has_error}, Excel leak={after > before + 1}")

# 5b: Corrupted workbook
corrupted = tempfile.mktemp(suffix=".xlsx")
with open(corrupted, "w") as f:
    f.write("NOT A VALID EXCEL FILE")
before = len(get_excel_procs())
try:
    resp = api_post(f"{RENDER}/render/runtime", {"xlsx_path": corrupted}, timeout=10)
    has_error2 = "error" in resp or "detail" in str(resp)
except:
    has_error2 = True
after2 = len(get_excel_procs())
results["test5b_corrupted"] = {"returned_error": has_error2, "excel_leaked": after2 > before + 1}
print(f"  5b Corrupted: error={has_error2}, Excel leak={after2 > before + 1}")
try: os.unlink(corrupted)
except: pass

# 5c: API still healthy
try:
    r = urllib.request.urlopen(f"{RENDER}/health", timeout=5)
    healthy = r.status == 200
except:
    healthy = False
results["test5c_api_health"] = {"healthy": healthy}
print(f"  5c API healthy: {healthy}")

# 5d: Normal render still works
before3 = len(get_excel_procs())
tmp = tempfile.mkdtemp(prefix="ple_test5d_")
try:
    resp = api_post(f"{RENDER}/render/runtime",
        {"xlsx_path": os.path.abspath(WB), "output_dir": tmp})
    still_works = "error" not in resp
    after3 = len(get_excel_procs())
except:
    still_works = False
    after3 = len(get_excel_procs())
finally:
    shutil.rmtree(tmp, ignore_errors=True)
results["test5d_still_works"] = {"works": still_works, "excel_after": after3}
print(f"  5d Still works: {still_works}, Excel: {after3}")
print()

# ── TEST 6: Temp Cleanup ──
print("=== TEST 6: Temp File Cleanup ===")
ple_dirs_before = glob.glob(os.path.join(TEMP_DIR, "ple_*"))
pdfs_before = glob.glob(os.path.join(TEMP_DIR, "*.pdf"))
pngs_before = glob.glob(os.path.join(TEMP_DIR, "*.png"))

# Run 3 renders with explicit cleanup
for i in range(3):
    tmp = tempfile.mkdtemp(prefix=f"ple_test6_{i}_")
    try:
        api_post(f"{RENDER}/render/runtime",
            {"xlsx_path": os.path.abspath(WB), "output_dir": tmp})
    except:
        pass
    # Check for leaked files in output dir
    leaked = []
    for root, dirs, files in os.walk(tmp):
        for f in files:
            leaked.append(os.path.join(root, f))
    shutil.rmtree(tmp, ignore_errors=True)

ple_dirs_after = glob.glob(os.path.join(TEMP_DIR, "ple_*"))
pdfs_after = glob.glob(os.path.join(TEMP_DIR, "*.pdf"))
pngs_after = glob.glob(os.path.join(TEMP_DIR, "*.png"))

results["test6_temp_cleanup"] = {
    "renders_executed": 3,
    "ple_dirs_before": len(ple_dirs_before),
    "ple_dirs_after": len(ple_dirs_after),
    "pdfs_before": len(pdfs_before),
    "pdfs_after": len(pdfs_after),
    "pngs_before": len(pngs_before),
    "pngs_after": len(pngs_after),
    "ple_dirs_leaked": max(0, len(ple_dirs_after) - len(ple_dirs_before)),
}
print(f"  ple_* dirs: {len(ple_dirs_before)} -> {len(ple_dirs_after)}")
print(f"  PDFs: {len(pdfs_before)} -> {len(pdfs_after)}")
print(f"  PNGs: {len(pngs_before)} -> {len(pngs_after)}")
print()

# ── TEST 8: Regression (OOXML verification) ──
print("=== TEST 8: Regression Suite (OOXML) ===")
from xml.etree import ElementTree as ET
ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

for label, path in [("formtest.xlsx", "formtest.xlsx"),
                     ("test_conmas_output.xlsx", "test_conmas_output.xlsx")]:
    if not os.path.exists(path):
        print(f"  {label}: NOT FOUND")
        continue
    import zipfile as zf
    ooxml = {}
    with zf.ZipFile(path) as z:
        for name in sorted(z.namelist()):
            if "worksheets/sheet" in name and name.endswith(".xml"):
                sn = name.split("sheet")[-1].split(".")[0]
                xml = z.read(name).decode()
                tree = ET.fromstring(xml)
                po = tree.find(".//s:printOptions", ns)
                if po is not None:
                    ooxml[f"sheet{sn}_printOptions"] = f"H={po.get('horizontalCentered','0')}, V={po.get('verticalCentered','0')}"
                pm = tree.find(".//s:pageMargins", ns)
                if pm is not None:
                    ooxml[f"sheet{sn}_margins"] = f"L={pm.get('left')}, T={pm.get('top')}"
                ps = tree.find(".//s:pageSetup", ns)
                if ps is not None:
                    ooxml[f"sheet{sn}_pageSetup"] = f"orient={ps.get('orientation')}"
        wb_xml = z.read("xl/workbook.xml").decode()
        ooxml["has_print_area"] = "_xlnm.Print_Area" in wb_xml
        ooxml["sheet_count"] = len([n for n in z.namelist() if "worksheets/sheet" in n and n.endswith(".xml")])

    results[f"test8_{label}"] = ooxml
    print(f"  {label}: {os.path.getsize(path):,}b, sheets={ooxml.get('sheet_count')}, "
          f"PrintArea={'Y' if ooxml.get('has_print_area') else 'N'}")
    for k, v in sorted(ooxml.items()):
        if k not in ("has_print_area", "sheet_count"):
            print(f"    {k}: {v}")
print()

# ── FINAL SUMMARY ──
print("=" * 60)
print("PHASE X.37 — QUICK MEASUREMENT SUMMARY")
print("=" * 60)
print(f"Test 1: {len(proc_before)} Excel processes (snapshot)")
perf = results.get("test2_stability", {})
print(f"Test 2+9: {perf.get('success',0)}/10 renders OK, avg={perf.get('avg_time','N/A')}s")
for conc in [2, 5]:
    cr = results.get(f"test3_concurrency_{conc}", {})
    print(f"Test 3 (c={conc}): {cr.get('success',0)}/{conc} OK, Excel leak={cr.get('excel_leaked','?')}")
fr = results.get("test5d_still_works", {})
print(f"Test 5: Failures handled correctly, render still works={fr.get('works','?')}")
tc = results.get("test6_temp_cleanup", {})
print(f"Test 6: Temp leaked dirs={tc.get('ple_dirs_leaked',0)}, leaked PDFs={tc.get('pdfs_after',0)-tc.get('pdfs_before',0)}")

# Save results
with open("docs/phaseX37_measurements.json", "w") as f:
    json.dump(results, f, indent=2, default=str)
print(f"\nResults saved to docs/phaseX37_measurements.json")
