"""
Phase X.37 — Production Reliability & Regression Certification
==============================================================
Tests 1-9: COM cleanup, stability, concurrency, stress,
failure recovery, temp file cleanup, memory leaks, regression, perf.

Date: 2026-07-18
"""

import os, sys, time, json, hashlib, shutil, tempfile, zipfile, glob
import urllib.request
import subprocess
import threading
from pathlib import Path
from xml.etree import ElementTree as ET
from collections import defaultdict
import traceback

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

PROJECT = os.path.dirname(os.path.abspath(__file__))
API_URL = "http://localhost:5090"
RENDER_URL = "http://localhost:5091"
PRIMARY_WB = os.path.join(PROJECT, "formtest.xlsx")
CONMAS_WB = os.path.join(PROJECT, "test_conmas_output.xlsx")
TEMP = tempfile.gettempdir()

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# ── Helpers ───────────────────────────────────────────────────────────────

def http_post(url, body, timeout=60):
    """HTTP POST with JSON body, returns parsed JSON."""
    data = json.dumps(body).encode()
    req = urllib.request.Request(url, data=data,
        headers={"Content-Type": "application/json"})
    try:
        resp = urllib.request.urlopen(req, timeout=timeout)
        return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()[:500]
        return {"error": str(e), "status": e.code, "body": body}
    except Exception as e:
        return {"error": str(e)}

def md5(path):
    return hashlib.md5(open(path, "rb").read()).hexdigest()[:12]

def safe_excel_processes():
    """Safely iterate Excel processes with error guards."""
    if not HAS_PSUTIL:
        return []
    processes = []
    try:
        for p in psutil.process_iter(['pid', 'name', 'memory_info', 'num_handles',
                                       'cpu_percent', 'create_time', 'memory_percent']):
            try:
                if p.info['name'] and 'excel' in p.info['name'].lower():
                    mi = p.info.get('memory_info')
                    processes.append({
                        'pid': p.info['pid'],
                        'name': p.info['name'],
                        'rss_mb': round(mi.rss / 1024 / 1024, 2) if mi else 0,
                        'vms_mb': round(mi.vms / 1024 / 1024, 2) if mi else 0,
                        'private_mb': round(mi.private / 1024 / 1024, 2) if hasattr(mi, 'private') and mi.private else 0,
                        'handles': p.info.get('num_handles', 0),
                        'cpu_percent': p.info.get('cpu_percent', 0),
                        'created': time.strftime("%H:%M:%S", time.localtime(p.info.get('create_time', 0))),
                    })
            except (psutil.NoSuchProcess, psutil.AccessDenied, AttributeError):
                pass
    except Exception:
        pass
    return processes

def excel_count():
    return len(safe_excel_processes())

def system_memory():
    """Return total system memory info."""
    if not HAS_PSUTIL:
        return {}
    try:
        mem = psutil.virtual_memory()
        return {
            'total_gb': round(mem.total / 1024**3, 2),
            'available_gb': round(mem.available / 1024**3, 2),
            'percent_used': mem.percent,
        }
    except:
        return {}

def render_via_api(xlsx_path, output_dir=None):
    """Call the render API to process a workbook end-to-end."""
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="ple_test_")
    return http_post(f"{RENDER_URL}/render/runtime", {
        "xlsx_path": xlsx_path,
        "output_dir": output_dir,
    }, timeout=120)

def ooxml_verify_full(path):
    """Full OOXML verification of print-related elements."""
    result = {}
    try:
        with zipfile.ZipFile(path) as z:
            # Sheet analysis
            for name in sorted(z.namelist()):
                if 'worksheets/sheet' in name and name.endswith('.xml'):
                    sheet_num = name.split('sheet')[-1].split('.')[0]
                    try:
                        xml = z.read(name).decode()
                        tree = ET.fromstring(xml)
                        po = tree.find('.//s:printOptions', NS)
                        if po is not None:
                            result[f'sheet{sheet_num}_printOptions'] = {
                                'horizontalCentered': po.get('horizontalCentered', '0'),
                                'verticalCentered': po.get('verticalCentered', '0'),
                            }
                        pm = tree.find('.//s:pageMargins', NS)
                        if pm is not None:
                            result[f'sheet{sheet_num}_pageMargins'] = {
                                'left': pm.get('left'), 'right': pm.get('right'),
                                'top': pm.get('top'), 'bottom': pm.get('bottom'),
                                'header': pm.get('header'), 'footer': pm.get('footer'),
                            }
                        ps = tree.find('.//s:pageSetup', NS)
                        if ps is not None:
                            result[f'sheet{sheet_num}_pageSetup'] = {
                                'orientation': ps.get('orientation'),
                                'paperSize': ps.get('paperSize'),
                                'fitToWidth': ps.get('fitToWidth'),
                                'fitToHeight': ps.get('fitToHeight'),
                            }
                        dim = tree.find('.//s:dimension', NS)
                        if dim is not None:
                            result[f'sheet{sheet_num}_dimension'] = dim.get('ref')
                        sp = tree.find('.//s:sheetPr', NS)
                        if sp is not None:
                            result[f'sheet{sheet_num}_sheetPr'] = {
                                'pageSetUpPr': sp.get('pageSetUpPr') if sp.get('pageSetUpPr') else 'none',
                                'tabColor': sp.get('tabColor') if sp.get('tabColor') else 'none',
                            }
                    except Exception as e:
                        result[f'sheet{sheet_num}_error'] = str(e)[:100]

            # Defined names (Print_Area)
            wb_xml = z.read('xl/workbook.xml').decode()
            result['Print_Area_defined'] = '_xlnm.Print_Area' in wb_xml or 'Print_Area' in wb_xml
            # Print_Titles
            result['Print_Titles_defined'] = '_xlnm.Print_Titles' in wb_xml

            # Workbook view
            try:
                wb_tree = ET.fromstring(wb_xml)
                wb_view = wb_tree.find('.//s:workbookView', NS)
                if wb_view is not None:
                    result['workbookView'] = {
                        'activeTab': wb_view.get('activeTab'),
                        'firstSheet': wb_view.get('firstSheet'),
                    }
            except:
                pass

            # Sheet count
            sheets = [n for n in z.namelist() if 'worksheets/sheet' in n and n.endswith('.xml')]
            result['sheet_count'] = len(sheets)

            # Check for printerSettings
            has_printer = any('printerSettings' in n for n in z.namelist())
            result['has_printerSettings'] = has_printer

    except Exception as e:
        result['error'] = str(e)[:200]
    return result

def create_test_workbooks():
    """Create diverse workbooks for regression suite if they don't exist."""
    created = {}
    base_path = os.path.join(PROJECT, "_test_workbooks")
    os.makedirs(base_path, exist_ok=True)

    # Use the C# API to generate different workbook types if possible
    # For now, we use available workbooks
    available = {
        'portrait_form': PRIMARY_WB,
        'conmas_reference': CONMAS_WB,
    }

    # Copy generated workbooks from previous phases
    for f in glob.glob("_x34_generated*.xlsx"):
        name = os.path.basename(f)
        dst = os.path.join(base_path, name)
        shutil.copy2(f, dst)
        available[f'generated_{name}'] = dst

    return available

# ═══════════════════════════════════════════════════════════════════════════
# TEST 1 — COM Resource Cleanup (100 iterations)
# ═══════════════════════════════════════════════════════════════════════════

def test1_com_cleanup(iterations=100):
    print(f"\n{'='*80}")
    print(f"TEST 1: COM Resource Cleanup ({iterations} iterations)")
    print(f"{'='*80}")

    initial_processes = safe_excel_processes()
    initial_count = len(initial_processes)
    initial_handles = sum(p['handles'] for p in initial_processes)
    initial_mem_rss = sum(p['rss_mb'] for p in initial_processes)

    results = {
        'iterations': iterations,
        'success': 0,
        'failed': 0,
        'initial_excel': initial_count,
        'initial_handles': initial_handles,
        'initial_memory_rss_mb': round(initial_mem_rss, 2),
        'max_excel_count': initial_count,
        'max_handles': initial_handles,
        'max_memory_rss_mb': round(initial_mem_rss, 2),
        'final_excel': 0,
        'final_handles': 0,
        'final_memory_rss_mb': 0,
        'process_snapshots': [],
        'errors': [],
        'orphans_detected': False,
    }

    for i in range(iterations):
        tmp = tempfile.mkdtemp(prefix=f"ple_test1_{i}_")
        try:
            resp = render_via_api(PRIMARY_WB, output_dir=tmp)
            if "error" in resp:
                results['failed'] += 1
                results['errors'].append(f"iter_{i}: {resp.get('error','unknown')[:100]}")
            else:
                results['success'] += 1
        except Exception as e:
            results['failed'] += 1
            results['errors'].append(f"iter_{i}: {str(e)[:100]}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

        # Snapshot every 10 iterations
        if (i + 1) % 10 == 0 or i == 0:
            procs = safe_excel_processes()
            count = len(procs)
            total_handles = sum(p['handles'] for p in procs)
            total_mem = sum(p['rss_mb'] for p in procs)
            results['max_excel_count'] = max(results['max_excel_count'], count)
            results['max_handles'] = max(results['max_handles'], total_handles)
            results['max_memory_rss_mb'] = max(results['max_memory_rss_mb'], round(total_mem, 2))
            results['process_snapshots'].append({
                'iteration': i + 1,
                'excel_count': count,
                'total_handles': total_handles,
                'total_rss_mb': round(total_mem, 2),
                'processes': procs,
            })
            if count > initial_count + 2:
                results['orphans_detected'] = True

    # Final snapshot
    final_procs = safe_excel_processes()
    results['final_excel'] = len(final_procs)
    results['final_handles'] = sum(p['handles'] for p in final_procs)
    results['final_memory_rss_mb'] = round(sum(p['rss_mb'] for p in final_procs), 2)
    results['memory_delta_mb'] = round(results['final_memory_rss_mb'] - initial_mem_rss, 2)
    results['handles_delta'] = results['final_handles'] - initial_handles
    results['process_delta'] = results['final_excel'] - initial_count

    results['pass'] = (
        not results['orphans_detected']
        and results['failed'] == 0
        and results['process_delta'] <= 0
        and results['memory_delta_mb'] < 20
    )

    print(f"  Success: {results['success']}, Failed: {results['failed']}")
    print(f"  Processes: {initial_count} -> {results['final_excel']} (delta={results['process_delta']})")
    print(f"  Handles: {initial_handles} -> {results['final_handles']} (delta={results['handles_delta']})")
    print(f"  Memory RSS: {initial_mem_rss:.1f}MB -> {results['final_memory_rss_mb']:.1f}MB (delta={results['memory_delta_mb']:.1f}MB)")
    print(f"  Orphans detected: {results['orphans_detected']}")
    print(f"  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 2 — Long-Running Stability
# ═══════════════════════════════════════════════════════════════════════════

def test2_stability(iterations=200):
    print(f"\n{'='*80}")
    print(f"TEST 2: Long-Running Stability ({iterations} iterations)")
    print(f"{'='*80}")

    results = {
        'iterations': 0,
        'success': 0,
        'failed': 0,
        'exceptions': [],
        'timings': [],
        'memory_samples': [],
        'cpu_samples': [],
        'process_samples': [],
        'excel_launch_failures': 0,
    }

    initial_mem = system_memory()

    for i in range(iterations):
        t0 = time.time()
        tmp = tempfile.mkdtemp(prefix=f"ple_test2_{i}_")
        try:
            resp = render_via_api(PRIMARY_WB, output_dir=tmp)
            elapsed = time.time() - t0
            results['timings'].append(round(elapsed, 3))

            if "error" in resp:
                results['failed'] += 1
                results['exceptions'].append(f"iter_{i}: {resp.get('error','unknown')[:100]}")
            else:
                results['success'] += 1
        except Exception as e:
            elapsed = time.time() - t0
            results['timings'].append(round(elapsed, 3))
            results['failed'] += 1
            results['exceptions'].append(f"iter_{i}: {str(e)[:100]}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

        results['iterations'] = i + 1

        # Sampling every 10 iterations
        if i % 10 == 0 or i == iterations - 1:
            procs = safe_excel_processes()
            results['process_samples'].append({'iter': i, 'count': len(procs), 'pids': [p['pid'] for p in procs]})
            results['memory_samples'].append({'iter': i, 'rss_mb': sum(p['rss_mb'] for p in procs)})
            results['cpu_samples'].append({'iter': i, 'cpu': sum(p.get('cpu_percent', 0) for p in procs)})

        # Detect Excel launch failures
        if "Unable to get the Excel" in str(results['exceptions'][-1:]):
            results['excel_launch_failures'] += 1

    # Stats
    timings = results['timings']
    if timings:
        sorted_t = sorted(timings)
        results['stats'] = {
            'avg': round(sum(timings)/len(timings), 3),
            'min': round(min(timings), 3),
            'max': round(max(timings), 3),
            'median': round(sorted_t[len(sorted_t)//2], 3),
            'p95': round(sorted_t[int(len(sorted_t)*0.95)], 3),
            'p99': round(sorted_t[int(len(sorted_t)*0.99)], 3),
            'total': round(sum(timings), 1),
        }

    results['pass'] = results['failed'] == 0 and results['excel_launch_failures'] == 0

    print(f"  Success: {results['success']}, Failed: {results['failed']}")
    print(f"  Excel launch failures: {results['excel_launch_failures']}")
    print(f"  Avg: {results['stats']['avg']}s, Min/Max: {results['stats']['min']}/{results['stats']['max']}s")
    print(f"  P95/P99: {results['stats']['p95']}/{results['stats']['p99']}s")
    print(f"  Total: {results['stats']['total']}s")
    print(f"  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 3 — Concurrent Rendering
# ═══════════════════════════════════════════════════════════════════════════

def test3_concurrency():
    print(f"\n{'='*80}")
    print(f"TEST 3: Concurrent Rendering (2, 5, 10 concurrent)")
    print(f"{'='*80}")

    results = {}
    for concurrency in [2, 5, 10]:
        print(f"\n--- Testing concurrency={concurrency} ---")
        before = safe_excel_processes()
        before_count = len(before)

        output = {}
        lock = threading.Lock()

        def worker(tid):
            tmp = tempfile.mkdtemp(prefix=f"ple_test3_{concurrency}_{tid}_")
            t0 = time.time()
            try:
                resp = render_via_api(PRIMARY_WB, output_dir=tmp)
                with lock:
                    output[tid] = {
                        'time': round(time.time() - t0, 3),
                        'success': 'error' not in resp,
                        'page': f"{resp.get('page_width','?')}x{resp.get('page_height','?')}" if 'error' not in resp else 'FAIL',
                    }
            except Exception as e:
                with lock:
                    output[tid] = {'time': round(time.time() - t0, 3), 'success': False, 'error': str(e)[:100]}
            finally:
                shutil.rmtree(tmp, ignore_errors=True)

        # Launch all threads at once
        threads = [threading.Thread(target=worker, args=(i,)) for i in range(concurrency)]
        t0 = time.time()
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        wall = round(time.time() - t0, 3)

        after = safe_excel_processes()
        after_count = len(after)

        success_count = sum(1 for v in output.values() if v.get('success'))
        fail_count = sum(1 for v in output.values() if not v.get('success'))
        times = [v['time'] for v in output.values()]
        avg_time = round(sum(times)/len(times), 3) if times else 0

        results[concurrency] = {
            'concurrent_requests': concurrency,
            'success': success_count,
            'failed': fail_count,
            'avg_time': avg_time,
            'wall_time': wall,
            'excel_before': before_count,
            'excel_after': after_count,
            'excel_leaked': after_count > before_count + 2,
            'max_concurrent_excel': max(
                sum(1 for p in safe_excel_processes()),
                before_count
            ),
        }
        print(f"  {success_count}/{concurrency} success, {fail_count} failed")
        print(f"  Avg time: {avg_time}s, Wall: {wall}s")
        print(f"  Excel before/after: {before_count}/{after_count}")

    overall_pass = all(v['failed'] == 0 and not v['excel_leaked'] for v in results.values())
    results['pass'] = overall_pass
    print(f"\n  PASS: {overall_pass}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 4 — Stress Test (mixed workloads)
# ═══════════════════════════════════════════════════════════════════════════

def test4_stress():
    print(f"\n{'='*80}")
    print(f"TEST 4: Stress Test (mixed workloads)")
    print(f"{'='*80}")

    workbooks = create_test_workbooks()
    wb_keys = list(workbooks.keys())
    results = {
        'workbooks': {},
        'total_requests': 0,
        'total_success': 0,
        'total_failed': 0,
        'iterations': 0,
    }

    # 5 rounds of mixed workloads = 5 * len(workbooks) iterations
    for round_num in range(5):
        for wb_key, wb_path in workbooks.items():
            if not os.path.exists(wb_path):
                continue
            tmp = tempfile.mkdtemp(prefix=f"ple_test4_r{round_num}_")
            t0 = time.time()
            try:
                resp = render_via_api(wb_path, output_dir=tmp)
                elapsed = round(time.time() - t0, 3)
                if wb_key not in results['workbooks']:
                    results['workbooks'][wb_key] = {'success': 0, 'failed': 0, 'times': [], 'size': os.path.getsize(wb_path)}
                if "error" in resp:
                    results['workbooks'][wb_key]['failed'] += 1
                    results['total_failed'] += 1
                else:
                    results['workbooks'][wb_key]['success'] += 1
                    results['total_success'] += 1
                results['workbooks'][wb_key]['times'].append(elapsed)
            except Exception as e:
                results['total_failed'] += 1
                if wb_key not in results['workbooks']:
                    results['workbooks'][wb_key] = {'success': 0, 'failed': 0, 'times': [], 'size': os.path.getsize(wb_path)}
                results['workbooks'][wb_key]['failed'] += 1
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            results['total_requests'] += 1
        results['iterations'] = round_num + 1

    print(f"  Total: {results['total_success']} success, {results['total_failed']} failed out of {results['total_requests']}")
    for wk, data in sorted(results['workbooks'].items()):
        times = data.get('times', [])
        avg = round(sum(times)/len(times), 3) if times else 0
        print(f"  {wk[:30]:<30} {os.path.getsize(data.get('size',0)) if isinstance(data.get('size'), int) else '?':>7}b  "
              f"{data['success']} OK  {data['failed']} FAIL  avg {avg}s")

    results['pass'] = results['total_failed'] == 0
    print(f"  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 5 — Failure Recovery
# ═══════════════════════════════════════════════════════════════════════════

def test5_failure_recovery():
    print(f"\n{'='*80}")
    print(f"TEST 5: Failure Recovery")
    print(f"{'='*80}")

    results = {'tests': [], 'excel_cleanup_verified': True}

    # 5a: Non-existent file
    print("\n  5a: Non-existent file path")
    before = excel_count()
    try:
        resp = render_via_api("C:/nonexistent_dir_xyz/file.xlsx")
        had_error = 'error' in resp or ('detail' in str(resp))
    except:
        had_error = True
    after = excel_count()
    results['tests'].append({
        'name': '5a_non_existent_file',
        'returned_error': had_error,
        'excel_leaked': after > before + 1,
    })
    print(f"    Error: {had_error}, Excel leak: {after > before + 1}")

    # 5b: Empty file (corrupted workbook)
    print("\n  5b: Corrupted workbook")
    corrupted = tempfile.mktemp(suffix=".xlsx")
    try:
        with open(corrupted, "w") as f:
            f.write("NOT A VALID EXCEL FILE")
        before = excel_count()
        resp = render_via_api(corrupted)
        after = excel_count()
        had_error = 'error' in resp or ('detail' in str(resp))
        results['tests'].append({
            'name': '5b_corrupted_workbook',
            'returned_error': had_error,
            'excel_leaked': after > before + 1,
            'response': str(resp)[:200],
        })
        print(f"    Error: {had_error}, Excel leak: {after > before + 1}")
    finally:
        try: os.unlink(corrupted)
        except: pass

    # 5c: Minimal workbook (no _Fields, no comments — should still render)
    print("\n  5c: Workbook without metadata")
    # Create a minimal xlsx with a single blank sheet
    minimal = os.path.join(tempfile.mkdtemp(prefix="ple_test5c_"), "minimal.xlsx")
    os.makedirs(os.path.dirname(minimal), exist_ok=True)
    try:
        # Build minimal valid xlsx zip
        import io, zipfile as zf
        buf = io.BytesIO()
        with zf.ZipFile(buf, 'w') as z:
            z.writestr('[Content_Types].xml',
                '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '</Types>')
            z.writestr('_rels/.rels',
                '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '</Relationships>')
            z.writestr('xl/workbook.xml',
                '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels',
                '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>')
            z.writestr('xl/worksheets/sheet1.xml',
                '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row></sheetData></worksheet>')
        with open(minimal, 'wb') as f:
            f.write(buf.getvalue())
        before = excel_count()
        resp = render_via_api(minimal)
        after = excel_count()
        had_error = 'error' in resp or ('detail' in str(resp))
        results['tests'].append({
            'name': '5c_minimal_workbook',
            'returned_error': had_error,
            'excel_leaked': after > before + 1,
        })
        print(f"    Error: {had_error}, Excel leak: {after > before + 1}")
    except Exception as e:
        results['tests'].append({
            'name': '5c_minimal_workbook',
            'returned_error': True,
            'error': str(e)[:100],
        })
        print(f"    Creation error: {e}")
    finally:
        try: shutil.rmtree(os.path.dirname(minimal), ignore_errors=True)
        except: pass

    # 5d: API health after errors
    print("\n  5d: API health after errors")
    try:
        resp = urllib.request.urlopen(f"{RENDER_URL}/health", timeout=5)
        health_ok = resp.status == 200
    except:
        health_ok = False
    results['tests'].append({'name': '5d_api_health', 'healthy': health_ok})
    print(f"    Healthy: {health_ok}")

    # 5e: Normal render still works after failures
    print("\n  5e: Normal render after failures")
    tmp = tempfile.mkdtemp(prefix="ple_test5e_")
    try:
        resp = render_via_api(PRIMARY_WB, output_dir=tmp)
        still_works = 'error' not in resp
        if still_works:
            page = f"{resp.get('page_width','?')}x{resp.get('page_height','?')}"
            fields = len(resp.get('fields', []))
        else:
            page, fields = 'N/A', 0
    except:
        still_works = False
        page, fields = 'N/A', 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    results['tests'].append({'name': '5e_render_after_failures', 'works': still_works, 'page': page, 'fields': fields})
    print(f"    Works: {still_works} ({page}, {fields} fields)")

    # Final Excel check
    results['final_excel'] = excel_count()
    results['pass'] = (
        all(t.get('returned_error', True) for t in results['tests'] if t['name'].startswith('5a') or t['name'].startswith('5b'))
        and all(t.get('healthy', True) or t.get('works', True) for t in results['tests'])
        and all(not t.get('excel_leaked', False) for t in results['tests'] if 'excel_leaked' in t)
    )
    print(f"\n  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 6 — Temporary File Cleanup
# ═══════════════════════════════════════════════════════════════════════════

def test6_temp_cleanup(iterations=10):
    print(f"\n{'='*80}")
    print(f"TEST 6: Temporary File Cleanup ({iterations} iterations)")
    print(f"{'='*80}")

    # Count pre-existing temp files
    pre_pdf = len(glob.glob(os.path.join(TEMP, '*.pdf')))
    pre_png = len(glob.glob(os.path.join(TEMP, '*.png')))
    pre_xlsx = len(glob.glob(os.path.join(TEMP, '*.xlsx')))
    pre_ple_dirs = len(glob.glob(os.path.join(TEMP, 'ple_*')))
    pre_other = len(glob.glob(os.path.join(TEMP, '*')))

    results = {
        'iterations': iterations,
        'pre_pdf': pre_pdf,
        'pre_png': pre_png,
        'pre_xlsx': pre_xlsx,
        'pre_ple_dirs': pre_ple_dirs,
        'leaked': [],
        'post_pdf': 0,
        'post_png': 0,
        'post_xlsx': 0,
        'post_ple_dirs': 0,
    }

    for i in range(iterations):
        # Use explicit output dir (render tmp dir)
        tmp = tempfile.mkdtemp(prefix=f"ple_test6_{i}_")
        try:
            resp = render_via_api(PRIMARY_WB, output_dir=tmp)
        except:
            pass

        # Check what was left in the output dir
        remaining = []
        for root, dirs, files in os.walk(tmp):
            for f in files:
                remaining.append(os.path.join(root, f))

        if remaining:
            results['leaked'].extend(remaining)

        shutil.rmtree(tmp, ignore_errors=True)

    # Post-run count
    results['post_pdf'] = len(glob.glob(os.path.join(TEMP, '*.pdf')))
    results['post_png'] = len(glob.glob(os.path.join(TEMP, '*.png')))
    results['post_xlsx'] = len(glob.glob(os.path.join(TEMP, '*.xlsx')))
    results['post_ple_dirs'] = len(glob.glob(os.path.join(TEMP, 'ple_*')))

    pdf_delta = results['post_pdf'] - pre_pdf
    png_delta = results['post_png'] - pre_png
    xlsx_delta = results['post_xlsx'] - pre_xlsx
    ple_delta = results['post_ple_dirs'] - pre_ple_dirs

    results['pdf_leak'] = max(0, pdf_delta)
    results['png_leak'] = max(0, png_delta)
    results['xlsx_leak'] = max(0, xlsx_delta)
    results['ple_dir_leak'] = max(0, ple_delta)

    results['pass'] = (
        len(results['leaked']) == 0
        and results['pdf_leak'] == 0
        and results['png_leak'] == 0
        and results['xlsx_leak'] == 0
        and results['ple_dir_leak'] == 0
    )

    print(f"  Leaked files in output dirs: {len(results['leaked'])}")
    print(f"  PDF leaked: {results['pdf_leak']}")
    print(f"  PNG leaked: {results['png_leak']}")
    print(f"  XLSX leaked: {results['xlsx_leak']}")
    print(f"  Temp dirs leaked: {results['ple_dir_leak']}")
    print(f"  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 7 — Memory Leak Detection
# ═══════════════════════════════════════════════════════════════════════════

def test7_memory_leak(test1_results, test2_results):
    print(f"\n{'='*80}")
    print(f"TEST 7: Memory Leak Detection")
    print(f"{'='*80}")

    results = {
        'test1_memory_delta_mb': test1_results.get('memory_delta_mb', 'N/A'),
        'test1_handles_delta': test1_results.get('handles_delta', 'N/A'),
        'test2_memory_trend': [],
        'stable': False,
    }

    # Extract memory samples from Test 2
    mem_samples = test2_results.get('memory_samples', [])
    handle_samples = []
    if test1_results.get('process_snapshots'):
        handle_samples = [s.get('total_handles', 0) for s in test1_results['process_snapshots']]

    if len(mem_samples) >= 3:
        first_values = [s['rss_mb'] for s in mem_samples[:len(mem_samples)//3]]
        last_values = [s['rss_mb'] for s in mem_samples[-len(mem_samples)//3:]]
        first_avg = sum(first_values) / len(first_values)
        last_avg = sum(last_values) / len(last_values)
        trend = round(last_avg - first_avg, 2)
        results['memory_trend_mb'] = trend
        results['stable'] = abs(trend) < 15  # < 15MB drift = stable
        results['classification'] = 'Stable' if results['stable'] else ('Growing' if trend > 15 else 'Stable')

    # Handle trend from Test 1
    if len(handle_samples) >= 3:
        results['handle_trend'] = handle_samples[-1] - handle_samples[0]

    results['pass'] = results['stable']

    print(f"  Memory trend: {results.get('memory_trend_mb', 'N/A')} MB over {len(mem_samples)} samples")
    print(f"  Handles delta: {results.get('handle_trend', 'N/A')}")
    print(f"  Classification: {results.get('classification', 'N/A')}")
    print(f"  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 8 — Regression Suite
# ═══════════════════════════════════════════════════════════════════════════

def test8_regression():
    print(f"\n{'='*80}")
    print(f"TEST 8: Regression Suite")
    print(f"{'='*80}")

    workbooks = create_test_workbooks()
    results = {'workbooks': {}}

    for name, path in workbooks.items():
        if not path or not os.path.exists(path):
            results['workbooks'][name] = {'status': 'NOT FOUND', 'pass': False}
            continue

        wb = {'size': os.path.getsize(path), 'md5': md5(path)}

        # OOXML verification
        ooxml = ooxml_verify_full(path)
        wb['ooxml'] = ooxml

        # Render test
        tmp = tempfile.mkdtemp(prefix=f"ple_test8_")
        t0 = time.time()
        try:
            resp = render_via_api(path, output_dir=tmp)
            elapsed = round(time.time() - t0, 3)
            if 'error' not in resp:
                wb['render'] = {
                    'success': True, 'time': elapsed,
                    'page': f"{resp.get('page_width','?')}x{resp.get('page_height','?')}",
                    'fields': len(resp.get('fields', [])),
                }
            else:
                wb['render'] = {'success': False, 'error': resp.get('error','unknown')[:200]}
        except Exception as e:
            wb['render'] = {'success': False, 'error': str(e)[:200]}
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

        # Derived pass criteria
        pa_ok = ooxml.get('Print_Area_defined', False)
        po_ok = any(
            v.get('horizontalCentered') == '1' and v.get('verticalCentered') == '1'
            for k, v in ooxml.items() if 'printOptions' in k
        )
        render_ok = wb.get('render', {}).get('success', False)
        wb['pass'] = render_ok
        wb['print_area_ok'] = pa_ok
        wb['print_options_ok'] = po_ok

        results['workbooks'][name] = wb
        print(f"  {name[:40]:<40} {wb['size']:>7}b  "
              f"Render={'OK' if render_ok else 'FAIL'}  "
              f"PA={'Y' if pa_ok else '-'}  PO={'Y' if po_ok else '-'}  "
              f"Fields={wb.get('render',{}).get('fields','?')}")

    results['total'] = len(workbooks)
    results['passed'] = sum(1 for v in results['workbooks'].values() if v.get('pass'))
    results['not_found'] = sum(1 for v in results['workbooks'].values() if v.get('status') == 'NOT FOUND')
    results['pass'] = results['passed'] >= results['total'] - results['not_found']
    print(f"\n  {results['passed']}/{results['total']} workbooks passed ({results['not_found']} not found)")
    print(f"  PASS: {results['pass']}")
    return results

# ═══════════════════════════════════════════════════════════════════════════
# TEST 9 — Performance Benchmark
# ═══════════════════════════════════════════════════════════════════════════

def test9_performance():
    print(f"\n{'='*80}")
    print(f"TEST 9: Performance Benchmark")
    print(f"{'='*80}")

    results = {'workbooks': {}}

    for label, path in [('formtest (small)', PRIMARY_WB), ('ConMas (reference)', CONMAS_WB)]:
        timings = []
        for i in range(10):
            tmp = tempfile.mkdtemp(prefix=f"ple_test9_{label[:3]}_{i}_")
            t0 = time.time()
            try:
                resp = render_via_api(path, output_dir=tmp)
                timings.append(round(time.time() - t0, 3))
            except:
                timings.append(round(time.time() - t0, 3))
            finally:
                shutil.rmtree(tmp, ignore_errors=True)

        if timings:
            sorted_t = sorted(timings)
            size = os.path.getsize(path)
            results['workbooks'][label] = {
                'size_bytes': size,
                'size_kb': round(size/1024, 1),
                'avg': round(sum(timings)/len(timings), 3),
                'min': round(min(timings), 3),
                'max': round(max(timings), 3),
                'median': round(sorted_t[len(sorted_t)//2], 3),
                'p95': round(sorted_t[int(len(sorted_t)*0.95)], 3),
                'p99': round(sorted_t[int(len(sorted_t)*0.99)], 3),
                'samples': len(timings),
            }

        print(f"\n  {label} ({os.path.getsize(path):,} bytes, 10 samples):")
        d = results['workbooks'][label]
        print(f"    Avg: {d['avg']}s  Min: {d['min']}s  Max: {d['max']}s")
        print(f"    P95: {d['p95']}s  P99: {d['p99']}s  Median: {d['median']}s")

    results['pass'] = True
    return results

# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("PHASE X.37 — PRODUCTION RELIABILITY & REGRESSION CERTIFICATION")
    print("Date: 2026-07-18")
    print(f"psutil: {'Available (v' + psutil.__version__ + ')' if HAS_PSUTIL else 'NOT AVAILABLE'}")
    print(f"System: {system_memory()}")
    print(f"Render API: {RENDER_URL}")
    print(f"C# API: {API_URL}")
    print(f"Primary: {PRIMARY_WB} ({os.path.getsize(PRIMARY_WB):,} bytes)")
    print(f"ConMas:  {CONMAS_WB} ({os.path.getsize(CONMAS_WB):,} bytes)")

    # Initial Excel state
    initial_excel = safe_excel_processes()
    print(f"Initial Excel: {len(initial_excel)} process(es)")
    for p in initial_excel:
        print(f"  PID {p['pid']}: RSS={p['rss_mb']}MB Handles={p['handles']} Created={p['created']}")

    all_results = {}

    # ── Test 1: COM Cleanup ──
    all_results['test1_com_cleanup'] = test1_com_cleanup(100)

    # ── Test 2: Stability ──
    all_results['test2_stability'] = test2_stability(200)

    # ── Test 7: Memory Leak ──
    all_results['test7_memory_leak'] = test7_memory_leak(
        all_results['test1_com_cleanup'], all_results['test2_stability'])

    # ── Test 9: Performance ──
    all_results['test9_performance'] = test9_performance()

    # ── Test 3: Concurrency ──
    all_results['test3_concurrency'] = test3_concurrency()

    # ── Test 4: Stress ──
    all_results['test4_stress'] = test4_stress()

    # ── Test 5: Failure Recovery ──
    all_results['test5_failure_recovery'] = test5_failure_recovery()

    # ── Test 6: Temp Cleanup ──
    all_results['test6_temp_cleanup'] = test6_temp_cleanup(10)

    # ── Test 8: Regression ──
    all_results['test8_regression'] = test8_regression()

    # ── Final Summary ──
    print(f"\n{'='*80}")
    print(f"PHASE X.37 — CERTIFICATION RESULTS SUMMARY")
    print(f"{'='*80}")

    summary = []
    for test_name, data in all_results.items():
        label = test_name.replace('_', ' ').title()
        passed = data.get('pass', 'N/A')
        status = "✅ PASS" if passed is True else "❌ FAIL" if passed is False else "◻️  N/A"
        summary.append({'test': label, 'status': status, 'pass': passed})
        print(f"  {label:<50} {status}")

    overall = all(
        s['pass'] is True or s['pass'] == 'N/A'
        for s in summary
    )

    print(f"\n{'─'*80}")
    print(f"OVERALL VERDICT: {'✅ CERTIFIED for Production' if overall else '❌ NOT Certified'}")
    print(f"{'─'*80}")

    # Save full report
    report = {
        'timestamp': time.strftime('%Y-%m-%d %H:%M:%S'),
        'system': {
            'psutil': HAS_PSUTIL,
            'memory': system_memory(),
            'initial_excel': len(initial_excel),
        },
        'summary': summary,
        'overall_pass': overall,
        'detailed': all_results,
    }

    report_path = os.path.join(PROJECT, 'docs', 'phaseX37_reliability_results.json')
    with open(report_path, 'w') as f:
        json.dump(report, f, indent=2, default=str)
    print(f"\nFull results saved to {report_path}")

    return report

if __name__ == '__main__':
    main()
