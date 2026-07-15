"""
Phase X.6 — Production Execution Trace & ConMas Comparison

Instruments the /upload/preview endpoint with per-stage timestamps,
thread IDs, COM threading state, and full exception capture.

Starts an instrumented server on port 5092, sends the Japanese workbook,
captures the full execution trace, then compares with the benchmark.
"""

import os
import sys
import time
import json
import io
import threading
import tempfile
import shutil
import traceback as tb
from pathlib import Path
from contextlib import asynccontextmanager

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

JAPANESE_WORKBOOK = r"C:\Users\MCF-JO~1\Documents\[V3.1_Sample]アンケート用紙.xlsx"
FORMTEST_WORKBOOK = r"C:\Users\MCF-JO~1\Documents\FormTest - Copy.xlsx"
INSTR_PORT = 5092
OUTPUT_DIR = os.path.join(_THIS_DIR, "timing_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# Part 1: Instrumented FastAPI endpoint
# ============================================================

log_entries = []  # thread-safe list of log dicts
_log_lock = threading.Lock()


def log_timestamp(stage: str, details: str = ""):
    """Log a timestamp with thread ID and process ID."""
    import os as _os
    entry = {
        "stage": stage,
        "details": details,
        "elapsed_ms": (time.perf_counter() - _t0) * 1000 if _t0 else 0,
        "thread_id": threading.current_thread().ident,
        "thread_name": threading.current_thread().name,
        "process_id": _os.getpid(),
    }
    with _log_lock:
        log_entries.append(entry)
    print(f"[{entry['elapsed_ms']:>8.0f}ms] [T{entry['thread_id']}] {stage}: {details}")


_t0 = None


def check_com_apartment():
    """Check and log the COM apartment state."""
    try:
        import pythoncom
        try:
            state = pythoncom._GetInterfaceCount()  # Just check if pythoncom is initialized
            apt_state = "Initialized"
        except Exception:
            apt_state = "Not initialized"
        log_timestamp("COM_APARTMENT_CHECK", f"pythoncom state: {apt_state}")
    except ImportError:
        log_timestamp("COM_APARTMENT_CHECK", "pythoncom not importable")


def make_instrumented_preview_endpoint():
    """Create a FastAPI app with instrumented /upload/preview endpoint."""

    import pythoncom
    import win32com.client
    import fitz
    import numpy as np
    from fastapi import FastAPI, HTTPException, UploadFile, File
    from fastapi.responses import JSONResponse

    app = FastAPI(title="Instrumented Preview Server")

    from render_service.upload_coordinate_generator import (
        _identify_clusters, sanitize_workbook, export_sanitized_pdf,
        render_pdf_to_image, scan_black_rectangles, split_merged_rects,
        normalize_rects, _sort_key_meta, generate_preview as original_generate_preview,
        generate_coordinates as original_generate_coordinates
    )
    from render_service.pdf_converter import xlsx_to_pdf
    from render_service.background_renderer import pdf_page_to_png, get_page_dimensions

    @app.post("/upload/preview_instrumented")
    async def upload_preview_instrumented(
        file: UploadFile = File(...),
        output_dir: str | None = None,
    ):
        global _t0
        _t0 = time.perf_counter()
        log_entries.clear()

        log_timestamp("REQUEST_RECEIVED", f"File: {file.filename}")

        # Initialize COM for this thread
        log_timestamp("COM_INIT", "Calling pythoncom.CoInitialize()")
        try:
            pythoncom.CoInitialize()
            log_timestamp("COM_INIT", "CoInitialize() succeeded")
        except Exception as e:
            log_timestamp("COM_INIT_ERROR", f"CoInitialize() failed: {e}")
            # Already initialized — that's OK

        tmp_dir = tempfile.mkdtemp(prefix="ple_preview_instr_")
        try:
            # Save uploaded file
            log_timestamp("FILE_READ_START", f"Reading {file.filename}")
            content = await file.read()
            log_timestamp("FILE_READ_DONE", f"Read {len(content)} bytes")

            xlsx_path = os.path.join(tmp_dir, file.filename or "upload.xlsx")
            with open(xlsx_path, "wb") as f:
                f.write(content)
            log_timestamp("FILE_SAVED", f"Saved to {xlsx_path}")

            out_dir = output_dir or tmp_dir

            # ── TIMED: generate_preview() ──────────────────────────────
            log_timestamp("GENERATE_PREVIEW_START", "Entering generate_preview()")

            log_timestamp("IDENTIFY_CLUSTERS_START", "Calling _identify_clusters()")
            check_com_apartment()
            try:
                cluster_meta = _identify_clusters(xlsx_path)
                log_timestamp("IDENTIFY_CLUSTERS_DONE", f"Found {len(cluster_meta)} clusters")
            except Exception as e:
                log_timestamp("IDENTIFY_CLUSTERS_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                cluster_meta = []

            if cluster_meta:
                log_timestamp("SANITIZE_WORKBOOK_START", "Calling sanitize_workbook()")
                try:
                    sanitized_path = sanitize_workbook(xlsx_path, cluster_meta)
                    log_timestamp("SANITIZE_WORKBOOK_DONE", f"Path: {sanitized_path}")
                except Exception as e:
                    log_timestamp("SANITIZE_WORKBOOK_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                    raise

                log_timestamp("EXPORT_SANITIZED_PDF_START", "Calling export_sanitized_pdf()")
                try:
                    pdf_path = export_sanitized_pdf(sanitized_path)
                    log_timestamp("EXPORT_SANITIZED_PDF_DONE", f"Path: {pdf_path}")
                except Exception as e:
                    log_timestamp("EXPORT_SANITIZED_PDF_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                    raise

                log_timestamp("RENDER_SANITIZED_PDF_START", "Rendering PDF to image")
                try:
                    img, img_w, img_h = render_pdf_to_image(pdf_path)
                    log_timestamp("RENDER_SANITIZED_PDF_DONE", f"Image: {img_w}x{img_h}")
                except Exception as e:
                    log_timestamp("RENDER_SANITIZED_PDF_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                    raise

                log_timestamp("SCAN_RECTANGLES_START", "Scanning for black rectangles")
                try:
                    rects = scan_black_rectangles(img)
                    log_timestamp("SCAN_RECTANGLES_DONE", f"Found {len(rects)} rectangles")
                except Exception as e:
                    log_timestamp("SCAN_RECTANGLES_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                    raise

                log_timestamp("SPLIT_MERGED_START", "Splitting merged rectangles")
                try:
                    split_rects = split_merged_rects(rects, cluster_meta)
                    log_timestamp("SPLIT_MERGED_DONE", f"After split: {len(split_rects)}")
                except Exception as e:
                    log_timestamp("SPLIT_MERGED_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                    raise

                log_timestamp("NORMALIZE_START", "Normalizing rectangles")
                normalized = normalize_rects(split_rects, img_w, img_h)

                # Sort and merge
                cluster_meta.sort(key=_sort_key_meta)
                normalized.sort(key=lambda r: (r["Top"], r["Left"]))
                fields = []
                for meta, rect in zip(cluster_meta, normalized):
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
                log_timestamp("NORMALIZE_DONE", f"Generated {len(fields)} fields")

                # Cleanup sanitized/PDF temp dirs
                for p in [sanitized_path, pdf_path]:
                    d = os.path.dirname(p)
                    try:
                        shutil.rmtree(d, ignore_errors=True)
                    except Exception:
                        pass
            else:
                fields = []

            log_timestamp("EXPORT_ORIGINAL_PDF_START", "Exporting original workbook PDF")
            try:
                orig_pdf_path = xlsx_to_pdf(xlsx_path)
                log_timestamp("EXPORT_ORIGINAL_PDF_DONE", f"Path: {orig_pdf_path}")
            except Exception as e:
                log_timestamp("EXPORT_ORIGINAL_PDF_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                raise

            log_timestamp("RENDER_BACKGROUND_PNG_START", "Rendering PNG for background")
            try:
                png_path = os.path.join(out_dir, f"preview_instr.png")
                pdf_page_to_png(orig_pdf_path, png_path, dpi=300)
                page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=300)
                log_timestamp("RENDER_BACKGROUND_PNG_DONE", f"PNG: {page_w}x{page_h}")
            except Exception as e:
                log_timestamp("RENDER_BACKGROUND_PNG_ERROR", f"Failed: {e}\n{tb.format_exc()}")
                raise

            # Cleanup original PDF
            try:
                os.unlink(orig_pdf_path)
                pdf_dir = os.path.dirname(orig_pdf_path)
                if pdf_dir and pdf_dir != tmp_dir:
                    shutil.rmtree(pdf_dir, ignore_errors=True)
            except Exception:
                pass

            log_timestamp("GENERATE_PREVIEW_DONE", "generate_preview() complete")

            # Build response
            log_timestamp("JSON_SERIALIZE_START", "Building JSON response")
            response_data = {
                "success": True,
                "backgroundImage": "preview_instr.png",
                "page": {"width": page_w, "height": page_h},
                "fields": fields,
            }
            log_timestamp("JSON_SERIALIZE_DONE", f"{len(json.dumps(response_data))} bytes")

            # CoUninitialize COM
            log_timestamp("COM_UNINIT", "Calling pythoncom.CoUninitialize()")
            try:
                pythoncom.CoUninitialize()
                log_timestamp("COM_UNINIT", "CoUninitialize() succeeded")
            except Exception as e:
                log_timestamp("COM_UNINIT_ERROR", f"CoUninitialize() failed: {e}")

            log_timestamp("RESPONSE_RETURNED", "Sending HTTP response")
            return JSONResponse(content=response_data)

        except Exception as e:
            log_timestamp("FATAL_ERROR", f"Unhandled exception: {e}\n{tb.format_exc()}")
            raise HTTPException(status_code=500, detail={
                "success": False,
                "message": str(e),
                "traceback": tb.format_exc(),
            })
        finally:
            # Cleanup temp dir
            if output_dir:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            # Always try CoUninitialize
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return app


# ============================================================
# Runner
# ============================================================

def run_instrumented_test():
    """Start the instrumented server, send a request, capture timestamps."""
    import uvicorn
    import requests
    import multiprocessing

    # Create the instrumented app
    app = make_instrumented_preview_endpoint()

    # Start server in a subprocess
    import subprocess
    import time as _time

    server_script = os.path.join(_THIS_DIR, "_instrumented_server_worker.py")

    # Write the worker script
    worker_code = '''
import sys, os
sys.path.insert(0, r"{project_root}")
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
import tempfile, shutil, time, json, threading, traceback as tb
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import pythoncom, win32com.client, fitz, numpy as np

app = FastAPI(title="Instrumented Preview Worker")

from render_service.upload_coordinate_generator import (
    _identify_clusters, sanitize_workbook, export_sanitized_pdf,
    render_pdf_to_image, scan_black_rectangles, split_merged_rects,
    normalize_rects, _sort_key_meta
)
from render_service.pdf_converter import xlsx_to_pdf
from render_service.background_renderer import pdf_page_to_png, get_page_dimensions

log_entries = []
_log_lock = threading.Lock()
_t0 = [0]

def log_ts(stage, details=""):
    entry = {{
        "stage": stage,
        "details": details,
        "elapsed_ms": (time.perf_counter() - _t0[0]) * 1000,
        "thread_id": threading.current_thread().ident,
        "process_id": os.getpid(),
    }}
    with _log_lock:
        log_entries.append(entry)
    print(f"[{{entry['elapsed_ms']:>8.0f}}ms] [T{{entry['thread_id']}}] {{stage}}: {{details}}", flush=True)

@app.post("/upload/preview")
async def preview_instrumented(file: UploadFile = File(...), output_dir: str | None = None):
    _t0[0] = time.perf_counter()
    log_entries.clear()
    log_ts("REQUEST_RECEIVED", f"File: {{file.filename}}")

    try:
        pythoncom.CoInitialize()
        log_ts("COM_INIT", "CoInitialize() succeeded")
    except Exception as e:
        log_ts("COM_INIT", f"Already initialized or error: {{e}}")

    tmp_dir = tempfile.mkdtemp(prefix="ple_instr_")
    try:
        content = await file.read()
        xlsx_path = os.path.join(tmp_dir, file.filename or "upload.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(content)
        log_ts("FILE_SAVED", f"{{len(content)}} bytes")

        out_dir = output_dir or tmp_dir

        log_ts("IDENTIFY_CLUSTERS_START")
        try:
            cluster_meta = _identify_clusters(xlsx_path)
            log_ts("IDENTIFY_CLUSTERS_DONE", f"{{len(cluster_meta)}} clusters")
        except Exception as e:
            log_ts("IDENTIFY_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("SANITIZE_START")
        try:
            sanitized_path = sanitize_workbook(xlsx_path, cluster_meta)
            log_ts("SANITIZE_DONE")
        except Exception as e:
            log_ts("SANITIZE_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("EXPORT_PDF_START")
        try:
            pdf_path = export_sanitized_pdf(sanitized_path)
            log_ts("EXPORT_PDF_DONE")
        except Exception as e:
            log_ts("EXPORT_PDF_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("RENDER_PDF_START")
        try:
            img, img_w, img_h = render_pdf_to_image(pdf_path)
            log_ts("RENDER_PDF_DONE", f"{{img_w}}x{{img_h}}")
        except Exception as e:
            log_ts("RENDER_PDF_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("SCAN_RECTS_START")
        try:
            rects = scan_black_rectangles(img)
            log_ts("SCAN_RECTS_DONE", f"{{len(rects)}} rects")
        except Exception as e:
            log_ts("SCAN_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("SPLIT_START")
        try:
            split_rects = split_merged_rects(rects, cluster_meta)
            log_ts("SPLIT_DONE", f"{{len(split_rects)}} after split")
        except Exception as e:
            log_ts("SPLIT_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("NORMALIZE_START")
        normalized = normalize_rects(split_rects, img_w, img_h)
        cluster_meta.sort(key=_sort_key_meta)
        normalized.sort(key=lambda r: (r["Top"], r["Left"]))
        fields = []
        for meta, rect in zip(cluster_meta, normalized):
            fields.append({{
                "name": meta["name"],
                "type": meta["type"],
                "cellAddr": meta["cellAddr"],
                "input_parameter": meta.get("input_parameter", ""),
                "left_ratio": round(rect["left_ratio"], 7),
                "top_ratio": round(rect["top_ratio"], 7),
                "right_ratio": round(rect["right_ratio"], 7),
                "bottom_ratio": round(rect["bottom_ratio"], 7),
            }})
        log_ts("NORMALIZE_DONE", f"{{len(fields)}} fields")

        # Cleanup sanitized temp
        for p in [sanitized_path, pdf_path]:
            try:
                shutil.rmtree(os.path.dirname(p), ignore_errors=True)
            except:
                pass

        # Original PDF background
        log_ts("ORIG_PDF_START")
        try:
            orig_pdf_path = xlsx_to_pdf(xlsx_path)
            log_ts("ORIG_PDF_DONE")
        except Exception as e:
            log_ts("ORIG_PDF_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        log_ts("RENDER_BG_START")
        try:
            png_path = os.path.join(out_dir, "preview_instr.png")
            pdf_page_to_png(orig_pdf_path, png_path, dpi=300)
            page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=300)
            log_ts("RENDER_BG_DONE", f"{{page_w}}x{{page_h}}")
        except Exception as e:
            log_ts("RENDER_BG_ERROR", f"{{e}}\\n{{tb.format_exc()}}")
            raise

        try:
            os.unlink(orig_pdf_path)
            shutil.rmtree(os.path.dirname(orig_pdf_path), ignore_errors=True)
        except:
            pass

        log_ts("JSON_SERIALIZE")
        response = json.dumps({{
            "success": True,
            "backgroundImage": "preview_instr.png",
            "page": {{"width": page_w, "height": page_h}},
            "fields": fields,
        }})
        log_ts("SERIALIZE_DONE", f"{{len(response)}} bytes")

        try:
            pythoncom.CoUninitialize()
            log_ts("COM_UNINIT", "CoUninitialize() succeeded")
        except Exception as e:
            log_ts("COM_UNINIT_ERROR", f"{{e}}")

        log_ts("RESPONSE_SENT")
        return JSONResponse(content=json.loads(response))

    except Exception as e:
        log_ts("FATAL", f"{{e}}\\n{{tb.format_exc()}}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return JSONResponse(status_code=500, content={{"success": False, "message": str(e), "traceback": tb.format_exc(), "log": log_entries}})
    finally:
        if output_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port={port})
'''.format(project_root=_PROJECT_ROOT, port=INSTR_PORT)

    with open(server_script, "w", encoding="utf-8") as f:
        f.write(worker_code)

    # Start server
    proc = subprocess.Popen(
        [sys.executable, server_script],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        cwd=_THIS_DIR,
    )
    _time.sleep(3)  # Wait for server to start

    report_path = os.path.join(OUTPUT_DIR, "production_trace.log")
    try:
        # Send request to instrumented endpoint
        for label, path in [("FormTest", FORMTEST_WORKBOOK), ("Japanese", JAPANESE_WORKBOOK)]:
            print(f"\nSending {label}...")
            url = f"http://127.0.0.1:{INSTR_PORT}/upload/preview"
            with open(path, "rb") as f:
                file_bytes = f.read()

            t_start = time.perf_counter()
            try:
                resp = requests.post(
                    url,
                    files={"file": (os.path.basename(path), io.BytesIO(file_bytes),
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    timeout=180,
                )
                elapsed = (time.perf_counter() - t_start) * 1000
                print(f"  Response: {resp.status_code} in {elapsed:.0f}ms")
                print(f"  Body: {resp.text[:500]}")
            except requests.Timeout:
                elapsed = (time.perf_counter() - t_start) * 1000
                print(f"  TIMEOUT after {elapsed:.0f}ms")
            except Exception as e:
                print(f"  ERROR: {e}")

            _time.sleep(2)  # Wait between tests

    finally:
        # Kill server
        proc.terminate()
        proc.wait(timeout=5)

        # Read server stdout (contains all timestamps)
        stdout_data = proc.stdout.read().decode("utf-8", errors="replace") if proc.stdout else ""

        # Save to file
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(stdout_data)

        # Cleanup worker script
        try:
            os.unlink(server_script)
        except:
            pass

    print(f"\nTrace saved to: {report_path}")
    return report_path


if __name__ == "__main__":
    run_instrumented_test()
