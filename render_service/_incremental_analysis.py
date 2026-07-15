"""
Phase X.3 — Incremental Sanitization Analysis
Tests each operation separately on the Japanese workbook.

Each step:
  1. Copy the original workbook
  2. Apply ONLY one operation
  3. Save as new XLSX
  4. Export to PDF (with IgnorePrintAreas=True, preserving FitToPages)
  5. Measure bounding box, pixel diff, coordinate positions
  6. Compare against original PDF

Operations tested:
  A: Original (baseline)
  B: Remove comments only
  C: Remove shapes only
  D: Remove headers/footers only
  E: Clear ALL cell values only
  F: Paint cluster cells BLACK only (no value clearing, no white fills)
  G: Full sanitizer (all operations combined)
"""
import sys, os, gc, json, tempfile, shutil, traceback
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
DOCS = r"C:\Users\MCF-JO~1\Documents"
sys.path.insert(0, ROOT)
os.chdir(ROOT)

JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")
DPI_VAL = 300
PT_TO_PX = DPI_VAL / 72.0
OUT_DIR = os.path.join(ROOT, "debug_output", "incremental_analysis")
os.makedirs(OUT_DIR, exist_ok=True)

import win32com.client
import numpy as np
import fitz

# ── Helpers ──
def pdf_info(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    doc.close()
    return {"media_box_pt": f"{w:.1f}x{h:.1f}", "w_pt": round(w, 1), "h_pt": round(h, 1)}

def render_img(pdf_path):
    zoom = DPI_VAL / 72.0
    doc = fitz.open(pdf_path)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    doc.close()
    return img, pix.width, pix.height

def content_bbox(img, white_thresh=240):
    """Get bounding box of non-white content. Returns (t,b,l,r) or None."""
    mask = ~np.all(img > white_thresh, axis=2)
    ys, xs = np.where(mask)
    if len(ys) == 0:
        return None
    return {"t": int(ys.min()), "b": int(ys.max()), "l": int(xs.min()), "r": int(xs.max())}

def pixel_diff_pct(img1, img2, thresh=20):
    diff = np.abs(img1.astype(np.int16) - img2.astype(np.int16))
    any_diff = np.any(diff > thresh, axis=2)
    return float(np.mean(any_diff) * 100)

def copy_workbook(excel, src_path):
    """Copy workbook to temp dir and return path to copy."""
    tmp = tempfile.mkdtemp(prefix="ple_incr_")
    dst = os.path.join(tmp, "copy.xlsx")
    wb = excel.Workbooks.Open(os.path.abspath(src_path))
    wb.SaveAs(os.path.abspath(dst))
    wb.Close(False)
    return dst, tmp

def open_copy(excel, path):
    """Open a workbook copy."""
    return excel.Workbooks.Open(os.path.abspath(path))

def export_pdf(ws_or_wb, pdf_path, ignorepa=True):
    """Export to PDF."""
    args = [0, os.path.abspath(pdf_path), 0, 0, True] if ignorepa else [0, os.path.abspath(pdf_path)]
    ws_or_wb.ExportAsFixedFormat(*args)

# ── Get clusters ──
print("Reading clusters...")
from render_service.upload_coordinate_generator import _identify_clusters
gc.collect()
cls = _identify_clusters(JAPAN)
print(f"  Found {len(cls)} clusters")
for c in cls:
    print(f"    {c['name']:20s}  {c['cellAddr']:15s}")

# Build cluster address set
CLUSTER_ADDRS = set(c["cellAddr"].upper().replace(" ", "") for c in cls)

# ── STAGE A: Original (baseline) ──
print(f"\n{'='*70}")
print(f"  STAGE A: Original workbook (baseline)")
print(f"{'='*70}")

all_stages = {}

# For each stage, we'll create the workbook, export PDF, render, and measure
def run_stage(label, fn_modify):
    """Run a stage: copy workbook, apply modification fn, export PDF, measure."""
    print(f"\n  --- {label} ---")
    stage_result = {"label": label}
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    
    try:
        # Copy original workbook
        copy_path, tmp_dir = copy_workbook(excel, JAPAN)
        print(f"    Copied to: {copy_path}")
        
        # Open the copy and apply modification
        wb = open_copy(excel, copy_path)
        
        # Apply the modification function
        fn_modify(excel, wb, copy_path)
        
        # Export to PDF (same settings as export_sanitized_pdf but with fix)
        pdf_tmp = tempfile.mkdtemp(prefix="ple_incr_pdf_")
        pdf_path = os.path.join(pdf_tmp, "output.pdf")
        
        # Delete _Fields sheet
        for i in range(wb.Sheets.Count, 0, -1):
            try:
                if wb.Sheets(i).Name in ("_Fields", "_RawData"):
                    wb.Sheets(i).Delete()
            except: pass
        
        ws = wb.ActiveSheet
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path), 0, 0, True)
        wb.Close(False)
        
        # Measure PDF
        pi = pdf_info(pdf_path)
        img, w_px, h_px = render_img(pdf_path)
        bb = content_bbox(img)
        
        stage_result["pdf"] = pi
        stage_result["img_dims"] = f"{w_px}x{h_px}"
        stage_result["bbox"] = bb
        
        print(f"    PDF: {pi['media_box_pt']}, Image: {w_px}x{h_px}")
        if bb:
            print(f"    Content bbox: ({bb['l']},{bb['t']})-({bb['r']},{bb['b']})")
            print(f"    Content dimensions: {bb['r']-bb['l']}x{bb['b']-bb['t']} px")
        else:
            print(f"    No non-white content found!")
        
        # Compare with baseline (Stage A)
        # We'll do the comparison after all stages are done
        
        # Save PDF for later comparison
        save_name = f"stage_{label.replace(' ','_')}.pdf"
        save_path = os.path.join(OUT_DIR, save_name)
        shutil.copy2(pdf_path, save_path)
        stage_result["saved_pdf"] = save_path
        print(f"    Saved: {save_path}")
        
        # Cleanup temp
        try: shutil.rmtree(pdf_tmp, ignore_errors=True)
        except: pass
        try: shutil.rmtree(tmp_dir, ignore_errors=True)
        except: pass
        
    except Exception as e:
        print(f"    ERROR: {e}")
        traceback.print_exc()
        stage_result["error"] = str(e)
    finally:
        try: excel.Quit()
        except: pass
    gc.collect()
    
    all_stages[label] = stage_result
    return stage_result

# ══════════════════════════════════════════════════════════════════════
# STAGE A: Original (no modification)
# ══════════════════════════════════════════════════════════════════════
def _noop(excel, wb, path):
    pass  # No modification

run_stage("A_Original", _noop)

# ══════════════════════════════════════════════════════════════════════
# STAGE B: Remove comments only
# ══════════════════════════════════════════════════════════════════════
def _remove_comments(excel, wb, path):
    for ws in wb.Worksheets:
        try:
            used = ws.UsedRange
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            for r in range(1, min(lr + 1, 200)):
                for c in range(1, min(lc + 1, 200)):
                    try:
                        cell = ws.Cells(r, c)
                        cm = cell.Comment
                        if cm:
                            cm.Delete()
                    except: pass
        except: pass

run_stage("B_RemoveComments", _remove_comments)

# ══════════════════════════════════════════════════════════════════════
# STAGE C: Remove shapes only
# ══════════════════════════════════════════════════════════════════════
def _remove_shapes(excel, wb, path):
    for ws in wb.Worksheets:
        try:
            for shape in list(ws.Shapes):
                shape.Delete()
        except: pass

run_stage("C_RemoveShapes", _remove_shapes)

# ══════════════════════════════════════════════════════════════════════
# STAGE D: Remove headers/footers only
# ══════════════════════════════════════════════════════════════════════
def _remove_headers(excel, wb, path):
    for ws in wb.Worksheets:
        try:
            ws.PageSetup.CenterHeader = ""
            ws.PageSetup.CenterFooter = ""
        except: pass

run_stage("D_RemoveHeaders", _remove_headers)

# ══════════════════════════════════════════════════════════════════════
# STAGE E: Clear all cell values only
# ══════════════════════════════════════════════════════════════════════
def _clear_values(excel, wb, path):
    for ws in wb.Worksheets:
        try:
            used = ws.UsedRange
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            for r in range(1, min(lr + 1, 200)):
                for c in range(1, min(lc + 1, 200)):
                    try:
                        ws.Cells(r, c).Value = ""
                    except: pass
        except: pass

run_stage("E_ClearValues", _clear_values)

# ══════════════════════════════════════════════════════════════════════
# STAGE F: Paint cluster cells black only
# ══════════════════════════════════════════════════════════════════════
def _paint_clusters(excel, wb, path):
    for ws in wb.Worksheets:
        try:
            used = ws.UsedRange
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            for r in range(1, min(lr + 1, 200)):
                for c in range(1, min(lc + 1, 200)):
                    try:
                        cell = ws.Cells(r, c)
                        ma = cell.MergeArea
                        if ma is not None:
                            addr = str(ma.Address).upper()
                            if addr in CLUSTER_ADDRS:
                                ma.Interior.Color = 1  # Black
                        else:
                            addr = str(cell.Address).upper()
                            if addr in CLUSTER_ADDRS:
                                cell.Interior.Color = 1  # Black
                    except: pass
        except: pass

run_stage("F_PaintClusters", _paint_clusters)

# ══════════════════════════════════════════════════════════════════════
# STAGE G: Full sanitizer (but WITHOUT the old Zoom/FitToPages clearing)
# ══════════════════════════════════════════════════════════════════════
def _full_sanitize(excel, wb, path):
    """Replicate the current sanitize_workbook function exactly."""
    for ws in wb.Worksheets:
        try:
            ws.PageSetup.CenterHeader = ""
            ws.PageSetup.CenterFooter = ""
        except: pass
        
        try:
            for shape in list(ws.Shapes):
                shape.Delete()
        except: pass
        
        try:
            used = ws.UsedRange
            if used is None: continue
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            
            for addr_str in CLUSTER_ADDRS:
                import re
                m = re.match(r"(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?", addr_str)
                if m:
                    def col_idx(s):
                        n = 0
                        for ch in s.replace("$", ""):
                            n = n * 26 + (ord(ch) - 64)
                        return n
                    c1 = col_idx(m.group(1))
                    c2 = col_idx(m.group(3)) if m.group(3) else c1
                    r1 = int(m.group(2).replace("$", ""))
                    r2 = int(m.group(4).replace("$", "")) if m.group(4) else r1
                    lc = max(lc, c1, c2)
                    lr = max(lr, r1, r2)
            
            for r in range(1, min(lr + 1, 200)):
                for c in range(1, min(lc + 1, 200)):
                    try:
                        cell = ws.Cells(r, c)
                        ma = cell.MergeArea
                        if ma is not None:
                            addr = str(ma.Address).upper()
                            is_cluster = addr in CLUSTER_ADDRS
                        else:
                            addr = str(cell.Address).upper()
                            is_cluster = addr in CLUSTER_ADDRS
                        
                        if is_cluster:
                            fc = ma if ma is not None else cell
                            fc.Interior.Color = 1  # Black
                            fc.Value = ""
                        else:
                            cell.Interior.Color = 0xFFFFFF  # White
                            cell.Value = ""
                    except: pass
            
            try:
                clear_range = ws.Range(ws.Cells(1, 1), ws.Cells(min(lr, 200), min(lc, 200)))
                clear_range.Borders.LineStyle = -4142  # xlNone
            except: pass
        except: pass

run_stage("G_FullSanitizer", _full_sanitize)

# ══════════════════════════════════════════════════════════════════════
# COMPARISON: All stages vs Stage A (Original)
# ══════════════════════════════════════════════════════════════════════
print(f"\n{'='*70}")
print(f"  COMPARISON: All stages vs Original")
print(f"{'='*70}")

baseline = all_stages.get("A_Original", {})
baseline_bbox = baseline.get("bbox")

# Render original for pixel comparisons
if baseline.get("saved_pdf"):
    img_orig, wo, ho = render_img(baseline["saved_pdf"])
else:
    img_orig = None

for label, data in all_stages.items():
    if label == "A_Original":
        continue
    print(f"\n  {label} vs Original:")
    
    bbox = data.get("bbox")
    if baseline_bbox and bbox:
        dx = bbox["l"] - baseline_bbox["l"]
        dy = bbox["t"] - baseline_bbox["t"]
        dw = (bbox["r"] - bbox["l"]) - (baseline_bbox["r"] - baseline_bbox["l"])
        dh = (bbox["b"] - bbox["t"]) - (baseline_bbox["b"] - baseline_bbox["t"])
        print(f"    Content bbox offset: ({dx}, {dy}) px")
        print(f"    Content size diff: {dw}x{dh} px")
        if abs(dx) > 5 or abs(dy) > 5:
            print(f"    ❌ GEOMETRY CHANGED (offset > 5px)")
        else:
            print(f"    ✅ Geometry preserved (offset <= 5px)")
    
    # Pixel diff
    if img_orig is not None and data.get("saved_pdf"):
        img_stage, _, _ = render_img(data["saved_pdf"])
        diff_pct = pixel_diff_pct(img_orig, img_stage)
        print(f"    Pixel diff vs original: {diff_pct:.3f}%")
        data["pixel_diff_vs_original"] = round(diff_pct, 3)

# ══════════════════════════════════════════════════════════════════════
# IDENTIFY FIRST DIVERGENCE
# ══════════════════════════════════════════════════════════════════════
print(f"\n{'='*70}")
print(f"  FIRST DIVERGENCE IDENTIFICATION")
print(f"{'='*70}")

# Find the first stage where geometry changes significantly
divergence_found = False
for label in ["B_RemoveComments", "C_RemoveShapes", "D_RemoveHeaders", 
              "E_ClearValues", "F_PaintClusters", "G_FullSanitizer"]:
    data = all_stages.get(label, {})
    bbox = data.get("bbox")
    if baseline_bbox and bbox:
        dx = abs(bbox["l"] - baseline_bbox["l"])
        dy = abs(bbox["t"] - baseline_bbox["t"])
        if (dx > 5 or dy > 5) and not divergence_found:
            print(f"\n  ❌ FIRST DIVERGENCE at: {label}")
            print(f"     Offset: ({dx}, {dy}) px")
            print(f"     This operation CHANGED the PDF geometry.")
            divergence_found = True
        elif dx > 5 or dy > 5:
            print(f"  ❌ ALSO divergent: {label} ({dx}, {dy}) px")
        else:
            print(f"  ✅ OK: {label} (no geometry change)")

if not divergence_found:
    print(f"  No geometry divergence found in any single operation.")
    print(f"  The issue may be cumulative.")

# ══════════════════════════════════════════════════════════════════════
# SAVE AND REPORT
# ══════════════════════════════════════════════════════════════════════
report_path = os.path.join(OUT_DIR, "incremental_results.json")
# Convert non-serializable types
class SimpleEnc(json.JSONEncoder):
    def default(self, o):
        return str(o)

with open(report_path, 'w', encoding='utf-8') as f:
    json.dump(all_stages, f, ensure_ascii=False, indent=2, cls=SimpleEnc)

print(f"\n  Full results: {report_path}")
print(f"  PDFs saved in: {OUT_DIR}")
print(f"  Done.")
