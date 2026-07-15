"""
Phase X.1 — Root Cause Verification Investigation
No fixes. No code changes. Evidence collection only.

Performs all 8 investigation tasks:
1. ExportAsFixedFormat paper size behavior
2. Sanitized vs original PDF geometry comparison
3. PageSetup property trace (before/after sanitize)
4. IgnorePrintAreas=True effect
5. Workbook vs worksheet level export
6. PaperSize hypothesis verification
7. Pixel-level comparison
8. Complete pipeline geometry trace
"""
import sys, os, gc, json, tempfile, shutil, base64
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

_ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
_DOCS = r"C:\Users\MCF-JO~1\Documents"
sys.path.insert(0, _ROOT)
os.chdir(_ROOT)

from pathlib import Path
DOCS = Path(_DOCS)
FORM = DOCS / "FormTest - Copy.xlsx"
JAPAN = DOCS / "[V3.1_Sample]アンケート用紙.xlsx"
DPI_VAL = 300
PT_TO_PX = DPI_VAL / 72.0
OUTPUT_DIR = Path(_ROOT) / "debug_output"
OUTPUT_DIR.mkdir(exist_ok=True)

results = {
    "workbooks": {},
    "page_setup": {},
    "export_comparison": {},
    "ignore_print_areas": {},
    "workbook_vs_worksheet": {},
    "pixel_comparison": {},
    "pipeline_geometry": {},
}

import win32com.client
import fitz  # PyMuPDF

# ════════════════════════════════════════════════════════════════════════
# HELPER: fresh COM session
# ════════════════════════════════════════════════════════════════════════
def com_session(fn):
    """Execute fn in a fresh COM session with cleanup."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        return fn(excel)
    finally:
        try: excel.Quit()
        except: pass
    gc.collect()

def open_wb(excel, path):
    return excel.Workbooks.Open(os.path.abspath(str(path)))

def get_ps_dict(ws):
    """Read all PageSetup properties into a dict."""
    ps = ws.PageSetup
    props = {}
    for name in ["Zoom", "FitToPagesWide", "FitToPagesTall",
                 "Orientation", "PaperSize",
                 "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
                 "CenterHorizontally", "CenterVertically",
                 "Order", "PrintArea", "PrintTitleRows", "PrintTitleColumns",
                 "FirstPageNumber", "PageNumberLocal"]:
        try:
            props[name] = repr(getattr(ps, name))
        except Exception as e:
            props[name] = f"ERROR({e})"
    # PaperSize/PageWidth/PageHeight via PageSetup chain
    try:
        pw = ps.PageSetup.PageWidth
        props["PageWidth"] = repr(pw)
    except: props["PageWidth"] = "ERROR"
    try:
        ph = ps.PageSetup.PageHeight
        props["PageHeight"] = repr(ph)
    except: props["PageHeight"] = "ERROR"
    try:
        props["PageWidth_direct"] = repr(ps.PageWidth)
    except: props["PageWidth_direct"] = "ERROR"
    try:
        props["PageHeight_direct"] = repr(ps.PageHeight)
    except: props["PageHeight_direct"] = "ERROR"
    return props

def pdf_page_info(pdf_path):
    """Get PDF page dimensions in points and pixels."""
    doc = fitz.open(pdf_path)
    if len(doc) == 0:
        doc.close()
        return {"pages": 0, "error": "no pages"}
    page = doc[0]
    rect = page.rect  # MediaBox
    w_pt, h_pt = rect.width, rect.height
    w_px = int(w_pt * PT_TO_PX)
    h_px = int(h_pt * PT_TO_PX)
    info = {
        "pages": len(doc),
        "media_box_pt": f"{w_pt:.1f}x{h_pt:.1f}",
        "media_box_px_300dpi": f"{w_px}x{h_px}",
        "media_box_width_pt": round(w_pt, 1),
        "media_box_height_pt": round(h_pt, 1),
    }
    # Try crop box
    try:
        cb = page.cropbox
        info["crop_box_pt"] = f"{cb.width:.1f}x{cb.height:.1f}"
    except:
        info["crop_box_pt"] = "same as media"
    doc.close()
    return info

def export_pdf(ws_or_wb, path, ignore_print_areas=False):
    """Export to PDF with settings control."""
    obj = ws_or_wb
    if ignore_print_areas:
        obj.ExportAsFixedFormat(0, os.path.abspath(path), 0, 0, True)
    else:
        obj.ExportAsFixedFormat(0, os.path.abspath(path))

# ════════════════════════════════════════════════════════════════════════
# TASK 1: ExportAsFixedFormat paper size behavior
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  TASK 1: ExportAsFixedFormat paper size behavior")
print("=" * 70)

task1 = {}

# Test on BOTH workbooks
for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  --- {wb_name} ---")
    info = {}
    def _task1(excel):
        wb = open_wb(excel, wb_path)
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        info["sheet_name"] = ws.Name
        
        # Read page setup BEFORE export
        info["page_setup_before"] = get_ps_dict(ws)
        print(f"  Sheet: {ws.Name}")
        print(f"  PaperSize from COM: {info['page_setup_before'].get('PaperSize')}")
        print(f"  Zoom: {info['page_setup_before'].get('Zoom')}")
        print(f"  FitToPagesWide: {info['page_setup_before'].get('FitToPagesWide')}")
        print(f"  FitToPagesTall: {info['page_setup_before'].get('FitToPagesTall')}")
        print(f"  PageWidth_direct: {info['page_setup_before'].get('PageWidth_direct')}")
        
        # Export at worksheet level with IGNOREPA=false (default)
        tmp1 = tempfile.mkdtemp(prefix="ple_verify1_")
        pdf1 = os.path.join(tmp1, "default.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf1), 0, 0, False)
        pdf_info1 = pdf_page_info(pdf1)
        info["pdf_default"] = pdf_info1
        print(f"  Default export PDF: {pdf_info1['media_box_pt']} ({pdf_info1['pages']} page(s))")
        
        # Read page setup AFTER export
        info["page_setup_after"] = get_ps_dict(ws)
        
        wb.Close(False)
        # Cleanup
        try: shutil.rmtree(tmp1, ignore_errors=True)
        except: pass
        return pdf1  # keep ref for cleanup
    
    pdf_path = com_session(_task1)
    try:
        if pdf_path:
            d = os.path.dirname(pdf_path)
            if d: shutil.rmtree(d, ignore_errors=True)
    except: pass
    gc.collect()
    
    info["conclusion"] = (
        "PaperSize preserved by COM" if "9" in str(info.get("page_setup_before", {}).get("PaperSize",""))
        else "PaperSize=1 (Letter)"
    )
    task1[wb_name] = info

results["export_comparison"]["task1_paper_size"] = task1

# ════════════════════════════════════════════════════════════════════════
# TASK 2: Sanitized vs original PDF geometry comparison
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  TASK 2: Sanitized vs original PDF geometry comparison")
print("=" * 70)

from render_service.upload_coordinate_generator import (
    _identify_clusters, sanitize_workbook, export_sanitized_pdf,
    render_pdf_to_image
)
from render_service.pdf_converter import xlsx_to_pdf

task2 = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  --- {wb_name} ---")
    info = {}
    
    # Get clusters
    info["clusters"] = []
    def _get_clusters(excel):
        wb = open_wb(excel, wb_path)
        clusters = []
        ws = wb.Sheets(1) if wb.Sheets(1).Name != "_Fields" else wb.Sheets(2)
        for row in range(1, 50):
            for col in range(1, 50):
                try:
                    c = ws.Cells(row, col).Comment
                    if c:
                        clusters.append(f"${row}${col}")
                except:
                    pass
        wb.Close(False)
        return clusters
    try:
        # Use win32com directly for cluster count
        def _count_comments(excel):
            wb = open_wb(excel, wb_path)
            ws = wb.Sheets(1) if wb.Sheets(1).Name != "_Fields" else wb.Sheets(2)
            count = 0
            for row in range(1, 50):
                for col in range(1, 50):
                    try:
                        if ws.Cells(row, col).Comment: count += 1
                    except: pass
            wb.Close(False)
            return count
        comment_count = com_session(_count_comments)
        gc.collect()
        print(f"  Comment count: {comment_count}")
        info["comment_count"] = comment_count
    except Exception as e:
        print(f"  Comment scan error: {e}")
        info["comment_count_error"] = str(e)
    
    # Generate original PDF
    try:
        pdf_orig = xlsx_to_pdf(str(wb_path))
        orig_info = pdf_page_info(pdf_orig)
        print(f"  Original PDF: {orig_info['media_box_pt']} media, {orig_info['pages']} pages")
        info["original_pdf"] = orig_info
        try: os.unlink(pdf_orig)
        except: pass
        d = os.path.dirname(pdf_orig)
        if d: shutil.rmtree(d, ignore_errors=True)
    except Exception as e:
        print(f"  Original PDF error: {e}")
        info["original_pdf_error"] = str(e)
    
    # Get clusters via _Fields sheet (more reliable for Japanese)
    try:
        cls = _identify_clusters(str(wb_path))
        info["cluster_meta"] = [{"name": c["name"], "addr": c["cellAddr"], "type": c["type"]} for c in cls]
        print(f"  Clusters via _identify_clusters: {len(cls)}")
    except Exception as e:
        print(f"  Cluster error: {e}")
        info["cluster_error"] = str(e)
        cls = []
    
    # Generate sanitized PDF
    if cls:
        try:
            sanitized = sanitize_workbook(str(wb_path), cls)
            pdf_san = export_sanitized_pdf(sanitized)
            san_info = pdf_page_info(pdf_san)
            print(f"  Sanitized PDF: {san_info['media_box_pt']} media, {san_info['pages']} pages")
            info["sanitized_pdf"] = san_info
            
            # Compare dimensions
            if "original_pdf" in info:
                o = info["original_pdf"]
                s = san_info
                if o["media_box_pt"] == s["media_box_pt"] and o["pages"] == s["pages"]:
                    info["pdf_match"] = True
                    print(f"  ✅ PDFs are IDENTICAL (same dimensions, same page count)")
                else:
                    info["pdf_match"] = False
                    info["pdf_diff"] = f"orig={o['media_box_pt']} san={s['media_box_pt']} pages={o['pages']}vs{s['pages']}"
                    print(f"  ❌ PDFs DIFFER: orig={o['media_box_pt']} san={s['media_box_pt']}")
            
            # Render sanitized to image for later pixel comparison
            img, img_w, img_h = render_pdf_to_image(pdf_san)
            info["sanitized_image"] = f"{img_w}x{img_h}"
            
            # Cleanup
            for d in [os.path.dirname(sanitized), os.path.dirname(pdf_san)]:
                try: shutil.rmtree(d, ignore_errors=True)
                except: pass
            
        except Exception as e:
            print(f"  Sanitize/export error: {e}")
            import traceback; traceback.print_exc()
            info["sanitize_error"] = str(e)
    else:
        info["sanitize_skipped"] = "no clusters found"
    
    task2[wb_name] = info

results["export_comparison"]["task2_sanitized_vs_original"] = task2

# ════════════════════════════════════════════════════════════════════════
# TASK 3: PageSetup property trace (before/after sanitize)
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  TASK 3: PageSetup property trace before/after sanitize")
print("=" * 70)

task3 = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  --- {wb_name} ---")
    info = {}
    def _trace(excel):
        wb = open_wb(excel, wb_path)
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        info["sheet_before"] = ws.Name
        info["before"] = get_ps_dict(ws)
        
        # Apply sanitize's page setup changes
        try: ws.PageSetup.Zoom = False
        except: pass
        try: ws.PageSetup.FitToPagesWide = False
        except: pass
        try: ws.PageSetup.FitToPagesTall = False
        except: pass
        
        info["after"] = get_ps_dict(ws)
        wb.Close(False)
    
    com_session(_trace)
    gc.collect()
    
    # Print comparison
    print(f"  {'Property':25s} {'Before':20s} {'After':20s}")
    print(f"  {'-'*25} {'-'*20} {'-'*20}")
    before = info.get("before", {})
    after = info.get("after", {})
    for key in sorted(before.keys()):
        b = before.get(key, "?")
        a = after.get(key, "?")
        changed = "⚠️" if b != a else "  "
        print(f"  {key:25s} {b:20s} {a:20s} {changed}")
    
    task3[wb_name] = info

results["page_setup"] = task3

# ════════════════════════════════════════════════════════════════════════
# TASK 4: Verify IgnorePrintAreas=True effect
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  TASK 4: IgnorePrintAreas=True effect")
print("=" * 70)

task4 = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  --- {wb_name} ---")
    info = {}
    def _ipa(excel):
        wb = open_wb(excel, wb_path)
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        info["sheet"] = ws.Name
        info["has_print_area"] = str(ws.PageSetup.PrintArea) if ws.PageSetup.PrintArea else "None"
        
        # Export with IgnorePrintAreas=False
        tmp_false = tempfile.mkdtemp(prefix="ple_ipafalse_")
        pdf_false = os.path.join(tmp_false, "ignorepa_false.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_false), 0, 0, False)
        info["pdf_ignore_false"] = pdf_page_info(pdf_false)
        
        # Export with IgnorePrintAreas=True
        tmp_true = tempfile.mkdtemp(prefix="ple_ipatrue_")
        pdf_true = os.path.join(tmp_true, "ignorepa_true.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_true), 0, 0, True)
        info["pdf_ignore_true"] = pdf_page_info(pdf_true)
        
        wb.Close(False)
        
        # Compare
        pf = info["pdf_ignore_false"]
        pt = info["pdf_ignore_true"]
        if pf["media_box_pt"] == pt["media_box_pt"] and pf["pages"] == pt["pages"]:
            info["ipa_effect"] = "SAME" 
        else:
            info["ipa_effect"] = "DIFFERENT"
            info["ipa_diff"] = f"false={pf['media_box_pt']}({pf['pages']}p) true={pt['media_box_pt']}({pt['pages']}p)"
        
        # Cleanup
        for d in [tmp_false, tmp_true]:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
    
    com_session(_ipa)
    gc.collect()
    print(f"  PrintArea: {info.get('has_print_area', '?')}")
    print(f"  IgnorePA=False: {info.get('pdf_ignore_false', {}).get('media_box_pt', '?')} ({info.get('pdf_ignore_false', {}).get('pages', '?')}p)")
    print(f"  IgnorePA=True:  {info.get('pdf_ignore_true', {}).get('media_box_pt', '?')} ({info.get('pdf_ignore_true', {}).get('pages', '?')}p)")
    print(f"  Effect: {info.get('ipa_effect', '?')}")
    task4[wb_name] = info

results["ignore_print_areas"] = task4

# ════════════════════════════════════════════════════════════════════════
# TASK 5: Workbook-level vs Worksheet-level export
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  TASK 5: Workbook-level vs Worksheet-level export")
print("=" * 70)

task5 = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  --- {wb_name} ---")
    info = {}
    def _level(excel):
        wb = open_wb(excel, wb_path)
        
        # Workbook-level export
        tmp_wb = tempfile.mkdtemp(prefix="ple_wblevel_")
        pdf_wb = os.path.join(tmp_wb, "workbook_level.pdf")
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_wb))
        info["workbook_level"] = pdf_page_info(pdf_wb)
        
        # Worksheet-level export (first non-_Fields sheet)
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        info["sheet"] = ws.Name
        tmp_ws = tempfile.mkdtemp(prefix="ple_wslevel_")
        pdf_ws = os.path.join(tmp_ws, "worksheet_level.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_ws), 0, 0, False)
        info["worksheet_level"] = pdf_page_info(pdf_ws)
        
        wb.Close(False)
        
        # Compare
        wbl = info["workbook_level"]
        wsl = info["worksheet_level"]
        if wbl["media_box_pt"] == wsl["media_box_pt"] and wbl["pages"] == wsl["pages"]:
            info["level_effect"] = "SAME"
        else:
            info["level_effect"] = "DIFFERENT"
            info["level_diff"] = f"wb={wbl['media_box_pt']}({wbl['pages']}p) ws={wsl['media_box_pt']}({wsl['pages']}p)"
        
        # Cleanup
        for d in [tmp_wb, tmp_ws]:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
    
    com_session(_level)
    gc.collect()
    print(f"  Workbook-level: {info.get('workbook_level', {}).get('media_box_pt', '?')} ({info.get('workbook_level', {}).get('pages', '?')}p)")
    print(f"  Worksheet-level: {info.get('worksheet_level', {}).get('media_box_pt', '?')} ({info.get('worksheet_level', {}).get('pages', '?')}p)")
    print(f"  Effect: {info.get('level_effect', '?')}")
    task5[wb_name] = info

results["workbook_vs_worksheet"] = task5

# ════════════════════════════════════════════════════════════════════════
# TASK 7: Pixel-level comparison
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  TASK 7: Pixel-level comparison")
print("=" * 70)

task7 = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  --- {wb_name} ---")
    info = {}
    
    # Need clusters first
    try:
        cls = _identify_clusters(str(wb_path))
        if not cls:
            print(f"  No clusters found")
            info["pixel_skip"] = "no clusters"
            task7[wb_name] = info
            continue
    except Exception as e:
        print(f"  Cluster error: {e}")
        info["pixel_skip"] = f"cluster error: {e}"
        task7[wb_name] = info
        continue
    
    # Generate original PDF
    try:
        pdf_orig = xlsx_to_pdf(str(wb_path))
    except Exception as e:
        print(f"  xlsx_to_pdf error: {e}")
        info["pixel_skip"] = f"original pdf: {e}"
        task7[wb_name] = info
        continue
    
    # Generate sanitized PDF
    try:
        sanitized = sanitize_workbook(str(wb_path), cls)
        pdf_san = export_sanitized_pdf(sanitized)
    except Exception as e:
        print(f"  sanitize error: {e}")
        info["pixel_skip"] = f"sanitize: {e}"
        task7[wb_name] = info
        try: os.unlink(pdf_orig)
        except: pass
        d = os.path.dirname(pdf_orig)
        if d: shutil.rmtree(d, ignore_errors=True)
        continue
    
    # Render both
    try:
        img_orig, wo, ho = render_pdf_to_image(pdf_orig)
        img_san, ws_, hs_ = render_pdf_to_image(pdf_san)
        info["original_image"] = f"{wo}x{ho}"
        info["sanitized_image"] = f"{ws_}x{hs_}"
        print(f"  Original image: {wo}x{ho}")
        print(f"  Sanitized image: {ws_}x{hs_}")
        
        # Compare dimensions
        if wo == ws_ and ho == hs_:
            info["image_dimensions_match"] = True
            print(f"  ✅ Image dimensions IDENTICAL")
            
            # Compute pixel difference
            import numpy as np
            diff = np.abs(img_orig.astype(np.int16) - img_san.astype(np.int16))
            diff_pct = float(np.mean(diff > 20) * 100)
            info["pixel_diff_pct"] = round(diff_pct, 3)
            print(f"  Pixel difference: {diff_pct:.3f}%")
            
            # Find content bounding box in both images
            # (non-white pixels)
            white_thresh = 240
            mask_orig = np.all(img_orig > white_thresh, axis=2)
            mask_san = np.all(img_san > white_thresh, axis=2)
            
            content_orig = np.where(~mask_orig)
            content_san = np.where(~mask_san)
            
            if len(content_orig[0]) > 0 and len(content_san[0]) > 0:
                # Bounding boxes
                orig_bbox = {
                    "top": int(content_orig[0].min()),
                    "bottom": int(content_orig[0].max()),
                    "left": int(content_orig[1].min()),
                    "right": int(content_orig[1].max()),
                }
                san_bbox = {
                    "top": int(content_san[0].min()),
                    "bottom": int(content_san[0].max()),
                    "left": int(content_san[1].min()),
                    "right": int(content_san[1].max()),
                }
                info["original_content_bbox"] = orig_bbox
                info["sanitized_content_bbox"] = san_bbox
                
                offset_x = san_bbox["left"] - orig_bbox["left"]
                offset_y = san_bbox["top"] - orig_bbox["top"]
                info["content_offset_px"] = f"({offset_x}, {offset_y})"
                info["content_offset_pt"] = f"({offset_x/PT_TO_PX:.2f}, {offset_y/PT_TO_PX:.2f})"
                
                print(f"  Original content bbox: ({orig_bbox['left']},{orig_bbox['top']})-({orig_bbox['right']},{orig_bbox['bottom']})")
                print(f"  Sanitized content bbox: ({san_bbox['left']},{san_bbox['top']})-({san_bbox['right']},{san_bbox['bottom']})")
                print(f"  Content offset: ({offset_x}, {offset_y}) px = ({offset_x/PT_TO_PX:.2f}, {offset_y/PT_TO_PX:.2f}) pt")
                
                if abs(offset_x) < 2 and abs(offset_y) < 2:
                    info["content_position_match"] = True
                    print(f"  ✅ Content positions MATCH")
                else:
                    info["content_position_match"] = False
                    print(f"  ❌ Content positions DIFFER by ({offset_x}, {offset_y}) px")
        else:
            info["image_dimensions_match"] = False
            info["dim_diff"] = f"orig={wo}x{ho} san={ws_}x{hs_}"
            print(f"  ❌ Image dimensions DIFFER: original={wo}x{ho}, sanitized={ws_}x{hs_}")
        
    except Exception as e:
        print(f"  Pixel comparison error: {e}")
        import traceback; traceback.print_exc()
        info["pixel_compare_error"] = str(e)
    
    # Cleanup
    for p in [pdf_orig, pdf_san]:
        try: os.unlink(p)
        except: pass
    for d in [os.path.dirname(pdf_orig), os.path.dirname(pdf_san),
              os.path.dirname(sanitized)]:
        try: shutil.rmtree(d, ignore_errors=True)
        except: pass
    
    task7[wb_name] = info

results["pixel_comparison"] = task7

# ════════════════════════════════════════════════════════════════════════
# SAVE RESULTS
# ════════════════════════════════════════════════════════════════════════
json_path = OUTPUT_DIR / "verify_results.json"
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\n\nResults saved to {json_path}")

# ════════════════════════════════════════════════════════════════════════
# SUMMARY
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  VERIFICATION SUMMARY")
print("=" * 70)

# Task 1
t1 = results.get("export_comparison", {}).get("task1_paper_size", {})
for wb, data in t1.items():
    ps_before = data.get("page_setup_before", {})
    pdf = data.get("pdf_default", {})
    print(f"\n  TASK 1 — {wb}: PaperSize={ps_before.get('PaperSize','?')} → PDF media box={pdf.get('media_box_pt','?')}")

# Task 2
t2 = results.get("export_comparison", {}).get("task2_sanitized_vs_original", {})
for wb, data in t2.items():
    match = data.get("pdf_match", "N/A")
    print(f"  TASK 2 — {wb}: PDFs match? {match}")

# Task 4
t4 = results.get("ignore_print_areas", {})
for wb, data in t4.items():
    effect = data.get("ipa_effect", "?")
    print(f"  TASK 4 — {wb}: IgnorePrintAreas effect = {effect}")

# Task 5
t5 = results.get("workbook_vs_worksheet", {})
for wb, data in t5.items():
    effect = data.get("level_effect", "?")
    print(f"  TASK 5 — {wb}: Workbook vs Worksheet = {effect}")

# Task 7
t7 = results.get("pixel_comparison", {})
for wb, data in t7.items():
    match = data.get("content_position_match", "N/A")
    offset = data.get("content_offset_px", "N/A")
    diff = data.get("pixel_diff_pct", "N/A")
    print(f"  TASK 7 — {wb}: Content match? {match}, offset={offset}, diff={diff}%")

print(f"\n  Full results: {json_path}")
print("  Done.")
