"""
Phase X.1 — Root Cause Verification Investigation
Each task runs in its own COM session. No shared state.
Writes structured results to JSON.
"""
import sys, os, gc, json, tempfile, shutil
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
DOCS = r"C:\Users\MCF-JO~1\Documents"
sys.path.insert(0, ROOT)
os.chdir(ROOT)
DPI_VAL = 300
PT_TO_PX = DPI_VAL / 72.0
OUTPUT = os.path.join(ROOT, "debug_output", "verify_full_results.json")

FORM = os.path.join(DOCS, "FormTest - Copy.xlsx")
JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")

import win32com.client
import fitz

def pdf_info(pdf_path):
    """Get PDF page dimensions."""
    doc = fitz.open(pdf_path)
    if len(doc) == 0:
        doc.close()
        return {"pages": 0, "error": "no pages"}
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    info = {
        "pages": len(doc),
        "media_box_pt": f"{w:.1f}x{h:.1f}",
        "width_pt": round(w, 1),
        "height_pt": round(h, 1),
    }
    doc.close()
    return info

# ════════════════════════════════════════════════════════════════════════
# TASK 1: PaperSize → PDF media box trace
# ════════════════════════════════════════════════════════════════════════
print("\n=== TASK 1: PaperSize → PDF media box ===")
task1 = {}

for name, path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  {name}:")
    info = {}
    if not os.path.exists(path):
        print(f"    NOT FOUND: {path}")
        task1[name] = {"error": "not found"}
        continue
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        
        ps = ws.PageSetup
        info["sheet"] = ws.Name
        info["PaperSize"] = int(ps.PaperSize)  # 1=Letter, 9=A4
        info["Orientation"] = int(ps.Orientation)
        info["Zoom"] = str(ps.Zoom)
        info["FitToPagesWide"] = int(getattr(ps, "FitToPagesWide", 0))
        info["FitToPagesTall"] = int(getattr(ps, "FitToPagesTall", 0))
        info["LeftMargin"] = round(float(ps.LeftMargin), 2)
        info["RightMargin"] = round(float(ps.RightMargin), 2)
        info["TopMargin"] = round(float(ps.TopMargin), 2)
        info["BottomMargin"] = round(float(ps.BottomMargin), 2)
        info["CenterHorizontally"] = bool(ps.CenterHorizontally)
        info["CenterVertically"] = bool(ps.CenterVertically)
        info["PrintArea"] = str(ps.PrintArea) if ps.PrintArea else "None"
        
        # Page dimensions - try multiple methods
        for method in ["PageWidth", "PageHeight"]:
            try:
                val = getattr(ps, method)
                info[f"ps_{method}"] = round(float(val), 1)
            except:
                info[f"ps_{method}"] = "ERROR"
        for method in ["PaperWidth", "PaperHeight"]:
            try:
                val = getattr(ps, method)
                info[f"ps_{method}"] = round(float(val), 1)
            except:
                info[f"ps_{method}"] = "ERROR"
        
        # Export to PDF (default params, no modification)
        tmp = tempfile.mkdtemp(prefix="ple_task1_")
        pdf = os.path.join(tmp, "out.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf), 0, 0, False)
        info["pdf"] = pdf_info(pdf)
        
        wb.Close(False)
        
        print(f"    Sheet: {info['sheet']}")
        print(f"    PaperSize COM: {info['PaperSize']} ({'Letter' if info['PaperSize']==1 else 'A4' if info['PaperSize']==9 else 'other'})")
        print(f"    Zoom: {info['Zoom']}, FitToPages: {info['FitToPagesWide']}x{info['FitToPagesTall']}")
        print(f"    Margins L/R/T/B: {info['LeftMargin']}/{info['RightMargin']}/{info['TopMargin']}/{info['BottomMargin']}")
        print(f"    Center H/V: {info['CenterHorizontally']}/{info['CenterVertically']}")
        print(f"    PDF media box: {info['pdf']['media_box_pt']} ({info['pdf']['pages']} page(s))")
        print(f"    PageWidth: {info.get('ps_PageWidth','?')}, PageHeight: {info.get('ps_PageHeight','?')}")
        print(f"    PaperWidth: {info.get('ps_PaperWidth','?')}, PaperHeight: {info.get('ps_PaperHeight','?')}")
        
        # Cleanup
        try: shutil.rmtree(tmp, ignore_errors=True)
        except: pass
        
    except Exception as e:
        print(f"    ERROR: {e}")
        import traceback; traceback.print_exc()
        info["error"] = str(e)
    finally:
        try: excel.Quit()
        except: pass
    gc.collect()
    
    task1[name] = info

# ════════════════════════════════════════════════════════════════════════
# TASK 3: PageSetup before and after sanitize modifications
# ════════════════════════════════════════════════════════════════════════
print("\n=== TASK 3: PageSetup before/after sanitize modifications ===")
task3 = {}

for name, path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  {name}:")
    info = {}
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        info["sheet"] = ws.Name
        
        before = {}
        after = {}
        
        props_to_check = ["Zoom", "FitToPagesWide", "FitToPagesTall",
                          "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
                          "CenterHorizontally", "CenterVertically",
                          "Orientation", "Order", "PaperSize",
                          "PrintArea", "PrintTitleRows", "PrintTitleColumns",
                          "FirstPageNumber", "PageNumberLocal",
                          "PageSetup.PageWidth", "PageSetup.PageHeight"]
        
        # Read BEFORE
        for p in props_to_check:
            try:
                if "." in p:
                    parts = p.split(".")
                    obj = ws.PageSetup
                    for part in parts:
                        obj = getattr(obj, part)
                    before[p] = repr(obj)
                else:
                    before[p] = repr(getattr(ws.PageSetup, p))
            except Exception as e:
                before[p] = f"ERR({e})"
        
        # Apply sanitize's modifications
        try: ws.PageSetup.Zoom = False
        except: pass
        try: ws.PageSetup.FitToPagesWide = False
        except: pass
        try: ws.PageSetup.FitToPagesTall = False
        except: pass
        
        # Read AFTER
        for p in props_to_check:
            try:
                if "." in p:
                    parts = p.split(".")
                    obj = ws.PageSetup
                    for part in parts:
                        obj = getattr(obj, part)
                    after[p] = repr(obj)
                else:
                    after[p] = repr(getattr(ws.PageSetup, p))
            except Exception as e:
                after[p] = f"ERR({e})"
        
        info["before"] = before
        info["after"] = after
        
        # Print comparison
        print(f"    {'Property':30s} {'Before':20s} {'After':20s}")
        print(f"    {'-'*30} {'-'*20} {'-'*20}")
        for key in sorted(before.keys()):
            b = before.get(key, "?")
            a = after.get(key, "?")
            ch = "⚠️" if b != a else "  "
            print(f"    {key:30s} {b:20s} {a:20s} {ch}")
        
        wb.Close(False)
    except Exception as e:
        print(f"    ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        try: excel.Quit()
        except: pass
    gc.collect()
    task3[name] = info

# ════════════════════════════════════════════════════════════════════════
# TASK 4: IgnorePrintAreas effect
# ════════════════════════════════════════════════════════════════════════
print("\n=== TASK 4: IgnorePrintAreas effect ===")
task4 = {}

for name, path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  {name}:")
    info = {}
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        
        info["sheet"] = ws.Name
        info["PrintArea"] = str(ws.PageSetup.PrintArea) if ws.PageSetup.PrintArea else "None"
        
        # Export with IgnorePrintAreas=False
        tmp1 = tempfile.mkdtemp(prefix="ple_ipaF_")
        p1 = os.path.join(tmp1, "ignorepa_false.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(p1), 0, 0, False)
        info["ignorePA_false"] = pdf_info(p1)
        
        # Export with IgnorePrintAreas=True
        tmp2 = tempfile.mkdtemp(prefix="ple_ipaT_")
        p2 = os.path.join(tmp2, "ignorepa_true.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(p2), 0, 0, True)
        info["ignorePA_true"] = pdf_info(p2)
        
        i1 = info["ignorePA_false"]
        i2 = info["ignorePA_true"]
        same_media = i1["media_box_pt"] == i2["media_box_pt"]
        same_pages = i1["pages"] == i2["pages"]
        info["identical"] = same_media and same_pages
        if not info["identical"]:
            info["diff"] = f"false={i1['media_box_pt']}({i1['pages']}p) true={i2['media_box_pt']}({i2['pages']}p)"
        
        print(f"    PrintArea: {info['PrintArea']}")
        print(f"    IgnorePA=False: {i1['media_box_pt']} ({i1['pages']}p)")
        print(f"    IgnorePA=True:  {i2['media_box_pt']} ({i2['pages']}p)")
        print(f"    Identical: {info['identical']}")
        
        wb.Close(False)
        for d in [tmp1, tmp2]:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
    except Exception as e:
        print(f"    ERROR: {e}")
        import traceback; traceback.print_exc()
    finally:
        try: excel.Quit()
        except: pass
    gc.collect()
    task4[name] = info

# ════════════════════════════════════════════════════════════════════════
# TASK 5: Workbook-level vs Worksheet-level export
# ════════════════════════════════════════════════════════════════════════
print("\n=== TASK 5: Workbook-level vs Worksheet-level export ===")
task5 = {}

for name, path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  {name}:")
    info = {}
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        ws = wb.Sheets(1)
        if ws.Name == "_Fields":
            ws = wb.Sheets(2)
        
        # Workbook-level export
        tmp1 = tempfile.mkdtemp(prefix="ple_wblvl_")
        p1 = os.path.join(tmp1, "workbook_level.pdf")
        wb.ExportAsFixedFormat(0, os.path.abspath(p1))
        info["workbook_level"] = pdf_info(p1)
        
        # Worksheet-level export
        tmp2 = tempfile.mkdtemp(prefix="ple_wslvl_")
        p2 = os.path.join(tmp2, "worksheet_level.pdf")
        ws.ExportAsFixedFormat(0, os.path.abspath(p2), 0, 0, False)
        info["worksheet_level"] = pdf_info(p2)
        
        i1 = info["workbook_level"]
        i2 = info["worksheet_level"]
        same_media = i1["media_box_pt"] == i2["media_box_pt"]
        same_pages = i1["pages"] == i2["pages"]
        info["identical"] = same_media and same_pages
        if not info["identical"]:
            info["diff"] = f"wb={i1['media_box_pt']}({i1['pages']}p) ws={i2['media_box_pt']}({i2['pages']}p)"
        
        print(f"    Workbook-level: {i1['media_box_pt']} ({i1['pages']}p)")
        print(f"    Worksheet-level: {i2['media_box_pt']} ({i2['pages']}p)")
        print(f"    Identical: {info['identical']}")
        
        wb.Close(False)
        for d in [tmp1, tmp2]:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
    except Exception as e:
        print(f"    ERROR: {e}")
    finally:
        try: excel.Quit()
        except: pass
    gc.collect()
    task5[name] = info

# ════════════════════════════════════════════════════════════════════════
# TASK 2+7: Sanitized vs original PDF comparison + pixel comparison
# ════════════════════════════════════════════════════════════════════════
print("\n=== TASK 2+7: Sanitized vs original — PDF geometry + pixel comparison ===")

from render_service.upload_coordinate_generator import (
    _identify_clusters, sanitize_workbook, export_sanitized_pdf,
    render_pdf_to_image
)
from render_service.pdf_converter import xlsx_to_pdf

task27 = {}

for name, path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  {name}:")
    info = {}
    
    # Get clusters
    try:
        cls = _identify_clusters(path)
        info["cluster_count"] = len(cls)
        print(f"    Clusters: {len(cls)}")
        if len(cls) == 0:
            print(f"    No clusters — skipping PDF comparison")
            task27[name] = {"error": "no clusters"}
            continue
    except Exception as e:
        print(f"    Cluster error: {e}")
        task27[name] = {"error": str(e)}
        continue
    
    # Generate original PDF
    try:
        pdf_orig = xlsx_to_pdf(path)
        info["original_pdf"] = pdf_info(pdf_orig)
        print(f"    Original PDF: {info['original_pdf']['media_box_pt']} ({info['original_pdf']['pages']}p)")
    except Exception as e:
        print(f"    xlsx_to_pdf error: {e}")
        import traceback; traceback.print_exc()
        task27[name] = {"error": f"original pdf: {e}"}
        continue
    
    # Generate sanitized PDF
    try:
        sanitized = sanitize_workbook(path, cls)
        pdf_san = export_sanitized_pdf(sanitized)
        info["sanitized_pdf"] = pdf_info(pdf_san)
        print(f"    Sanitized PDF: {info['sanitized_pdf']['media_box_pt']} ({info['sanitized_pdf']['pages']}p)")
    except Exception as e:
        print(f"    Sanitize error: {e}")
        info["sanitized_error"] = str(e)
        # Still proceed with original PDF render
        pdf_san = None
    
    # Render both to images and compare pixels
    try:
        img_orig, wo, ho = render_pdf_to_image(pdf_orig)
        info["orig_img"] = f"{wo}x{ho}"
        
        if pdf_san and os.path.exists(pdf_san):
            try:
                img_san, ws_, hs_ = render_pdf_to_image(pdf_san)
                info["san_img"] = f"{ws_}x{hs_}"
                
                # Compare dimensions
                if wo == ws_ and ho == hs_:
                    info["img_match"] = True
                    print(f"    ✅ Images are {wo}x{ho}")
                    
                    # Pixel-level comparison
                    import numpy as np
                    diff = np.abs(img_orig.astype(np.int16) - img_san.astype(np.int16))
                    diff_pct = float(np.mean(np.any(diff > 20, axis=2)) * 100)
                    info["pixel_diff_pct"] = round(diff_pct, 3)
                    
                    # Check content bounding boxes (non-white pixels)
                    wh = 240
                    m_orig = ~np.all(img_orig > wh, axis=2)
                    m_san = ~np.all(img_san > wh, axis=2)
                    
                    co = np.where(m_orig)
                    cs_ = np.where(m_san)
                    
                    if len(co[0]) > 0 and len(cs_[0]) > 0:
                        obb = {"t": int(co[0].min()), "b": int(co[0].max()),
                               "l": int(co[1].min()), "r": int(co[1].max())}
                        sbb = {"t": int(cs_[0].min()), "b": int(cs_[0].max()),
                               "l": int(cs_[1].min()), "r": int(cs_[1].max())}
                        info["orig_bbox"] = obb
                        info["san_bbox"] = sbb
                        ox = sbb["l"] - obb["l"]
                        oy = sbb["t"] - obb["t"]
                        info["content_offset_px"] = f"({ox}, {oy})"
                        print(f"    Original content: ({obb['l']},{obb['t']})-({obb['r']},{obb['b']})")
                        print(f"    Sanitized content: ({sbb['l']},{sbb['t']})-({sbb['r']},{sbb['b']})")
                        print(f"    Content offset: ({ox}, {oy}) px = ({ox/PT_TO_PX:.2f}, {oy/PT_TO_PX:.2f}) pt")
                        info["content_match"] = abs(ox) < 2 and abs(oy) < 2
                        print(f"    Pixel diff: {diff_pct:.3f}%")
                    else:
                        info["content_bbox_error"] = "no content found in one image"
                        print(f"    ⚠️ Could not compute content bounding box")
                else:
                    info["img_match"] = False
                    info["img_diff"] = f"orig={wo}x{ho} san={ws_}x{hs_}"
                    print(f"    ❌ Image dimensions differ: orig={wo}x{ho} san={ws_}x{hs_}")
            except Exception as e:
                print(f"    Sanitized render error: {e}")
                info["san_render_error"] = str(e)
    except Exception as e:
        print(f"    Render/comparison error: {e}")
        import traceback; traceback.print_exc()
        info["comparison_error"] = str(e)
    
    # Cleanup
    for p in [pdf_orig, pdf_san]:
        try: os.unlink(p) if p and os.path.exists(p) else None
        except: pass
    for d in [os.path.dirname(pdf_orig), 
              os.path.dirname(pdf_san) if pdf_san else None,
              os.path.dirname(sanitized) if 'sanitized' in dir() else None]:
        try: shutil.rmtree(d, ignore_errors=True) if d and os.path.exists(d) else None
        except: pass
    
    task27[name] = info

# ════════════════════════════════════════════════════════════════════════
# SAVE ALL RESULTS
# ════════════════════════════════════════════════════════════════════════
all_results = {
    "task1_paper_size_pdf": task1,
    "task3_pagesetup_before_after": task3,
    "task4_ignore_print_areas": task4,
    "task5_workbook_vs_worksheet": task5,
    "task27_sanitized_vs_original_pixel": task27,
}

with open(OUTPUT, 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)
print(f"\n\nAll results saved to: {OUTPUT}")

# ════════════════════════════════════════════════════════════════════════
# FINAL SUMMARY
# ════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("  VERIFICATION SUMMARY")
print("=" * 70)

for section, data in all_results.items():
    print(f"\n  {section}:")
    for wb, d in data.items():
        if "error" in d:
            print(f"    {wb}: ❌ ERROR - {d['error']}")
            continue
        if section == "task1_paper_size_pdf":
            ps = d.get("PaperSize", "?")
            pdf = d.get("pdf", {})
            print(f"    {wb}: PaperSize={ps}, PDF={pdf.get('media_box_pt','?')}")
        elif section == "task4_ignore_print_areas":
            print(f"    {wb}: Identical={d.get('identical','?')}")
        elif section == "task5_workbook_vs_worksheet":
            print(f"    {wb}: Identical={d.get('identical','?')}")
        elif section == "task27_sanitized_vs_original_pixel":
            match = d.get("img_match", "N/A")
            cm = d.get("content_match", "N/A")
            off = d.get("content_offset_px", "N/A")
            dp = d.get("pixel_diff_pct", "N/A")
            print(f"    {wb}: ImgMatch={match}, ContentMatch={cm}, Offset={off}, Diff%={dp}")

print("\n  Done.")
