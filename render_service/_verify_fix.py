"""
Phase X.2 — Fix verification.
Compares sanitized vs original PDF geometry after the export_sanitized_pdf fix.
Also runs full coordinate generation on both workbooks.
"""
import sys, os, gc, json, tempfile, shutil
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
DOCS = r"C:\Users\MCF-JO~1\Documents"
sys.path.insert(0, ROOT)
os.chdir(ROOT)

DPI_VAL = 300
PT_TO_PX = DPI_VAL / 72.0
OUTPUT_JSON = os.path.join(ROOT, "debug_output", "verify_fix_results.json")

FORM = os.path.join(DOCS, "FormTest - Copy.xlsx")
JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")

from render_service.upload_coordinate_generator import (
    _identify_clusters, sanitize_workbook, export_sanitized_pdf,
    render_pdf_to_image, generate_coordinates
)
from render_service.pdf_converter import xlsx_to_pdf

import fitz
import numpy as np

def pdf_info(pdf_path):
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

results = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n{'='*70}")
    print(f"  WORKBOOK: {wb_name}")
    print(f"{'='*70}")
    wb_result = {}
    
    if not os.path.exists(wb_path):
        print(f"  NOT FOUND: {wb_path}")
        wb_result["error"] = "not found"
        results[wb_name] = wb_result
        continue
    
    # ── 1. Get clusters ──
    print(f"\n  [1] Identifying clusters...")
    try:
        cls = _identify_clusters(wb_path)
        wb_result["cluster_count"] = len(cls)
        print(f"      Found {len(cls)} clusters")
    except Exception as e:
        print(f"      ERROR: {e}")
        wb_result["error"] = f"cluster: {e}"
        results[wb_name] = wb_result
        continue
    
    # ── 2. Generate original PDF ──
    print(f"  [2] Generating original PDF...")
    try:
        pdf_orig = xlsx_to_pdf(wb_path)
        orig_info = pdf_info(pdf_orig)
        wb_result["original_pdf"] = orig_info
        print(f"      {orig_info['media_box_pt']} ({orig_info['pages']}p)")
    except Exception as e:
        print(f"      ERROR: {e}")
        wb_result["error"] = f"orig_pdf: {e}"
        results[wb_name] = wb_result
        continue
    
    # ── 3. Sanitize workbook ──
    print(f"  [3] Sanitizing workbook (WITH fix)...")
    try:
        sanitized = sanitize_workbook(wb_path, cls)
        print(f"      Sanitized OK")
    except Exception as e:
        print(f"      ERROR: {e}")
        wb_result["error"] = f"sanitize: {e}"
        results[wb_name] = wb_result
        # Cleanup
        try: os.unlink(pdf_orig)
        except: pass
        d = os.path.dirname(pdf_orig)
        if d: shutil.rmtree(d, ignore_errors=True)
        continue
    
    # ── 4. Generate sanitized PDF (WITH FIX) ──
    print(f"  [4] Generating sanitized PDF (WITH FIX)...")
    try:
        pdf_san = export_sanitized_pdf(sanitized)
        san_info = pdf_info(pdf_san)
        wb_result["sanitized_pdf"] = san_info
        print(f"      {san_info['media_box_pt']} ({san_info['pages']}p)")
    except Exception as e:
        print(f"      ERROR: {e}")
        wb_result["error"] = f"san_pdf: {e}"
        results[wb_name] = wb_result
        for p in [pdf_orig, None]:
            try: os.unlink(p)
            except: pass
        for d in [os.path.dirname(pdf_orig), os.path.dirname(sanitized)]:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
        continue
    
    # ── 5. Compare dimensions ──
    same_dims = (orig_info["media_box_pt"] == san_info["media_box_pt"] and
                 orig_info["pages"] == san_info["pages"])
    wb_result["dimensions_match"] = same_dims
    print(f"  [5] Dimensions match: {same_dims}")
    
    # ── 6. Render BOTH to images and compare pixels ──
    print(f"  [6] Rendering both to images...")
    try:
        img_orig, wo, ho = render_pdf_to_image(pdf_orig)
        img_san, ws_, hs_ = render_pdf_to_image(pdf_san)
        wb_result["orig_image"] = f"{wo}x{ho}"
        wb_result["san_image"] = f"{ws_}x{hs_}"
        
        if wo == ws_ and ho == hs_:
            wb_result["image_dimensions_match"] = True
            print(f"      Both: {wo}x{ho}")
            
            # Compare pixel content
            diff = np.abs(img_orig.astype(np.int16) - img_san.astype(np.int16))
            any_diff = np.any(diff > 20, axis=2)
            diff_pct = float(np.mean(any_diff) * 100)
            wb_result["pixel_diff_pct"] = round(diff_pct, 3)
            
            # Content bounding boxes (non-white pixels)
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
                wb_result["original_bbox"] = obb
                wb_result["sanitized_bbox"] = sbb
                
                ox = sbb["l"] - obb["l"]
                oy = sbb["t"] - obb["t"]
                wb_result["content_offset_px"] = f"({ox}, {oy})"
                wb_result["content_offset_pt"] = f"({ox/PT_TO_PX:.2f}, {oy/PT_TO_PX:.2f})"
                
                print(f"      Original content: ({obb['l']},{obb['t']})-({obb['r']},{obb['b']})")
                print(f"      Sanitized content: ({sbb['l']},{sbb['t']})-({sbb['r']},{sbb['b']})")
                print(f"      Content offset: ({ox}, {oy}) px = ({ox/PT_TO_PX:.2f}, {oy/PT_TO_PX:.2f}) pt")
                
                if abs(ox) < 5 and abs(oy) < 5:
                    wb_result["content_match"] = True
                    print(f"      ✅ Content positions MATCH (offset < 5 px)")
                else:
                    wb_result["content_match"] = False
                    print(f"      ⚠️ Content positions differ by ({ox}, {oy}) px")
                
                print(f"      Pixel diff: {diff_pct:.3f}%")
            else:
                print(f"      ⚠️ Could not compute bbox for one image")
                wb_result["bbox_error"] = "no content in one image"
        else:
            wb_result["image_dimensions_match"] = False
            wb_result["dim_diff"] = f"orig={wo}x{ho} san={ws_}x{hs_}"
            print(f"      ❌ Dimensions differ: orig={wo}x{ho} san={ws_}x{hs_}")
    except Exception as e:
        print(f"      Pixel comparison error: {e}")
        import traceback; traceback.print_exc()
        wb_result["pixel_error"] = str(e)
    
    # ── 7. Run full coordinate generation ──
    print(f"  [7] Running full generate_coordinates()...")
    try:
        fields = generate_coordinates(wb_path)
        wb_result["field_count"] = len(fields)
        wb_result["fields"] = []
        for f in fields:
            wb_result["fields"].append({
                "name": f["name"],
                "cellAddr": f["cellAddr"],
                "left_ratio": f["left_ratio"],
                "top_ratio": f["top_ratio"],
                "right_ratio": f["right_ratio"],
                "bottom_ratio": f["bottom_ratio"],
            })
        print(f"      Generated {len(fields)} fields:")
        for f in fields:
            print(f"        {f['name']:20s}  {f['cellAddr']:15s}  L={f['left_ratio']:.5f} T={f['top_ratio']:.5f} R={f['right_ratio']:.5f} B={f['bottom_ratio']:.5f}")
    except Exception as e:
        print(f"      ERROR: {e}")
        wb_result["generate_error"] = str(e)
    
    # ── Cleanup ──
    for p in [pdf_orig, pdf_san]:
        try: os.unlink(p)
        except: pass
    for d in [os.path.dirname(pdf_orig), os.path.dirname(pdf_san),
              os.path.dirname(sanitized)]:
        try: shutil.rmtree(d, ignore_errors=True)
        except: pass
    gc.collect()
    
    results[wb_name] = wb_result

# ── Save ──
with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\nResults saved to {OUTPUT_JSON}")

# ── Summary ──
print(f"\n{'='*70}")
print(f"  VERIFICATION SUMMARY")
print(f"{'='*70}")
for wb, r in results.items():
    print(f"\n  {wb}:")
    if "error" in r:
        print(f"    ❌ {r['error']}")
        continue
    print(f"    Clusters: {r.get('cluster_count', '?')}")
    print(f"    Dimensions match: {r.get('dimensions_match', '?')}")
    print(f"    Content offset: {r.get('content_offset_px', 'N/A')}")
    print(f"    Content match (<5px): {r.get('content_match', 'N/A')}")
    print(f"    Pixel diff: {r.get('pixel_diff_pct', 'N/A')}%")
    print(f"    Fields generated: {r.get('field_count', '?')}")
    if r.get('content_match') == True:
        print(f"    ✅ VERIFIED: Coordinates will align with background")
    elif r.get('content_match') == False:
        print(f"    ❌ MISMATCH: Coordinates may not align")
    else:
        print(f"    ⚠️ Unknown")

print(f"\n  Done.")
