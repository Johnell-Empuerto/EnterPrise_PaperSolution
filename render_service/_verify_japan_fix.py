"""
Standalone verification for the Japanese workbook.
Uses separate COM sessions for each stage, with GC cleanup between them.
"""
import sys, os, gc, json, tempfile, shutil
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
DOCS = r"C:\Users\MCF-JO~1\Documents"
sys.path.insert(0, ROOT)
os.chdir(ROOT)

JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")
DPI_VAL = 300
PT_TO_PX = DPI_VAL / 72.0

import win32com.client
import numpy as np
import fitz

def pdf_info(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    doc.close()
    return {"media_box_pt": f"{w:.1f}x{h:.1f}", "width_pt": round(w, 1), "height_pt": round(h, 1)}

def render_pdf(pdf_path):
    zoom = DPI_VAL / 72.0
    doc = fitz.open(pdf_path)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    doc.close()
    return img, pix.width, pix.height

print("=" * 70)
print("  JAPANESE WORKBOOK — STANDALONE VERIFICATION")
print("=" * 70)

# ── STEP 1: Identify clusters (fresh COM) ──
print("\n[1] Identifying clusters...")
from render_service.upload_coordinate_generator import _identify_clusters
gc.collect()
cls = _identify_clusters(JAPAN)
print(f"    Found {len(cls)} clusters")
for c in cls:
    print(f"      {c['name']:20s}  {c['cellAddr']:15s}  {c['type']}")

# ── STEP 2: Sanitize workbook (fresh COM) ──
print("\n[2] Sanitizing workbook...")
from render_service.upload_coordinate_generator import sanitize_workbook
gc.collect()
gc.collect()
sanitized = sanitize_workbook(JAPAN, cls)
print(f"    Sanitized -> {sanitized}")

# ── STEP 3: Generate original PDF via xlsx_to_pdf ──
print("\n[3] Generating original PDF...")
from render_service.pdf_converter import xlsx_to_pdf
gc.collect()
pdf_orig = xlsx_to_pdf(JAPAN)
print(f"    Original PDF dimensions: {pdf_info(pdf_orig)['media_box_pt']}")

# ── STEP 4: Generate sanitized PDF ──
print("\n[4] Generating sanitized PDF via export_sanitized_pdf...")
from render_service.upload_coordinate_generator import export_sanitized_pdf

# Try with retry
pdf_san = None
for attempt in range(3):
    try:
        gc.collect()
        gc.collect()
        pdf_san = export_sanitized_pdf(sanitized)
        print(f"    Sanitized PDF dimensions: {pdf_info(pdf_san)['media_box_pt']}")
        break
    except Exception as e:
        print(f"    Attempt {attempt+1} failed: {e}")
        gc.collect()
        if attempt < 2:
            import time
            time.sleep(2)

if pdf_san is None:
    print("    ❌ All attempts failed for export_sanitized_pdf")
    print("    Falling back: generating sanitized PDF via direct COM...")
    
    # Manual fallback: open the sanitized workbook and export directly
    for attempt in range(3):
        try:
            gc.collect()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            try:
                wb = excel.Workbooks.Open(os.path.abspath(sanitized))
                ws = wb.ActiveSheet
                # delete _Fields
                for i in range(wb.Sheets.Count, 0, -1):
                    try:
                        if wb.Sheets(i).Name in ("_Fields", "_RawData"):
                            wb.Sheets(i).Delete()
                    except: pass
                tmp = tempfile.mkdtemp(prefix="ple_japan_fallback_")
                pdf_san = os.path.join(tmp, "fallback.pdf")
                ws.ExportAsFixedFormat(0, os.path.abspath(pdf_san), 0, 0, True)
                wb.Close(False)
                print(f"    Fallback succeeded: {pdf_info(pdf_san)['media_box_pt']}")
                break
            finally:
                try: excel.Quit()
                except: pass
            gc.collect()
        except Exception as e2:
            print(f"    Fallback attempt {attempt+1} failed: {e2}")
            gc.collect()
            import time
            time.sleep(2)

# ── STEP 5: Pixel comparison ──
if pdf_san and os.path.exists(pdf_san):
    print("\n[5] Pixel comparison...")
    try:
        img_orig, wo, ho = render_pdf(pdf_orig)
        img_san, ws_, hs_ = render_pdf(pdf_san)
        print(f"    Original: {wo}x{ho}  Sanitized: {ws_}x{hs_}")
        
        if wo == ws_ and ho == hs_:
            print(f"    ✅ Image dimensions match")
            
            # Content bboxes
            wh = 240
            m_orig = ~np.all(img_orig > wh, axis=2)
            m_san = ~np.all(img_san > wh, axis=2)
            co = np.where(m_orig)
            cs_ = np.where(m_san)
            
            if len(co[0]) > 0 and len(cs_[0]) > 0:
                ox = int(cs_[1].min()) - int(co[1].min())
                oy = int(cs_[0].min()) - int(co[0].min())
                
                print(f"    Original bbox: ({int(co[1].min())},{int(co[0].min())})-({int(co[1].max())},{int(co[0].max())})")
                print(f"    Sanitized bbox: ({int(cs_[1].min())},{int(cs_[0].min())})-({int(cs_[1].max())},{int(cs_[0].max())})")
                print(f"    Content offset: ({ox}, {oy}) px = ({ox/PT_TO_PX:.2f}, {oy/PT_TO_PX:.2f}) pt")
                
                diff = np.abs(img_orig.astype(np.int16) - img_san.astype(np.int16))
                diff_pct = float(np.mean(np.any(diff > 20, axis=2)) * 100)
                print(f"    Pixel diff: {diff_pct:.3f}%")
                
                if abs(ox) < 10 and abs(oy) < 10:
                    print(f"    ✅ CRITICAL: Content positions MATCH (offset < 10px)")
                    print(f"       The fix WORKED for the Japanese workbook!")
                else:
                    print(f"    ⚠️ Content positions STILL differ by ({ox}, {oy}) px")
                    print(f"       BEFORE fix: (1593, 600) px")
                    print(f"       AFTER  fix: ({ox}, {oy}) px")
        else:
            print(f"    ❌ Different dimensions: orig={wo}x{ho} san={ws_}x{hs_}")
    except Exception as e:
        print(f"    Pixel comparison error: {e}")
        import traceback; traceback.print_exc()
else:
    print("\n[5] Pixel comparison SKIPPED — no sanitized PDF available")

# ── STEP 6: Generate coordinates ──
print("\n[6] Running generate_coordinates()...")
gc.collect()
gc.collect()
try:
    from render_service.upload_coordinate_generator import generate_coordinates
    fields = generate_coordinates(JAPAN)
    print(f"    Generated {len(fields)} fields:")
    for f in fields:
        print(f"      {f['name']:20s}  {f['cellAddr']:15s}  L={f['left_ratio']:.5f} T={f['top_ratio']:.5f} R={f['right_ratio']:.5f} B={f['bottom_ratio']:.5f}")
except Exception as e:
    print(f"    ERROR: {e}")
    import traceback; traceback.print_exc()

# Cleanup
print("\n  Cleaning up...")
for p in [pdf_orig, pdf_san]:
    try: os.unlink(p) if p and os.path.exists(p) else None
    except: pass
for d in [os.path.dirname(pdf_orig)]:
    try: shutil.rmtree(d, ignore_errors=True) if os.path.exists(d) else None
    except: pass
try:
    sd = os.path.dirname(sanitized)
    shutil.rmtree(sd, ignore_errors=True)
except: pass

print("\n  Done.")
