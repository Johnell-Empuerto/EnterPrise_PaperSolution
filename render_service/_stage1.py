"""
Stage 1: Cluster identification and PDF dimension comparison.
"""
import sys, os, gc, json, tempfile, shutil
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

_ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
sys.path.insert(0, _ROOT)

DOCS = r"C:\Users\MCF-JO~1\Documents"
FORM = os.path.join(DOCS, "FormTest - Copy.xlsx")
JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")

results = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n{'='*60}")
    print(f"  WORKBOOK: {wb_name}")
    print(f"{'='*60}")
    
    if not os.path.exists(wb_path):
        print(f"  FILE NOT FOUND: {wb_path}")
        continue
    
    # ── IDENTIFY CLUSTERS ──
    from render_service.upload_coordinate_generator import _identify_clusters
    print(f"  Identifying clusters...")
    gc.collect()
    try:
        cls = _identify_clusters(wb_path)
        print(f"  FOUND {len(cls)} clusters:")
        for c in sorted(cls, key=lambda x: x.get("cellAddr","")):
            print(f"    {c['name']:20s}  {c['cellAddr']:15s}  {c['type']}")
        results[wb_name] = {"clusters": len(cls), "cluster_list": cls}
    except Exception as e:
        print(f"  CLUSTER ERROR: {e}")
        import traceback; traceback.print_exc()
        continue
    
    # ── PDF DIMENSIONS (original workbook) ──
    from render_service.pdf_converter import xlsx_to_pdf
    from render_service.background_renderer import get_page_dimensions
    from render_service.upload_coordinate_generator import DPI
    
    print(f"  Generating PDF from ORIGINAL workbook...")
    gc.collect()
    try:
        pdf_orig = xlsx_to_pdf(wb_path)
        w_orig, h_orig = get_page_dimensions(pdf_orig, dpi=DPI)
        print(f"  ORIGINAL PDF: {w_orig}x{h_orig} px ({w_orig/(DPI/72):.0f}x{h_orig/(DPI/72):.0f} pt)")
        results[wb_name]["pdf_orig"] = f"{w_orig}x{h_orig}"
        
        # Cleanup
        try: os.unlink(pdf_orig)
        except: pass
        d = os.path.dirname(pdf_orig)
        if d: shutil.rmtree(d, ignore_errors=True)
    except Exception as e:
        print(f"  ORIGINAL PDF ERROR: {e}")
        import traceback; traceback.print_exc()
    
    # ── PDF DIMENSIONS (sanitized workbook) ──
    from render_service.upload_coordinate_generator import sanitize_workbook, export_sanitized_pdf, render_pdf_to_image
    
    print(f"  Sanitizing workbook...")
    gc.collect()
    try:
        sanitized = sanitize_workbook(wb_path, cls)
        print(f"  Generating PDF from SANITIZED workbook...")
        gc.collect()
        pdf_san = export_sanitized_pdf(sanitized)
        img, w_san, h_san = render_pdf_to_image(pdf_san)
        print(f"  SANITIZED PDF:  {w_san}x{h_san} px ({w_san/(DPI/72):.0f}x{h_san/(DPI/72):.0f} pt)")
        results[wb_name]["pdf_san"] = f"{w_san}x{h_san}"
        
        # Compare
        if "pdf_orig" in results[wb_name]:
            w_orig_r, h_orig_r = [int(x) for x in results[wb_name]["pdf_orig"].split("x")]
            if w_orig_r == w_san and h_orig_r == h_san:
                print(f"  ✅ SAME dimensions: original and sanitized PDFs match")
                results[wb_name]["pdf_match"] = True
            else:
                print(f"  ❌ DIFFERENT dimensions!")
                results[wb_name]["pdf_match"] = False
    except Exception as e:
        print(f"  SANITIZE/PDF ERROR: {e}")
        import traceback; traceback.print_exc()
    
    # Cleanup sanitized temp dirs
    try:
        for t in [os.path.dirname(sanitized), os.path.dirname(pdf_san)]:
            try: shutil.rmtree(t, ignore_errors=True)
            except: pass
    except: pass

# ── SUMMARY ──
print(f"\n{'='*60}")
print(f"  SUMMARY")
print(f"{'='*60}")
for wb_name, data in results.items():
    print(f"  {wb_name}:")
    print(f"    Clusters: {data.get('clusters', '?')}")
    print(f"    Original PDF: {data.get('pdf_orig', '?')}")
    print(f"    Sanitized PDF: {data.get('pdf_san', '?')}")
    print(f"    PDFs match: {data.get('pdf_match', '?')}")

# Save as JSON
json_path = os.path.join(_ROOT, "render_service", "_stage1_results.json")
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\nResults saved to {json_path}")
