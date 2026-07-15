"""
Coordinate Divergence Investigation — Phase X

Compares the Python pixel scan pipeline outputs for two workbooks:
  1. FormTest - Copy.xlsx (known to align)
  2. [V3.1_Sample]アンケート用紙.xlsx (known to NOT align)

Captures coordinate values at every stage to find the FIRST divergence point.

Usage:
    python render_service/debug_investigation.py

No code changes — evidence collection only.
"""

import sys, os, json, traceback, tempfile, shutil
# Force UTF-8 for stdout to handle Unicode paths
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────
# Try the full username path first, then the short 8.3 path
_USER = Path(r"C:\Users\MCF-JOHNELLEEMUERTO")
if not _USER.exists():
    _USER = Path(r"C:\Users\MCF-JO~1")
DOCS   = _USER / "Documents"
FORM   = DOCS / "FormTest - Copy.xlsx"
JAPAN  = DOCS / "[V3.1_Sample]アンケート用紙.xlsx"
OUTPUT = _USER / "Documents" / "Johnell" / "PaperLess Enterprise" / "docs" / "phaseX_divergence_report.md"

print(f"Using user path: {_USER}")
print(f"FormTest path: {FORM}")
print(f"Japan path: {JAPAN}")

# ── Setup module path ───────────────────────────────────────────────────
ROOT = _USER / "Documents" / "Johnell" / "PaperLess Enterprise"
sys.path.insert(0, str(ROOT))
os.chdir(str(ROOT))
print(f"Project root: {ROOT}")
print(f"sys.path[0]: {sys.path[0]}")
print(f"Python version: {sys.version}")

# ── Helpers ─────────────────────────────────────────────────────────────
def stage(name: str):
    print(f"\n{'='*70}")
    print(f"  STAGE: {name}")
    print(f"{'='*70}")

def info(msg: str):
    print(f"  [i] {msg}")

def field_section(fields, label: str):
    print(f"\n  --- {label} ---")
    for i, f in enumerate(fields):
        lr = f.get("left_ratio", "?")
        tr = f.get("top_ratio", "?")
        rr = f.get("right_ratio", "?")
        br = f.get("bottom_ratio", "?")
        addr = f.get("cellAddr", "?")
        name = f.get("name", f.get("id", f"field_{i}"))
        print(f"    {name:20s}  addr={addr:15s}  L={lr} T={tr} R={rr} B={br}")

def markdown_table(fields_form, fields_japan, label: str):
    """Print a markdown table comparing the same field index across both workbooks."""
    lines = []
    lines.append(f"### {label}")
    max_len = max(len(fields_form), len(fields_japan))
    lines.append("| Field | FormTest L | FormTest T | FormTest R | FormTest B | Japan L | Japan T | Japan R | Japan B |")
    lines.append("|-------|------------|------------|------------|------------|---------|---------|---------|---------|")
    for i in range(max_len):
        f_f = fields_form[i] if i < len(fields_form) else {}
        f_j = fields_japan[i] if i < len(fields_japan) else {}
        name = f_f.get("name", f_j.get("name", f"field_{i}"))
        fl = f_f.get("left_ratio", "—")
        ft = f_f.get("top_ratio", "—")
        fr = f_f.get("right_ratio", "—")
        fb = f_f.get("bottom_ratio", "—")
        jl = f_j.get("left_ratio", "—")
        jt = f_j.get("top_ratio", "—")
        jr = f_j.get("right_ratio", "—")
        jb = f_j.get("bottom_ratio", "—")
        lines.append(f"| {name:20s} | {fl} | {ft} | {fr} | {fb} | {jl} | {jt} | {jr} | {jb} |")
    lines.append("")
    return "\n".join(lines)

# ══════════════════════════════════════════════════════════════════════════
# STAGE 0: File existence
# ══════════════════════════════════════════════════════════════════════════
stage("0: File existence")

results = []  # Collects (stage_name, form_fields, japan_fields) for comparison

for name, path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    exists = path.exists()
    size = path.stat().st_size if exists else 0
    info(f"{name}: exists={exists}, size={size} bytes")
    if not exists:
        print(f"ERROR: {name} workbook not found at {path}")
        sys.exit(1)

# ══════════════════════════════════════════════════════════════════════════
# STAGE 1: Python pixel scan pipeline (generate_coordinates)
# ══════════════════════════════════════════════════════════════════════════
stage("1: Python pixel scan — generate_coordinates()")

from render_service.upload_coordinate_generator import (
    generate_coordinates, _identify_clusters,
    sanitize_workbook, export_sanitized_pdf,
    render_pdf_to_image, scan_black_rectangles,
    split_merged_rects, normalize_rects,
    _sort_key_meta, DPI
)

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  ═══ {wb_name} ═══")

    # Step 1a: Cluster identification
    info("1a: Identifying clusters...")
    cls = _identify_clusters(str(wb_path))
    info(f"Found {len(cls)} clusters")
    for c in cls:
        info(f"  {c['name']:20s}  addr={c['cellAddr']:15s}  type={c['type']}")

    # Step 1b: Sanitize workbook — wrap in fresh COM dispatch to avoid threading issues
    info("1b: Sanitizing workbook...")
    try:
        import gc
        gc.collect()  # Clean up any lingering COM objects
        sanitized = sanitize_workbook(str(wb_path), cls)
        info(f"Sanitized -> {sanitized}")
    except Exception as e:
        info(f"SANITIZE FAILED: {e}")
        traceback.print_exc()
        continue

    # Step 1c: Export sanitized PDF
    info("1c: Exporting sanitized PDF...")
    try:
        import gc
        gc.collect()
        pdf_path = export_sanitized_pdf(sanitized)
        info(f"PDF -> {pdf_path}")
    except Exception as e:
        info(f"PDF EXPORT FAILED: {e}")
        traceback.print_exc()
        continue

    # Step 1d: Render PDF to image
    info("1d: Rendering PDF to image...")
    try:
        img, img_w, img_h = render_pdf_to_image(pdf_path)
        info(f"Image: {img_w}x{img_h} px")
    except Exception as e:
        info(f"RENDER FAILED: {e}")
        traceback.print_exc()
        continue

    # Step 1e: Pixel scan
    info("1e: Scanning black rectangles...")
    try:
        rects = scan_black_rectangles(img)
        info(f"Found {len(rects)} rectangles from pixel scan")
        for r in rects:
            info(f"  rect: L={r['Left']:.1f} T={r['Top']:.1f} R={r['Right']:.1f} B={r['Bottom']:.1f}")
    except Exception as e:
        info(f"SCAN FAILED: {e}")
        traceback.print_exc()
        continue

    # Step 1f: Split merged rects
    info("1f: Splitting merged rects...")
    try:
        split = split_merged_rects(rects, cls)
        info(f"After split: {len(split)} rectangles")
        for r in split:
            info(f"  split: L={r['Left']:.1f} T={r['Top']:.1f} R={r['Right']:.1f} B={r['Bottom']:.1f}")
    except Exception as e:
        info(f"SPLIT FAILED: {e}")
        traceback.print_exc()
        continue

    # Step 1g: Normalize
    info("1g: Normalizing...")
    try:
        normalized = normalize_rects(split, img_w, img_h)
        info(f"Normalized: {len(normalized)} rects")
    except Exception as e:
        info(f"NORMALIZE FAILED: {e}")
        traceback.print_exc()
        continue

    # Step 1h: Final output
    info("1h: Merging with cluster metadata...")
    cls_sorted = sorted(cls, key=_sort_key_meta)
    norm_sorted = sorted(normalized, key=lambda r: (r["Top"], r["Left"]))
    
    final = []
    for meta, rect in zip(cls_sorted, norm_sorted[:len(cls_sorted)]):
        final.append({
            "name": meta["name"],
            "type": meta["type"],
            "cellAddr": meta["cellAddr"],
            "left_ratio": round(rect["left_ratio"], 7),
            "top_ratio": round(rect["top_ratio"], 7),
            "right_ratio": round(rect["right_ratio"], 7),
            "bottom_ratio": round(rect["bottom_ratio"], 7),
            # Pixel values for comparison
            "left_px": round(rect["left_ratio"] * img_w, 1),
            "top_px": round(rect["top_ratio"] * img_h, 1),
            "right_px": round(rect["right_ratio"] * img_w, 1),
            "bottom_px": round(rect["bottom_ratio"] * img_h, 1),
            "img_w": img_w,
            "img_h": img_h,
        })
    
    results.append(("1_python_pixel_scan", wb_name, final))
    info(f"Generated {len(final)} fields")
    for f in final:
        info(f"  {f['name']:20s}  L={f['left_ratio']:.5f} T={f['top_ratio']:.5f} R={f['right_ratio']:.5f} B={f['bottom_ratio']:.5f}  px: {f['left_px']:.0f},{f['top_px']:.0f},{f['right_px']:.0f},{f['bottom_px']:.0f}")

    # Cleanup temp dirs
    for tmp_dir in {os.path.dirname(sanitized), os.path.dirname(pdf_path)}:
        try: shutil.rmtree(tmp_dir, ignore_errors=True)
        except: pass

# ══════════════════════════════════════════════════════════════════════════
# STAGE 2: PDF export settings comparison
# ══════════════════════════════════════════════════════════════════════════
stage("2: PDF export settings comparison")

from render_service.pdf_converter import xlsx_to_pdf
from render_service.background_renderer import pdf_page_to_png, get_page_dimensions

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n  ═══ {wb_name} ── Export settings comparison ═══")
    
    # Render both PDFs (sanitized vs original) and compare dimensions
    try:
        # Run coordinate generation to get sanitized PDF stats
        cls = _identify_clusters(str(wb_path))
        sanitized = sanitize_workbook(str(wb_path), cls)
        pdf_sanitized = export_sanitized_pdf(sanitized)
        img_san, w_san, h_san = render_pdf_to_image(pdf_sanitized)
        info(f"Sanitized PDF: {w_san}x{h_san} px at {DPI} DPI ({w_san / (DPI/72):.0f}x{h_san / (DPI/72):.0f} pt)")
        
        # Original workbook PDF
        pdf_original = xlsx_to_pdf(str(wb_path))
        img_orig, w_orig, h_orig = render_pdf_to_image(pdf_original)
        info(f"Original PDF:  {w_orig}x{h_orig} px at {DPI} DPI ({w_orig / (DPI/72):.0f}x{h_orig / (DPI/72):.0f} pt)")
        
        # Compare dimensions
        if w_san == w_orig and h_san == h_orig:
            info("✅ Same page dimensions: PDFs match")
        else:
            info(f"❌ DIFFERENT page dimensions: sanitized={w_san}x{h_san}, original={w_orig}x{h_orig}")
            info(f"   Difference: {w_san - w_orig}px wide, {h_san - h_orig}px tall")
        
        # Check if the PDFs have the same number of pages
        import fitz
        doc_san = fitz.open(pdf_sanitized)
        doc_orig = fitz.open(pdf_original)
        info(f"Sanitized pages: {len(doc_san)}, Original pages: {len(doc_orig)}")
        if len(doc_san) != len(doc_orig):
            info(f"❌ DIFFERENT page count!")
        doc_san.close()
        doc_orig.close()
        
        # Cleanup
        for tmp_dir in {os.path.dirname(sanitized), os.path.dirname(pdf_sanitized), os.path.dirname(pdf_original)}:
            try: shutil.rmtree(tmp_dir, ignore_errors=True)
            except: pass
            
    except Exception as e:
        info(f"Export analysis FAILED: {e}")
        traceback.print_exc()

# ══════════════════════════════════════════════════════════════════════════
# STAGE 3: Check existing runtime metadata on disk
# ══════════════════════════════════════════════════════════════════════════
stage("3: Existing runtime metadata")

forms_dir = ROOT / "ExcelAPI" / "ExcelAPI" / "Forms"
meta_files = list(forms_dir.glob("*.runtime.json")) + list(forms_dir.glob("*.meta.json")) + list(forms_dir.glob("*.xlsx"))
info(f"Found {len(meta_files)} files in Forms/ directory")

# Check for any runtime JSON files that might correspond to these workbooks
for f in meta_files:
    info(f"  {f.name} ({f.stat().st_size} bytes)")

# ══════════════════════════════════════════════════════════════════════════
# SUMMARY: First divergence analysis
# ══════════════════════════════════════════════════════════════════════════

print("\n\n")
print("=" * 70)
print("  DIVERGENCE ANALYSIS SUMMARY")
print("=" * 70)

# Write the debug output to a file for later reference
output_path = OUTPUT
print(f"\nFull report -> {output_path}")
