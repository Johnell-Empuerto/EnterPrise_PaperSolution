"""
Final Validation Investigation — Prove Whether Excel or LibreOffice Is the Root Cause

Experiments 1-6 all in one script.
"""

import os, sys, tempfile, shutil, json, math
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import fitz  # PyMuPDF
import numpy as np
from PIL import Image
import cv2

from render_service.pdf_converter import xlsx_to_pdf
from render_service.xml_field_provider import get_cluster_fielddefs

DPI = 300
PT_TO_PX = DPI / 72.0

TEMPLATES = {
    547: {
        "xlsx": os.path.join(_PROJECT_ROOT, "[V3.1_Sample]アンケート用紙.xlsx"),
        "label": "547 (A4 - アンケート用紙)",
    },
    546: {
        "xlsx": os.path.join(_PROJECT_ROOT, "Investigation_546", "original.xlsx"),
        "label": "546 (Letter - Investigation)",
    },
}

OUTPUT_DIR = os.path.join(_PROJECT_ROOT, "ExcelAPI", "ExcelAPI", "Preview")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def generate_pdf_libreoffice(xlsx_path: str) -> str:
    """Generate PDF via LibreOffice, return the PDF path."""
    tmp = tempfile.mkdtemp(prefix="ple_lo_")
    pdf = xlsx_to_pdf(xlsx_path, output_dir=tmp)
    return pdf, tmp


def generate_pdf_excel_com(xlsx_path: str) -> str:
    """Generate PDF via Excel COM ExportAsFixedFormat, return the PDF path."""
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    tmp = tempfile.mkdtemp(prefix="ple_excel_")
    pdf_path = os.path.join(tmp, "output.pdf")
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  # xlTypePDF = 0
        wb.Close(False)
    finally:
        excel.Quit()
    return pdf_path, tmp


def rasterize_pdf(pdf_path: str, dpi: int = 300, page_index: int = 0) -> tuple:
    """Rasterize a PDF page to a numpy array (RGB) and get metadata."""
    doc = fitz.open(pdf_path)
    page = doc[page_index]
    
    # Collect PDF box information
    rect = page.rect
    mediabox = page.mediabox
    cropbox = page.cropbox
    
    # Check for BleedBox, TrimBox, ArtBox (may be None)
    bleedbox = page.bleedbox if hasattr(page, 'bleedbox') else None
    trimbox = page.trimbox if hasattr(page, 'trimbox') else None
    artbox = page.artbox if hasattr(page, 'artbox') else None
    rotation = page.rotation
    
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    
    # Convert to numpy array
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    img_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)  # OpenCV uses BGR
    
    info = {
        "pdf_path": pdf_path,
        "page_index": page_index,
        "dpi": dpi,
        "rect": {"width": rect.width, "height": rect.height},
        "mediabox": {"x0": mediabox.x0, "y0": mediabox.y0, "x1": mediabox.x1, "y1": mediabox.y1,
                     "width": mediabox.width, "height": mediabox.height},
        "cropbox": {"x0": cropbox.x0, "y0": cropbox.y0, "x1": cropbox.x1, "y1": cropbox.y1,
                    "width": cropbox.width, "height": cropbox.height},
        "bleedbox": {"x0": bleedbox.x0, "y0": bleedbox.y0, "x1": bleedbox.x1, "y1": bleedbox.y1,
                     "width": bleedbox.width, "height": bleedbox.height} if bleedbox else None,
        "trimbox": {"x0": trimbox.x0, "y0": trimbox.y0, "x1": trimbox.x1, "y1": trimbox.y1,
                    "width": trimbox.width, "height": trimbox.height} if trimbox else None,
        "artbox": {"x0": artbox.x0, "y0": artbox.y0, "x1": artbox.x1, "y1": artbox.y1,
                   "width": artbox.width, "height": artbox.height} if artbox else None,
        "rotation": rotation,
        "pixmap_width": pix.width,
        "pixmap_height": pix.height,
    }
    
    doc.close()
    return img_rgb, info


def detect_grid_lines(img_bgr: np.ndarray) -> tuple:
    """Detect horizontal and vertical grid lines using OpenCV.
    Returns (horiz_lines, vert_lines) as sorted lists of pixel coordinates."""
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    
    # Adaptive threshold
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, 51, 8
    )
    
    h, w = img_bgr.shape[:2]
    
    # Morphological close to connect broken grid lines
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (max(w // 40, 21), 1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(h // 40, 21)))
    lines_h = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel_h)
    lines_v = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel_v)
    
    # Hough lines
    hough_h = cv2.HoughLinesP(
        lines_h, rho=1, theta=np.pi / 180, threshold=100,
        minLineLength=w * 0.3, maxLineGap=20
    )
    hough_v = cv2.HoughLinesP(
        lines_v, rho=1, theta=np.pi / 180, threshold=100,
        minLineLength=h * 0.3, maxLineGap=20
    )
    
    def extract_and_cluster(lines):
        if lines is None:
            return []
        positions = []
        for line in lines:
            x1, y1, x2, y2 = line[0]
            if abs(y2 - y1) < abs(x2 - x1):  # horizontal
                pos = (y1 + y2) // 2
                positions.append(pos)
            else:  # vertical
                pos = (x1 + x2) // 2
                positions.append(pos)
        if not positions:
            return []
        positions = sorted(set(positions))
        clusters = []
        current = [positions[0]]
        for p in positions[1:]:
            if abs(p - current[-1]) <= 8:
                current.append(p)
            else:
                clusters.append(int(np.mean(current)))
                current = [p]
        clusters.append(int(np.mean(current)))
        return clusters
    
    horiz = extract_and_cluster(hough_h)
    vert = extract_and_cluster(hough_v)
    return sorted(horiz), sorted(vert)


def compute_xor_heatmap(img1: np.ndarray, img2: np.ndarray) -> np.ndarray:
    """Compute pixel difference between two images. Returns heatmap image."""
    gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    
    # Resize larger to smaller if dimensions differ
    h = min(gray1.shape[0], gray2.shape[0])
    w = min(gray1.shape[1], gray2.shape[1])
    gray1 = cv2.resize(gray1, (w, h))
    gray2 = cv2.resize(gray2, (w, h))
    
    diff = cv2.absdiff(gray1, gray2)
    # Normalize to 0-255 for visibility
    heatmap = cv2.normalize(diff, None, 0, 255, cv2.NORM_MINMAX)
    heatmap_color = cv2.applyColorMap(heatmap, cv2.COLORMAP_JET)
    
    stats = {
        "total_pixels": w * h,
        "changed_pixels": int(np.sum(diff > 10)),
        "mean_diff": float(np.mean(diff)),
        "max_diff": float(np.max(diff)),
        "pct_changed": float(np.sum(diff > 10) / (w * h) * 100),
    }
    return heatmap_color, stats, diff


def estimate_translation(img1: np.ndarray, img2: np.ndarray) -> dict:
    """Use phase correlation to estimate translation between two images."""
    gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    
    h = min(gray1.shape[0], gray2.shape[0])
    w = min(gray1.shape[1], gray2.shape[1])
    gray1 = cv2.resize(gray1, (w, h))
    gray2 = cv2.resize(gray2, (w, h))
    
    # Phase correlation
    fft1 = np.fft.fft2(gray1.astype(np.float32))
    fft2 = np.fft.fft2(gray2.astype(np.float32))
    cross_power = (fft1 * fft2.conj()) / (np.abs(fft1 * fft2) + 1e-10)
    phase_corr = np.fft.ifft2(cross_power).real
    
    # Find peak
    max_idx = np.unravel_index(np.argmax(phase_corr), phase_corr.shape)
    dy = max_idx[0] if max_idx[0] <= h // 2 else max_idx[0] - h
    dx = max_idx[1] if max_idx[1] <= w // 2 else max_idx[1] - w
    
    return {"dx": dx, "dy": dy, "confidence": float(phase_corr[max_idx])}


def match_grid_lines(lines1: list, lines2: list) -> dict:
    """Match corresponding grid lines between two images and compute statistics."""
    if not lines1 or not lines2:
        return {"matched_pairs": 0, "mean_delta": None, "max_delta": None, "std_delta": None}
    
    matched = []
    used = set()
    for l1 in lines1:
        best_dist = float('inf')
        best_l2 = None
        for j, l2 in enumerate(lines2):
            if j in used:
                continue
            dist = abs(l1 - l2)
            if dist < best_dist and dist < 20:  # 20px max match distance
                best_dist = dist
                best_l2 = j
        if best_l2 is not None:
            matched.append(best_dist)
            used.add(best_l2)
    
    if not matched:
        return {"matched_pairs": 0, "mean_delta": None, "max_delta": None, "std_delta": None}
    
    return {
        "matched_pairs": len(matched),
        "mean_delta": float(np.mean(matched)),
        "max_delta": float(np.max(matched)),
        "std_delta": float(np.std(matched)),
    }


# =============================================================================
# MAIN EXPERIMENT
# =============================================================================

def main():
    results = {}
    
    for template_id, info in TEMPLATES.items():
        xlsx_path = info["xlsx"]
        label = info["label"]
        
        print(f"\n{'='*80}")
        print(f"TEMPLATE {label}")
        print(f"{'='*80}")
        
        if not os.path.exists(xlsx_path):
            print(f"  SKIP: {xlsx_path} not found")
            continue
        
        print(f"\n  XLSX: {xlsx_path}")
        
        # ---------------------------------------------------------------
        # EXPERIMENT 4: Compare PDF geometry
        # ---------------------------------------------------------------
        print(f"\n  --- EXPERIMENT 4: PDF Geometry Comparison ---")
        
        # Generate PDF via LibreOffice
        print(f"  Generating LibreOffice PDF...")
        lo_pdf, lo_tmp = generate_pdf_libreoffice(xlsx_path)
        
        # Generate PDF via Excel COM
        print(f"  Generating Excel COM PDF...")
        try:
            excel_pdf, excel_tmp = generate_pdf_excel_com(xlsx_path)
            excel_available = True
        except Exception as e:
            print(f"  Excel COM FAILED: {e}")
            excel_available = False
        
        # Analyze PDF geometry
        for engine_name, pdf_path in [("LibreOffice", lo_pdf), ("Excel COM", excel_pdf)]:
            if not excel_available and engine_name == "Excel COM":
                continue
            doc = fitz.open(pdf_path)
            page = doc[0]
            
            print(f"\n  [{engine_name}] PDF Geometry:")
            print(f"    Page count: {len(doc)}")
            print(f"    Rect:        ({page.rect.x0:.2f}, {page.rect.y0:.2f}) -> ({page.rect.x1:.2f}, {page.rect.y1:.2f}) = {page.rect.width:.2f} x {page.rect.height:.2f}")
            print(f"    MediaBox:    ({page.mediabox.x0:.2f}, {page.mediabox.y0:.2f}) -> ({page.mediabox.x1:.2f}, {page.mediabox.y1:.2f}) = {page.mediabox.width:.2f} x {page.mediabox.height:.2f}")
            print(f"    CropBox:     ({page.cropbox.x0:.2f}, {page.cropbox.y0:.2f}) -> ({page.cropbox.x1:.2f}, {page.cropbox.y1:.2f}) = {page.cropbox.width:.2f} x {page.cropbox.height:.2f}")
            if hasattr(page, 'bleedbox') and page.bleedbox:
                print(f"    BleedBox:    ({page.bleedbox.x0:.2f}, {page.bleedbox.y0:.2f}) -> ({page.bleedbox.x1:.2f}, {page.bleedbox.y1:.2f})")
            if hasattr(page, 'trimbox') and page.trimbox:
                print(f"    TrimBox:     ({page.trimbox.x0:.2f}, {page.trimbox.y0:.2f}) -> ({page.trimbox.x1:.2f}, {page.trimbox.y1:.2f})")
            if hasattr(page, 'artbox') and page.artbox:
                print(f"    ArtBox:      ({page.artbox.x0:.2f}, {page.artbox.y0:.2f}) -> ({page.artbox.x1:.2f}, {page.artbox.y1:.2f})")
            print(f"    Rotation:    {page.rotation}")
            doc.close()
        
        # ---------------------------------------------------------------
        # EXPERIMENT 2: Same Rasterizer (PyMuPDF), Different PDF Engine
        # ---------------------------------------------------------------
        print(f"\n  --- EXPERIMENT 2: Same Rasterizer, Different PDF Engine ---")
        
        # Rasterize both PDFs with PyMuPDF at 300dpi
        lo_img, lo_info = rasterize_pdf(lo_pdf, dpi=DPI)
        print(f"\n  [LibreOffice PNG] Size: {lo_info['pixmap_width']} x {lo_info['pixmap_height']} px")
        print(f"    Expected (rect×{DPI}/72): {lo_info['rect']['width']*DPI/72:.1f} x {lo_info['rect']['height']*DPI/72:.1f} px")
        
        if excel_available:
            excel_img, excel_info = rasterize_pdf(excel_pdf, dpi=DPI)
            print(f"\n  [Excel COM PNG]    Size: {excel_info['pixmap_width']} x {excel_info['pixmap_height']} px")
            print(f"    Expected (rect×{DPI}/72): {excel_info['rect']['width']*DPI/72:.1f} x {excel_info['rect']['height']*DPI/72:.1f} px")
            
            # Compare dimensions
            w_diff = excel_info['pixmap_width'] - lo_info['pixmap_width']
            h_diff = excel_info['pixmap_height'] - lo_info['pixmap_height']
            print(f"\n  DIMENSION DIFFERENCE (Excel - LibreOffice): {w_diff} px W, {h_diff} px H")
            
            if w_diff == 0 and h_diff == 0:
                print("  *** PNG dimensions are IDENTICAL ***")
            else:
                print(f"  *** PNG dimensions DIFFER by {abs(w_diff)}x{abs(h_diff)} px ***")
            
            # ---------------------------------------------------------------
            # EXPERIMENT 3: Pixel Difference Heatmap
            # ---------------------------------------------------------------
            print(f"\n  --- EXPERIMENT 3: Pixel Difference Heatmap ---")
            
            # Resize to same dimensions for comparison
            h_min = min(lo_img.shape[0], excel_img.shape[0])
            w_min = min(lo_img.shape[1], excel_img.shape[1])
            lo_resized = cv2.resize(lo_img, (w_min, h_min))
            excel_resized = cv2.resize(excel_img, (w_min, h_min))
            
            heatmap, xor_stats, raw_diff = compute_xor_heatmap(lo_resized, excel_resized)
            print(f"    Total pixels: {xor_stats['total_pixels']}")
            print(f"    Changed pixels (>10): {xor_stats['changed_pixels']} ({xor_stats['pct_changed']:.2f}%)")
            print(f"    Mean pixel diff: {xor_stats['mean_diff']:.2f}")
            print(f"    Max pixel diff: {xor_stats['max_diff']}")
            
            # Save heatmap
            heatmap_path = os.path.join(OUTPUT_DIR, f"heatmap_{template_id}.png")
            cv2.imwrite(heatmap_path, heatmap)
            print(f"    Heatmap saved: {heatmap_path}")
            
            # Estimate translation
            trans = estimate_translation(lo_resized, excel_resized)
            print(f"\n  PHASE CORRELATION TRANSLATION ESTIMATE:")
            print(f"    dx = {trans['dx']} px, dy = {trans['dy']} px")
            print(f"    confidence: {trans['confidence']:.4f}")
            
            # ---------------------------------------------------------------
            # EXPERIMENT 5: Grid Line Alignment
            # ---------------------------------------------------------------
            print(f"\n  --- EXPERIMENT 5: Grid Line Alignment ---")
            
            lo_horiz, lo_vert = detect_grid_lines(lo_resized)
            excel_horiz, excel_vert = detect_grid_lines(excel_resized)
            
            print(f"\n  [LibreOffice] Grid lines: {len(lo_horiz)}H x {len(lo_vert)}V")
            print(f"  [Excel COM]    Grid lines: {len(excel_horiz)}H x {len(excel_vert)}V")
            
            if lo_horiz and excel_horiz:
                h_match = match_grid_lines(lo_horiz, excel_horiz)
                print(f"\n  Horizontal line matching:")
                print(f"    Matched pairs: {h_match['matched_pairs']}")
                if h_match['mean_delta'] is not None:
                    print(f"    Mean delta:    {h_match['mean_delta']:.2f} px")
                    print(f"    Max delta:     {h_match['max_delta']:.2f} px")
                    print(f"    Std delta:     {h_match['std_delta']:.2f} px")
            
            if lo_vert and excel_vert:
                v_match = match_grid_lines(lo_vert, excel_vert)
                print(f"\n  Vertical line matching:")
                print(f"    Matched pairs: {v_match['matched_pairs']}")
                if v_match['mean_delta'] is not None:
                    print(f"    Mean delta:    {v_match['mean_delta']:.2f} px")
                    print(f"    Max delta:     {v_match['max_delta']:.2f} px")
                    print(f"    Std delta:     {v_match['std_delta']:.2f} px")
            
            # Save comparison overlay
            overlay = np.zeros((h_min, w_min * 2, 3), dtype=np.uint8)
            overlay[:, :w_min] = lo_resized
            overlay[:, w_min:] = excel_resized
            compare_path = os.path.join(OUTPUT_DIR, f"compare_{template_id}.png")
            cv2.imwrite(compare_path, overlay)
            print(f"\n    Side-by-side comparison saved: {compare_path}")
            
            # ---------------------------------------------------------------
            # EXPERIMENT 6: Ratio Validation
            # ---------------------------------------------------------------
            print(f"\n  --- EXPERIMENT 6: Ratio Validation ---")
            
            try:
                fields, raw_clusters = get_cluster_fielddefs(template_id)
                print(f"  Fields from DB: {len(fields)}")
                
                for i, (field, cluster) in enumerate(zip(fields, raw_clusters)):
                    if i >= 5:  # Check first 5 fields
                        break
                    
                    rl = cluster.get("left_position")
                    rt = cluster.get("top_position")
                    rr = cluster.get("right_position")
                    rb = cluster.get("bottom_position")
                    
                    if rl is None:
                        continue
                    
                    rl, rt, rr, rb = float(rl), float(rt), float(rr), float(rb)
                    
                    print(f"\n  Field {field.addr}:")
                    print(f"    Ratios: L={rl:.6f} T={rt:.6f} R={rr:.6f} B={rb:.6f}")
                    
                    # Compute pixel positions for both engines
                    lo_left   = rl * lo_info['pixmap_width']
                    lo_top    = rt * lo_info['pixmap_height']
                    lo_right  = rr * lo_info['pixmap_width']
                    lo_bottom = rb * lo_info['pixmap_height']
                    
                    excel_left   = rl * excel_info['pixmap_width']
                    excel_top    = rt * excel_info['pixmap_height']
                    excel_right  = rr * excel_info['pixmap_width']
                    excel_bottom = rb * excel_info['pixmap_height']
                    
                    print(f"    LibreOffice: L={lo_left:.2f} T={lo_top:.2f} R={lo_right:.2f} B={lo_bottom:.2f}  ({lo_right-lo_left:.1f}×{lo_bottom-lo_top:.1f})")
                    print(f"    Excel COM:   L={excel_left:.2f} T={excel_top:.2f} R={excel_right:.2f} B={excel_bottom:.2f}  ({excel_right-excel_left:.1f}×{excel_bottom-excel_top:.1f})")
                    
                    delta_l = excel_left - lo_left
                    delta_t = excel_top - lo_top
                    delta_r = excel_right - lo_right
                    delta_b = excel_bottom - lo_bottom
                    
                    print(f"    Delta (Excel - LO): L={delta_l:.2f} T={delta_t:.2f} R={delta_r:.2f} B={delta_b:.2f} px")
                    
                    # In points
                    print(f"    Delta in points:     L={delta_l/PT_TO_PX:.3f} T={delta_t/PT_TO_PX:.3f} R={delta_r/PT_TO_PX:.3f} B={delta_b/PT_TO_PX:.3f}")
            except Exception as e:
                print(f"  Ratio validation failed: {e}")
        
        # Save the LibreOffice PNG as our reference
        lo_png = os.path.join(OUTPUT_DIR, f"lo_{template_id}_page1.png")
        cv2.imwrite(lo_png, lo_img)
        
        if excel_available:
            excel_png = os.path.join(OUTPUT_DIR, f"excel_{template_id}_page1.png")
            cv2.imwrite(excel_png, excel_img)
        
        # Cleanup temp dirs
        shutil.rmtree(lo_tmp, ignore_errors=True)
        if excel_available:
            shutil.rmtree(excel_tmp, ignore_errors=True)
        
        # Store results
        results[template_id] = {
            "excel_available": excel_available,
            "lo_dims": (lo_info['pixmap_width'], lo_info['pixmap_height']),
            "excel_dims": (excel_info['pixmap_width'], excel_info['pixmap_height']) if excel_available else None,
        }
    
    # =========================================================================
    # FINAL ANSWER
    # =========================================================================
    print(f"\n\n{'='*80}")
    print("FINAL ANSWER")
    print(f"{'='*80}")
    print()
    
    for template_id in TEMPLATES:
        r = results.get(template_id)
        if not r or not r.get("excel_available"):
            continue
        
        lo_w, lo_h = r["lo_dims"]
        ex_w, ex_h = r["excel_dims"]
        
        if lo_w == ex_w and lo_h == ex_h:
            print(f"Template {template_id}: PNG dimensions MATCH (both {ex_w}x{ex_h})")
        else:
            print(f"Template {template_id}: PNG dimensions DIFFER - LO: {lo_w}x{lo_h}, Excel: {ex_w}x{ex_h}")
            ratio_w = ex_w / lo_w if lo_w > 0 else 0
            ratio_h = ex_h / lo_h if lo_h > 0 else 0
            print(f"  Scale factor: {ratio_w:.6f} x {ratio_h:.6f}")
    
    print()
    print("To answer the final question definitively, check:") 
    print("  1. If PNG dimensions are identical → translation estimate tells us if content shifted")
    print("  2. If grid line deltas are 0 on average → same positioning")
    print("  3. If ratio×PNG_dim gives same pixel values for both engines → zero-calibration works")


if __name__ == "__main__":
    main()
