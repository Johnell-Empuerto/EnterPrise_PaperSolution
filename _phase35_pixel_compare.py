"""
Phase X.35 - Tests 3, 4, 5, 6: Pixel comparison, overlay, consistency, header/footer
"""
import os, sys, hashlib, json
try:
    import fitz  # PyMuPDF
except ImportError:
    import pypdfium2 as fitz
    fitz = None

from PIL import Image
import numpy as np

PDF_DIR = "."
PDF_NAMES = ["Original", "ConMas", "Generated1", "Generated2"]
PDF_FILES = {n: os.path.join(PDF_DIR, f"_x35_pdf_{n}.pdf") for n in PDF_NAMES}

PNG_DIR = "_x35_pngs"
os.makedirs(PNG_DIR, exist_ok=True)

DPI = 300

def pdf_to_pngs(pdf_path, prefix):
    """Convert PDF to list of PNG images (one per page)."""
    doc = fitz.open(pdf_path)
    images = []
    for i, page in enumerate(doc):
        mat = fitz.Matrix(DPI/72, DPI/72)  # scale to DPI
        pix = page.get_pixmap(matrix=mat)
        img_path = os.path.join(PNG_DIR, f"{prefix}_page{i+1}.png")
        pix.save(img_path)
        # Also load as PIL Image for comparison
        pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append({
            "path": img_path,
            "pil": pil_img,
            "width": pix.width,
            "height": pix.height,
            "page": i+1
        })
        print(f"  Page {i+1}: {pix.width}x{pix.height}px -> {img_path}")
    doc.close()
    return images

def compare_images(img1, img2, label1, label2):
    """Compare two PIL images, return difference metrics."""
    w = min(img1.width, img2.width)
    h = min(img1.height, img2.height)
    
    arr1 = np.array(img1.resize((w, h)))
    arr2 = np.array(img2.resize((w, h)))
    
    diff = np.abs(arr1.astype(np.int16) - arr2.astype(np.int16))
    max_diff = int(diff.max())
    avg_diff = float(diff.mean())
    total_pixels = w * h
    diff_pixels = int(np.sum(diff > 1))  # pixels with >1 per-channel difference
    diff_percent = (diff_pixels / (total_pixels * 3)) * 100  # per channel
    
    # Create difference image
    diff_viz = np.abs(arr1.astype(np.int16) - arr2.astype(np.int16)).astype(np.uint8)
    diff_viz = np.clip(diff_viz * 10, 0, 255).astype(np.uint8)  # amplify for visibility
    diff_img = Image.fromarray(diff_viz)
    
    return {
        "max_deviation": max_diff,
        "avg_deviation": round(avg_diff, 4),
        "diff_pixels": diff_pixels,
        "diff_percent": round(diff_percent, 4),
        "image_width": w,
        "image_height": h,
        "diff_image": diff_img
    }

def overlay_analysis(img_base, img_overlay, label_base, label_overlay):
    """Analyze alignment of two images using edge detection."""
    arr_base = np.array(img_base.convert("L"))
    arr_over = np.array(img_overlay.convert("L"))
    
    h = min(arr_base.shape[0], arr_over.shape[0])
    w = min(arr_base.shape[1], arr_over.shape[1])
    
    arr_base = arr_base[:h, :w]
    arr_over = arr_over[:h, :w]
    
    # Simple edge detection (Sobel-like)
    edges_base = np.abs(np.diff(arr_base.astype(np.int16), axis=1))
    edges_over = np.abs(np.diff(arr_over.astype(np.int16), axis=1))
    
    # Find first non-empty row/col (content boundaries)
    col_sums_base = np.sum(arr_base, axis=0)
    col_sums_over = np.sum(arr_over, axis=0)
    row_sums_base = np.sum(arr_base, axis=1)
    row_sums_over = np.sum(arr_over, axis=1)
    
    threshold = np.mean(arr_base) * 0.1  # content threshold
    
    def find_first(pixels, thresh):
        return int(np.argmax(pixels > thresh * arr_base.shape[0]))
    
    def find_last(pixels, thresh):
        return int(arr_base.shape[1] - np.argmax(pixels[::-1] > thresh * arr_base.shape[0]))
    
    # Content area boundaries
    left_b = find_first(col_sums_base, threshold)
    left_o = find_first(col_sums_over, threshold)
    right_b = find_last(col_sums_base, threshold)
    right_o = find_last(col_sums_over, threshold)
    top_b = find_first(row_sums_base, threshold)
    top_o = find_first(row_sums_over, threshold)
    bottom_b = find_last(row_sums_base, threshold)
    bottom_o = find_last(row_sums_over, threshold)
    
    # Displacement in pixels and mm
    px_mm = 25.4 / DPI
    
    return {
        "left_offset_px": left_o - left_b,
        "left_offset_mm": round((left_o - left_b) * px_mm, 4),
        "right_offset_px": right_o - right_b,
        "right_offset_mm": round((right_o - right_b) * px_mm, 4),
        "top_offset_px": top_o - top_b,
        "top_offset_mm": round((top_o - top_b) * px_mm, 4),
        "bottom_offset_px": bottom_o - bottom_b,
        "bottom_offset_mm": round((bottom_o - bottom_b) * px_mm, 4),
        "content_width_base": right_b - left_b,
        "content_width_overlay": right_o - left_o,
        "content_height_base": bottom_b - top_b,
        "content_height_overlay": bottom_o - top_o,
    }

# ===== MAIN =====
print("=" * 80)
print("PHASE X.35 - Tests 3, 4, 5, 6: PDF -> PNG -> Pixel Comparison")
print("=" * 80)

# Verify all PDFs exist
print("\n--- Checking PDFs ---")
for name, path in PDF_FILES.items():
    if os.path.exists(path):
        size = os.path.getsize(path)
        md5 = hashlib.md5(open(path, "rb").read()).hexdigest()
        print(f"  {name}: {size:,} bytes, MD5={md5[:12]}")
    else:
        print(f"  {name}: MISSING")

# Test 3: Convert all to PNGs
print("\n" + "=" * 80)
print("TEST 3: PDF -> PNG Conversion (300 DPI)")
print("=" * 80)
all_pages = {}
for name in PDF_NAMES:
    path = PDF_FILES[name]
    if not os.path.exists(path):
        print(f"  {name}: PDF not found, skipping")
        continue
    print(f"\n{name}:")
    pages = pdf_to_pngs(path, name)
    all_pages[name] = pages

# Test 3: Pixel comparison
print("\n" + "=" * 80)
print("TEST 3: Pixel Comparison (ConMas vs others)")
print("=" * 80)

if "ConMas" in all_pages and len(all_pages["ConMas"]) > 0:
    conmas_page = all_pages["ConMas"][0]["pil"]
    for name in PDF_NAMES:
        if name == "ConMas" or name not in all_pages or len(all_pages[name]) == 0:
            continue
        print(f"\n  {name} vs ConMas (Page 1):")
        result = compare_images(all_pages[name][0]["pil"], conmas_page, name, "ConMas")
        print(f"    Image size: {result['image_width']}x{result['image_height']}px")
        print(f"    Max pixel deviation: {result['max_deviation']}")
        print(f"    Avg pixel deviation: {result['avg_deviation']:.4f}")
        print(f"    Different pixels: {result['diff_pixels']}")
        print(f"    Difference %: {result['diff_percent']:.4f}%")
        
        # Save diff image
        diff_path = os.path.join(PNG_DIR, f"diff_{name}_vs_ConMas.png")
        result['diff_image'].save(diff_path)
        print(f"    Difference image saved: {diff_path}")
        
        # If identical, note it
        if result['max_deviation'] == 0:
            print(f"    >>> IDENTICAL: No pixel differences")

# Test 4: Overlay analysis
print("\n" + "=" * 80)
print("TEST 4: Overlay Analysis (Generated1 over ConMas)")
print("=" * 80)

if "ConMas" in all_pages and "Generated1" in all_pages:
    if len(all_pages["ConMas"]) > 0 and len(all_pages["Generated1"]) > 0:
        overlay = overlay_analysis(
            conmas_page, all_pages["Generated1"][0]["pil"],
            "ConMas", "Generated1"
        )
        print(f"  Content Boundary Offsets (Generated1 relative to ConMas):")
        print(f"    Left:   {overlay['left_offset_px']:+d}px ({overlay['left_offset_mm']:+.4f}mm)")
        print(f"    Right:  {overlay['right_offset_px']:+d}px ({overlay['right_offset_mm']:+.4f}mm)")
        print(f"    Top:    {overlay['top_offset_px']:+d}px ({overlay['top_offset_mm']:+.4f}mm)")
        print(f"    Bottom: {overlay['bottom_offset_px']:+d}px ({overlay['bottom_offset_mm']:+.4f}mm)")
        print(f"  Content Dimensions:")
        print(f"    ConMas:     {overlay['content_width_base']}x{overlay['content_height_base']}px")
        print(f"    Generated1: {overlay['content_width_overlay']}x{overlay['content_height_overlay']}px")

# Test 5: Print engine consistency
print("\n" + "=" * 80)
print("TEST 5: Print Engine Consistency (hash comparison)")
print("=" * 80)

print("  (Consistency will be verified in Test 5b - re-export)")
print("  First-pass PDFs hashes for reference:")
for name in PDF_NAMES:
    path = PDF_FILES[name]
    if os.path.exists(path):
        md5 = hashlib.md5(open(path, "rb").read()).hexdigest()
        sha = hashlib.sha256(open(path, "rb").read()).hexdigest()
        size = os.path.getsize(path)
        print(f"    {name}: {size:,} bytes, MD5={md5[:16]}, SHA256={sha[:16]}...")

# Test 6: Header/Footer investigation
print("\n" + "=" * 80)
print("TEST 6: Header/Footer Margin Rendering Investigation")
print("=" * 80)

# Check if page content reaches top or bottom of the page
if "ConMas" in all_pages and "Generated1" in all_pages:
    if len(all_pages["ConMas"]) > 0 and len(all_pages["Generated1"]) > 0:
        conmas_arr = np.array(all_pages["ConMas"][0]["pil"].convert("L"))
        gen1_arr = np.array(all_pages["Generated1"][0]["pil"].convert("L"))
        
        h = min(conmas_arr.shape[0], gen1_arr.shape[0])
        w = min(conmas_arr.shape[1], gen1_arr.shape[1])
        
        # Check if top/bottom 50px have any content
        top_strip_c = conmas_arr[:50, :w]
        top_strip_g = gen1_arr[:50, :w]
        bottom_strip_c = conmas_arr[h-50:, :w]
        bottom_strip_g = gen1_arr[h-50:, :w]
        
        # A strip with mostly white (>95% white) = no content in header/footer zone
        white_threshold = 240
        top_white_c = np.mean(top_strip_c > white_threshold) * 100
        top_white_g = np.mean(top_strip_g > white_threshold) * 100
        bottom_white_c = np.mean(bottom_strip_c > white_threshold) * 100
        bottom_white_g = np.mean(bottom_strip_g > white_threshold) * 100
        
        print(f"  Header Zone (top 50px) whiteness:")
        print(f"    ConMas:     {top_white_c:.1f}% white")
        print(f"    Generated1: {top_white_g:.1f}% white")
        print(f"  Footer Zone (bottom 50px) whiteness:")
        print(f"    ConMas:     {bottom_white_c:.1f}% white")
        print(f"    Generated1: {bottom_white_g:.1f}% white")
        
        # Also check if the first non-white pixel is at same position
        def first_content_row(arr):
            for r in range(arr.shape[0]):
                if np.mean(arr[r] < 240) > 0.05:
                    return r
            return -1
        
        def last_content_row(arr):
            for r in range(arr.shape[0]-1, -1, -1):
                if np.mean(arr[r] < 240) > 0.05:
                    return r
            return -1
        
        fc_c = first_content_row(conmas_arr)
        fc_g = first_content_row(gen1_arr)
        lc_c = last_content_row(conmas_arr)
        lc_g = last_content_row(gen1_arr)
        
        print(f"  First content row:")
        print(f"    ConMas:     row {fc_c} ({fc_c/DPI*25.4:.2f}mm from top)")
        print(f"    Generated1: row {fc_g} ({fc_g/DPI*25.4:.2f}mm from top)")
        print(f"    Difference: {fc_g - fc_c}px ({(fc_g-fc_c)/DPI*25.4:.4f}mm)")
        print(f"  Last content row:")
        print(f"    ConMas:     row {lc_c} ({lc_c/DPI*25.4:.2f}mm from top)")
        print(f"    Generated1: row {lc_g} ({lc_g/DPI*25.4:.2f}mm from top)")
        print(f"    Difference: {lc_g - lc_c}px ({(lc_g-lc_c)/DPI*25.4:.4f}mm)")

print("\n" + "=" * 80)
print("Phase X.35 Tests 3-6 COMPLETE")
print("=" * 80)
