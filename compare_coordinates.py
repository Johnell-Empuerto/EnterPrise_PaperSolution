"""
Coordinate Validation Script — Phase 14
Compares our runtime coordinates with legacy ConMas database for template 546.
Generates overlay comparison images in Comparison_Output/ folder.
"""

import json
import os
import math

# ──────────── Legacy Database Coordinates (irepodb.def_cluster for top_id=546) ────────────
LEGACY_FIELDS = [
    {"cell": "$A$1:$B$2",  "left_r": 0.3364706, "top_r": 0.3845454,  "right_r": 0.4982353, "bottom_r": 0.4218182},
    {"cell": "$C$1:$D$2",  "left_r": 0.5,        "top_r": 0.3845454,  "right_r": 0.6635294, "bottom_r": 0.4218182},
    {"cell": "$A$3:$D$4",  "left_r": 0.3364706, "top_r": 0.4231818,  "right_r": 0.6635294, "bottom_r": 0.4604546},
    {"cell": "$A$6:$D$7",  "left_r": 0.3364706, "top_r": 0.4809091,  "right_r": 0.6635294, "bottom_r": 0.5181818},
    {"cell": "$A$9:$D$10", "left_r": 0.3364706, "top_r": 0.5386364,  "right_r": 0.6635294, "bottom_r": 0.5759091},
    {"cell": "$A$12",      "left_r": 0.3364706, "top_r": 0.5963637,  "right_r": 0.4164706, "bottom_r": 0.615},
]

# ──────────── Page Setup (from Investigation_546/page_setup_dump.json) ────────────
PAGE_WIDTH_PT = 612.0
PAGE_HEIGHT_PT = 792.0
LEFT_MARGIN_PT = 51.0236220472441
RIGHT_MARGIN_PT = 51.0236220472441
TOP_MARGIN_PT = 53.85826771653544
BOTTOM_MARGIN_PT = 53.85826771653544
CENTER_H = True
CENTER_V = True

PRINTABLE_W = PAGE_WIDTH_PT - LEFT_MARGIN_PT - RIGHT_MARGIN_PT  # 509.952pt
PRINTABLE_H = PAGE_HEIGHT_PT - TOP_MARGIN_PT - BOTTOM_MARGIN_PT  # 684.284pt

# ──────────── COM Range Data (from Investigation_546/com_pipeline_dump.json) ────────────
PRINT_AREA_ADDR = "$A$1:$D$12"
RANGE_WIDTH_PT = 200.3757  # PAW
RANGE_HEIGHT_PT = 180.0     # PAH

# Cell positions (worksheet-relative points from COM)
CELL_POSITIONS = {
    "$A$1:$B$2":  {"left_pt": 0,        "top_pt": 0,       "width_pt": 100.1878, "height_pt": 30.0},
    "$C$1:$D$2":  {"left_pt": 100.1878,  "top_pt": 0,       "width_pt": 100.1878, "height_pt": 30.0},
    "$A$3:$D$4":  {"left_pt": 0,        "top_pt": 30.0,     "width_pt": 200.3757, "height_pt": 30.0},
    "$A$6:$D$7":  {"left_pt": 0,        "top_pt": 75.0,     "width_pt": 200.3757, "height_pt": 30.0},
    "$A$9:$D$10": {"left_pt": 0,        "top_pt": 120.0,    "width_pt": 200.3757, "height_pt": 30.0},
    "$A$12":      {"left_pt": 0,        "top_pt": 165.0,    "width_pt": 50.0939,  "height_pt": 15.0},
}

# ──────────── ConMas Effective Dimensions (from investigation) ────────────
EFF_W_PT = 200.16
EFF_H_PT = 182.8801

# ──────────── PNG dimensions (Letter at 300 DPI) ────────────
PNG_W = 2550
PNG_H = 3299
SCALE_X = PNG_W / PAGE_WIDTH_PT  # 4.166667
SCALE_Y = PNG_H / PAGE_HEIGHT_PT  # 4.165404 (slight rounding difference)

DPI = 300.0
POINTS_TO_PIXELS = DPI / 72.0  # 4.166667

# ======= Coordinate Computation Methods =======

def compute_origin_standard(margin_pt, printable_pt, content_pt, center):
    """Standard formula: origin = margin + (printable - content) / 2 if centering"""
    origin = margin_pt
    if center and content_pt < printable_pt:
        origin += (printable_pt - content_pt) / 2.0
    return origin

def compute_origin_conmas(margin_pt, printable_pt, effective_pt):
    """ConMas formula: origin = margin + (printable - effective) / 2"""
    return margin_pt + (printable_pt - effective_pt) / 2.0

def compute_pixel_standard(cell_pt, print_area_origin_pt, origin_pt, scale):
    """Standard: pixel = (origin + (cell - printArea)) * scale"""
    return round((origin_pt + (cell_pt - print_area_origin_pt)) * scale, 1)

def compute_pixel_conmas(cell_pt, print_area_origin_pt, origin_pt, eff_dim, range_dim, scale):
    """ConMas: pixel = (origin + (cell - printArea) * (eff_dim / range_dim)) * scale"""
    if range_dim <= 0 or eff_dim <= 0:
        factor = 1.0
    else:
        factor = eff_dim / range_dim
    return round((origin_pt + (cell_pt - print_area_origin_pt) * factor) * scale, 1)

# ======= Compute All Coordinates =======

print("=" * 90)
print(f"{'Field':<15} {'Metric':<10} {'Legacy(px)':<12} {'Std(px)':<10} {'ConMas(px)':<12} {'DStd':<10} {'DCMas':<10}")
print("-" * 90)

results = []
for field in LEGACY_FIELDS:
    cell = field["cell"]
    pos = CELL_POSITIONS[cell]
    
    # Legacy pixel coordinates (from DB ratios)
    leg_left = round(field["left_r"] * PNG_W, 1)
    leg_top = round(field["top_r"] * PNG_H, 1)
    leg_right = round(field["right_r"] * PNG_W, 1)
    leg_bottom = round(field["bottom_r"] * PNG_H, 1)
    leg_w = round(leg_right - leg_left, 1)
    leg_h = round(leg_bottom - leg_top, 1)
    
    # Standard formula (current CoordinateTransformer)
    origin_x_std = compute_origin_standard(LEFT_MARGIN_PT, PRINTABLE_W, RANGE_WIDTH_PT, CENTER_H)
    origin_y_std = compute_origin_standard(TOP_MARGIN_PT, PRINTABLE_H, RANGE_HEIGHT_PT, CENTER_V)
    
    std_left = compute_pixel_standard(pos["left_pt"], 0, origin_x_std * POINTS_TO_PIXELS, 1.0)
    std_top = compute_pixel_standard(pos["top_pt"], 0, origin_y_std * POINTS_TO_PIXELS, 1.0)
    std_w = round(pos["width_pt"] * POINTS_TO_PIXELS, 1)
    std_h = round(pos["height_pt"] * POINTS_TO_PIXELS, 1)
    
    # ConMas formula (our fix)
    origin_x_cm = compute_origin_conmas(LEFT_MARGIN_PT, PRINTABLE_W, EFF_W_PT)
    origin_y_cm = compute_origin_conmas(TOP_MARGIN_PT, PRINTABLE_H, EFF_H_PT)
    
    cm_left = compute_pixel_conmas(pos["left_pt"], 0, origin_x_cm, EFF_W_PT, RANGE_WIDTH_PT, POINTS_TO_PIXELS)
    cm_top = compute_pixel_conmas(pos["top_pt"], 0, origin_y_cm, EFF_H_PT, RANGE_HEIGHT_PT, POINTS_TO_PIXELS)
    cm_w = round(pos["width_pt"] * (EFF_W_PT / RANGE_WIDTH_PT) * POINTS_TO_PIXELS, 1)
    cm_h = round(pos["height_pt"] * (EFF_H_PT / RANGE_HEIGHT_PT) * POINTS_TO_PIXELS, 1)
    
    # Calculate deltas
    d_left_std = round(leg_left - std_left, 1)
    d_top_std = round(leg_top - std_top, 1)
    d_left_cm = round(leg_left - cm_left, 1)
    d_top_cm = round(leg_top - cm_top, 1)
    
    print(f"{cell:<15} {'Left':<10} {leg_left:<12} {std_left:<10} {cm_left:<12} {d_left_std:<+10} {d_left_cm:<+10}")
    print(f"{'':<15} {'Top':<10} {leg_top:<12} {std_top:<10} {cm_top:<12} {d_top_std:<+10} {d_top_cm:<+10}")
    print(f"{'':<15} {'Width':<10} {leg_w:<12} {std_w:<10} {cm_w:<12}")
    print(f"{'':<15} {'Height':<10} {leg_h:<12} {std_h:<10} {cm_h:<12}")
    print("-" * 90)
    
    results.append({
        "cell": cell,
        "legacy": {"left": leg_left, "top": leg_top, "width": leg_w, "height": leg_h},
        "std": {"left": std_left, "top": std_top, "width": std_w, "height": std_h},
        "conmas": {"left": cm_left, "top": cm_top, "width": cm_w, "height": cm_h},
        "delta_std": {"left": d_left_std, "top": d_top_std},
        "delta_cm": {"left": d_left_cm, "top": d_top_cm}
    })

# Print summary
print(f"\n{'='*90}")
print("SUMMARY")
print(f"{'='*90}")
avg_d_left_std = sum(r["delta_std"]["left"] for r in results) / len(results)
avg_d_top_std = sum(r["delta_std"]["top"] for r in results) / len(results)
avg_d_left_cm = sum(r["delta_cm"]["left"] for r in results) / len(results)
avg_d_top_cm = sum(r["delta_cm"]["top"] for r in results) / len(results)

print(f"Standard formula:  Avg dLeft={avg_d_left_std:+.1f}px, Avg dTop={avg_d_top_std:+.1f}px")
print(f"ConMas formula:    Avg dLeft={avg_d_left_cm:+.1f}px, Avg dTop={avg_d_top_cm:+.1f}px")
print(f"Target:            0.0px offset = perfect alignment with legacy database")
print()

# Now generate overlay images
print("Generating comparison overlays...")

try:
    from PIL import Image, ImageDraw
    
    # Check if a preview PNG exists
    preview_found = False
    for f in os.listdir("."):
        if f.endswith(".png") and os.path.getsize(f) > 10000:
            preview_path = f
            preview_found = True
            break
    
    if not preview_found:
        # Try to find one in Preview directory
        for p in ["ExcelAPI/ExcelAPI/Preview", "Preview"]:
            if os.path.exists(p):
                for f in os.listdir(p):
                    if f.startswith("page_") and f.endswith(".png"):
                        preview_path = os.path.join(p, f)
                        preview_found = True
                        break
            if preview_found:
                break
    
    if not preview_found:
        # Use a plain white image
        print("No preview PNG found, creating blank comparison")
        img = Image.new("RGB", (PNG_W, PNG_H), (255, 255, 255))
    else:
        img = Image.open(preview_path)
        print(f"Using preview image: {preview_path}")
    
    # Scale to displayable size (1/4 for speed)
    scale_display = 0.25
    display_w = int(PNG_W * scale_display)
    display_h = int(PNG_H * scale_display)
    
    # Create comparison output directory
    os.makedirs("Comparison_Output", exist_ok=True)
    
    # Full overlay (all fields on one image)
    overlay = img.copy()
    draw = ImageDraw.Draw(overlay)
    
    for r in results:
        cell = r["cell"]
        leg = r["legacy"]
        cm = r["conmas"]
        
        def scale_rect(l, t, w, h):
            return (l, t, l + w, t + h)
        
        # Legacy = Red (thick)
        draw.rectangle(scale_rect(leg["left"], leg["top"], leg["width"], leg["height"]),
                       outline=(255, 0, 0), width=3)
        # ConMas = Green (medium)
        draw.rectangle(scale_rect(cm["left"], cm["top"], cm["width"], cm["height"]),
                       outline=(0, 200, 0), width=2)
        
        # Label
        label = f"{cell} CM:{cm['left']:.0f},{cm['top']:.0f}"
        draw.text((leg["left"] + 5, leg["top"] - 18), label, fill=(0, 0, 0))
    
    overlay.save("Comparison_Output/legacy_vs_conmas.png")
    print("  Saved: Comparison_Output/legacy_vs_conmas.png")
    
    # Individual field comparisons
    for r in results:
        cell_safe = r["cell"].replace("$", "").replace(":", "_")
        leg = r["legacy"]
        cm = r["conmas"]
        
        # Crop region with padding
        min_x = max(0, int(min(leg["left"], cm["left"])) - 50)
        min_y = max(0, int(min(leg["top"], cm["top"])) - 50)
        max_x = min(PNG_W, int(max(leg["left"] + leg["width"], cm["left"] + cm["width"])) + 50)
        max_y = min(PNG_H, int(max(leg["top"] + leg["height"], cm["top"] + cm["height"])) + 50)
        
        crop = img.crop((min_x, min_y, max_x, max_y))
        draw_crop = ImageDraw.Draw(crop)
        
        # Legacy = Red
        draw_crop.rectangle((leg["left"] - min_x, leg["top"] - min_y,
                             leg["left"] - min_x + leg["width"], leg["top"] - min_y + leg["height"]),
                            outline=(255, 0, 0), width=2)
        # ConMas = Green
        draw_crop.rectangle((cm["left"] - min_x, cm["top"] - min_y,
                             cm["left"] - min_x + cm["width"], cm["top"] - min_y + cm["height"]),
                            outline=(0, 200, 0), width=2)
        
        # Legend
        draw_crop.text((5, 5), f"Red=Legacy Green=Ours", fill=(0, 0, 0))
        draw_crop.text((5, 20), f"dL={r['delta_cm']['left']:+.1f} dT={r['delta_cm']['top']:+.1f}", fill=(0, 0, 0))
        
        crop_path = f"Comparison_Output/compare_{cell_safe}.png"
        crop.save(crop_path)
        print(f"  Saved: {crop_path}")
    
    print("\nOverlay generation complete!")
    
except ImportError:
    print("PIL not installed — overlay images cannot be generated.")
    print("Install with: pip install Pillow")
except Exception as e:
    print(f"Error generating overlays: {e}")
    import traceback
    traceback.print_exc()
