"""
Enhanced validation overlay: detects actual form element boundaries from the
background PNG using OpenCV and compares them against the ratio-based field
coordinates. Draws:
  - Green  + blue outline = CV-detected cell boundaries (SOURCE OF TRUTH)
  - Yellow + red outline  = ratio-based overlay coordinates
  - Offset annotation (px) if green and yellow differ by > 2 px

Usage:
    python render_service/validate_alignment.py 547
"""

import os
import sys
from pathlib import Path

# Ensure project root is on the path
_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

try:
    import cv2
    import numpy as np
    HAS_CV = True
except ImportError:
    HAS_CV = False

from PIL import Image, ImageDraw, ImageFont


# ── colour constants ──────────────────────────────────────────────────────────
CV_FILL   = (0, 255, 0, 48)    # green  translucent
CV_OUTLINE = (0, 100, 255, 220)  # orange-blue outline
RATIO_FILL   = (255, 255, 0, 48)   # yellow translucent
RATIO_OUTLINE = (255, 0, 0, 220)   # red outline
OFFSET_COLOUR = (255, 0, 255, 220) # magenta for offset arrows


def detect_table_lines(img_bgr: np.ndarray, debug_dir: str = None) -> tuple[
        list[float], list[float]]:
    """
    Detect horizontal and vertical table grid lines on a form background
    using Hough Line Transform + line clustering.

    Returns (horiz_lines, vert_lines) where each is a list of pixel
    coordinates sorted ascending.
    """
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    # Sharpening + adaptive threshold works better than Canny for printed forms
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, 51, 8
    )

    if debug_dir:
        cv2.imwrite(os.path.join(debug_dir, "01_thresh.png"), thresh)

    # Morphological close to connect broken grid lines
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (max(img_bgr.shape[1] // 40, 21), 1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(img_bgr.shape[0] // 40, 21)))
    lines_h = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel_h)
    lines_v = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel_v)

    if debug_dir:
        cv2.imwrite(os.path.join(debug_dir, "02_lines_h.png"), lines_h)
        cv2.imwrite(os.path.join(debug_dir, "03_lines_v.png"), lines_v)

    # Hough lines for horizontal
    hough_h = cv2.HoughLinesP(
        lines_h, rho=1, theta=np.pi / 180, threshold=100,
        minLineLength=img_bgr.shape[1] * 0.3, maxLineGap=20
    )
    hough_v = cv2.HoughLinesP(
        lines_v, rho=1, theta=np.pi / 180, threshold=100,
        minLineLength=img_bgr.shape[0] * 0.3, maxLineGap=20
    )

    def extract_and_cluster(lines):
        """Extract line positions and cluster nearby ones."""
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
        # Cluster: group positions within 8px of each other
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

    # Filter short lines (sometimes noise sneaks in)
    if debug_dir:
        debug_lines = img_bgr.copy()
        for y in horiz:
            cv2.line(debug_lines, (0, y), (img_bgr.shape[1], y), (0, 255, 0), 2)
        for x in vert:
            cv2.line(debug_lines, (x, 0), (x, img_bgr.shape[0]), (255, 0, 0), 2)
        cv2.imwrite(os.path.join(debug_dir, "04_detected_lines.png"), debug_lines)

    return sorted(horiz), sorted(vert)


def find_cell_for_ratio(ratio_left: float, ratio_top: float,
                        ratio_right: float, ratio_bottom: float,
                        horiz: list[float], vert: list[float],
                        page_w: int, page_h: int,
                        max_distance: int = 30) -> dict:
    """
    Find the nearest detected grid lines for a ratio-based cell and compute
    the actual cell boundaries. Closest line must be within max_distance px
    or the detected value is reported as None.
    Returns detected vs ratio coordinates with delta.
    """
    def nearest(sorted_vals, target):
        if not sorted_vals:
            return None
        best = min(sorted_vals, key=lambda v: abs(v - target))
        if abs(best - target) > max_distance:
            return None
        return best

    left_detected = nearest(vert, ratio_left * page_w)
    right_detected = nearest(vert, ratio_right * page_w)
    top_detected = nearest(horiz, ratio_top * page_h)
    bottom_detected = nearest(horiz, ratio_bottom * page_h)

    result = {
        "ratio": {
            "left": round(ratio_left * page_w),
            "top": round(ratio_top * page_h),
            "right": round(ratio_right * page_w),
            "bottom": round(ratio_bottom * page_h),
        },
        "detected": {
            "left": round(left_detected) if left_detected is not None else None,
            "top": round(top_detected) if top_detected is not None else None,
            "right": round(right_detected) if right_detected is not None else None,
            "bottom": round(bottom_detected) if bottom_detected is not None else None,
        }
    }
    if all(v is not None for v in result["detected"].values()):
        result["delta"] = {
            "left": result["detected"]["left"] - result["ratio"]["left"],
            "top": result["detected"]["top"] - result["ratio"]["top"],
            "right": result["detected"]["right"] - result["ratio"]["right"],
            "bottom": result["detected"]["bottom"] - result["ratio"]["bottom"],
        }
    else:
        result["delta"] = None
    return result


def generate_validation_overlay(bg_path: str, fields: list,
                                output_path: str,
                                page_w: int, page_h: int,
                                debug_dir: str = None):
    """
    Generate a validation PNG that shows:
      - Green = CV-detected cell boundaries
      - Yellow = ratio-based overlay
      - Magenta = offset annotations

    No calibration offsets are applied — the ConMas ratio formula
    (pixel = ratio × page_dim) is used as-is.
    """
    if not HAS_CV:
        print("[validate] OpenCV not available, skipping CV detection")
        return

    bg_bgr = cv2.imread(bg_path)
    if bg_bgr is None:
        print(f"[validate] Cannot read background: {bg_path}")
        return
    bg_rgb = cv2.cvtColor(bg_bgr, cv2.COLOR_BGR2RGB)

    # Detect table lines
    horiz, vert = detect_table_lines(bg_bgr, debug_dir=debug_dir)

    print(f"[validate] Detected {len(horiz)} horizontal + {len(vert)} vertical grid lines")

    # Create PIL image for annotation overlay
    img_pil = Image.fromarray(bg_rgb).convert("RGBA")
    overlay_pil = Image.new("RGBA", img_pil.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay_pil)

    try:
        font = ImageFont.truetype("arial.ttf", 16)
        small_font = ImageFont.truetype("arial.ttf", 12)
    except:
        font = ImageFont.load_default()
        small_font = font

    all_aligned = True

    for f in fields:
        # ── Extract ratio coordinates ────────────────────────────────────
        rl = getattr(f, "ratio_left", None)
        rt = getattr(f, "ratio_top", None)
        rr = getattr(f, "ratio_right", None)
        rb = getattr(f, "ratio_bottom", None)

        if rl is None:
            # Skip fields without ratio data (should not happen when called from renderer)
            continue
        rx = round(rl * page_w)
        ry = round(rt * page_h)
        rw = round((rr - rl) * page_w)
        rh = round((rb - rt) * page_h)

        # ── Draw ratio-based overlay (yellow) ────────────────────────────
        draw.rectangle([rx, ry, rx + rw, ry + rh],
                       fill=RATIO_FILL, outline=RATIO_OUTLINE, width=2)
        draw.text((rx + 2, ry + 2), f.addr, fill=RATIO_OUTLINE, font=small_font)

        # ── Find detected cell boundaries ────────────────────────────────
        if horiz and vert:
            cell = find_cell_for_ratio(rl, rt, rr, rb, horiz, vert, page_w, page_h)
            det = cell["detected"]

            if all(v is not None for v in det.values()):
                # Draw detected boundary (green)
                draw.rectangle([det["left"], det["top"],
                                det["right"], det["bottom"]],
                               fill=CV_FILL, outline=CV_OUTLINE, width=2)

                # Compute delta between ratio-based yellow box and CV-detected green box
                d_left = round(det["left"] - (rl * page_w))
                d_top = round(det["top"] - (rt * page_h))
                d_right = round(det["right"] - (rr * page_w))
                d_bottom = round(det["bottom"] - (rb * page_h))

                if any(abs(v) > 2 for v in (d_left, d_top, d_right, d_bottom)):
                    all_aligned = False
                    mid_det = ((det["left"] + det["right"]) // 2,
                               (det["top"] + det["bottom"]) // 2)
                    mid_rat = (rx + rw // 2, ry + rh // 2)
                    draw.line([mid_det, mid_rat], fill=OFFSET_COLOUR, width=2)
                    offset_str = f"ΔL={d_left:+d} ΔT={d_top:+d}"
                    draw.text((mid_rat[0] + 5, mid_rat[1] + 5),
                              offset_str, fill=OFFSET_COLOUR, font=small_font)
                else:
                    draw.text((rx + rw + 5, ry + 2),
                              "✓", fill=(0, 200, 0, 255), font=font)

    # ── Draw summary ──────────────────────────────────────────────────
    summary = f"Detected: {len(horiz)}H x {len(vert)}V grid lines"
    summary += " | ALL ALIGNED" if all_aligned else " | OFFSET DETECTED!"
    draw.text((10, 10), summary, fill=(0, 0, 0, 200) if all_aligned else OFFSET_COLOUR,
              font=font)

    # Compose and save
    img_pil = Image.alpha_composite(img_pil, overlay_pil)
    img_pil.convert("RGB").save(output_path, "PNG")
    print(f"[validate] Validation overlay saved: {output_path}")
    print(f"[validate] Grid lines detected: {len(horiz)}H x {len(vert)}V")
    result_str = "ALL ALIGNED" if all_aligned else "OFFSET FOUND!"
    print(f"[validate] Alignment result: {result_str}")


# ── CLI entry point ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    from render_service.xml_field_provider import get_cluster_fielddefs

    parser = argparse.ArgumentParser(description="Validate field overlay alignment")
    parser.add_argument("template_id", type=int, help="Template ID (e.g. 547)")
    parser.add_argument("--bg", default=None, help="Path to background PNG (auto-resolved if not given)")
    parser.add_argument("--output", default=None, help="Output validation PNG path")
    parser.add_argument("--dpi", type=int, default=300, help="DPI used for rendering (default: 300)")
    parser.add_argument("--debug-dir", default=None, help="Directory for debug images (optional)")
    args = parser.parse_args()

    preview_dir = os.path.join(_PROJECT_ROOT, "ExcelAPI", "ExcelAPI", "Preview")
    if args.bg is None:
        args.bg = os.path.join(preview_dir, f"runtime_{args.template_id}_page1.png")
    if args.output is None:
        args.output = os.path.join(preview_dir, f"runtime_{args.template_id}_validation.png")

    if args.debug_dir:
        os.makedirs(args.debug_dir, exist_ok=True)
    print(f"[validate] Template: {args.template_id}")
    print(f"[validate] Background: {args.bg}")
    print(f"[validate] Output: {args.output}")
    print(f"[validate] DPI: {args.dpi}")

    # Load fields
    fields, _ = get_cluster_fielddefs(args.template_id)
    print(f"[validate] Fields: {len(fields)}")

    # Get page dimensions from the actual PNG
    from PIL import Image as PILImage
    with PILImage.open(args.bg) as img:
        page_w, page_h = img.size
    print(f"[validate] Page: {page_w} x {page_h}")

    # Run validation (no calibration offsets)
    generate_validation_overlay(
        args.bg, fields, args.output,
        page_w, page_h,
        debug_dir=args.debug_dir,
    )
