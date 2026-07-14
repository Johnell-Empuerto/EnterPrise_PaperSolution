"""
Enhanced debug visualization generator.

Produces multiple debug images highlighting different aspects of geometry:
  - debug_overlay.png     -- blue rectangles on rendered PDF
  - debug_grid.png        -- cell boundary grid overlaid on PNG
  - debug_pdf_vectors.png -- PDF vector elements highlighted
  - debug_field_centers.png -- field center markers
  - debug_errors.png      -- error heatmap / difference visualization
"""

import json
import math
import os
from typing import List, Optional, Tuple
from CoordinateEngine.models.field_def import FieldGeometry, LayoutInfo, WorkbookInfo

# Accumulator for invalid rectangles across all generator calls.
# Populated by validate_rect calls, consumed by _write_report.
_invalid_rects: List[dict] = []


def validate_rect(
    rect,
    label: str = "unknown",
    worksheet: str = "Sheet1",
    cell_ref: str = "",
    geometry: Optional[dict] = None,
) -> Tuple[bool, Optional[str]]:
    """Validate rectangle coordinates before passing to PyMuPDF.

    Accepts a fitz.Rect or any 4-element sequence (x0, y0, x1, y1).

    Returns:
        (True, None) if valid.
        (False, reason_string) if invalid.  Also appends a diagnostic
        record to the module-level _invalid_rects list.
    """
    try:
        if hasattr(rect, "x0"):
            x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
        else:
            x0, y0, x1, y1 = rect[0], rect[1], rect[2], rect[3]
    except (TypeError, IndexError, AttributeError):
        reason = "Cannot unpack coordinates"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason

    # None check
    if any(v is None for v in (x0, y0, x1, y1)):
        reason = "Contains None"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason

    # Non-numeric check
    if not all(isinstance(v, (int, float)) for v in (x0, y0, x1, y1)):
        reason = "Contains non-numeric value"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason

    # NaN check
    if any(math.isnan(v) for v in (x0, y0, x1, y1)):
        reason = "Contains NaN"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason

    # Infinity check
    if any(math.isinf(v) for v in (x0, y0, x1, y1)):
        reason = "Contains Infinity"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason

    width = x1 - x0
    height = y1 - y0

    if width <= 0:
        reason = f"Zero or negative width ({width})"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason
    if height <= 0:
        reason = f"Zero or negative height ({height})"
        _invalid_rects.append(_build_invalid_record(label, worksheet, cell_ref, reason, geometry))
        return False, reason

    return True, None


def _build_invalid_record(
    label: str,
    worksheet: str,
    cell_ref: str,
    reason: str,
    geometry: Optional[dict],
) -> dict:
    return {
        "field": label,
        "worksheet": worksheet,
        "cellReference": cell_ref,
        "reason": reason,
        "geometry": geometry or {},
    }


def _write_report(out_dir: str, total_fields: int):
    """Write invalid_rectangles.json and debug_report.md to the debug folder."""
    global _invalid_rects

    # invalid_rectangles.json
    json_path = os.path.join(out_dir, "invalid_rectangles.json")
    with open(json_path, "w") as f:
        json.dump(_invalid_rects, f, indent=2)
    print(f"  Invalid rectangles: {json_path}")

    # Count reasons
    reasons = {}
    for r in _invalid_rects:
        reason = r["reason"]
        reasons[reason] = reasons.get(reason, 0) + 1

    valid_count = total_fields - len(_invalid_rects)
    md_lines = [
        "# Debug Report",
        "",
        "## Summary",
        f"- **Total fields:** {total_fields}",
        f"- **Valid rectangles:** {valid_count}",
        f"- **Invalid rectangles:** {len(_invalid_rects)}",
        "",
    ]

    if _invalid_rects:
        md_lines.append("## Invalid Rectangles")
        md_lines.append("")
        for r in _invalid_rects:
            md_lines.append(f"- **{r['field']}** (`{r['cellReference']}`, sheet `{r['worksheet']}`): {r['reason']}")
            if r["geometry"]:
                g = r["geometry"]
                md_lines.append(f"  - page_left_pt={g.get('page_left_pt')}, page_top_pt={g.get('page_top_pt')}")
                md_lines.append(f"  - page_width_pt={g.get('page_width_pt')}, page_height_pt={g.get('page_height_pt')}")

        md_lines.append("")
        md_lines.append("## Reasons Breakdown")
        md_lines.append("")
        for reason, count in sorted(reasons.items(), key=lambda x: -x[1]):
            md_lines.append(f"- **{reason}:** {count} field(s)")

    report_path = os.path.join(out_dir, "debug_report.md")
    with open(report_path, "w") as f:
        f.write("\n".join(md_lines))
    print(f"  Debug report: {report_path}")

    # Clear accumulator for next run
    _invalid_rects = []


def _geo_to_dict(g: FieldGeometry) -> dict:
    """Convert a FieldGeometry to a plain dict for logging."""
    return {
        "page_left_pt": g.page_left_pt,
        "page_top_pt": g.page_top_pt,
        "page_width_pt": g.page_width_pt,
        "page_height_pt": g.page_height_pt,
        "range_left_pt": g.range_left_pt,
        "range_top_pt": g.range_top_pt,
        "range_width_pt": g.range_width_pt,
        "range_height_pt": g.range_height_pt,
        "left_px": g.left_px,
        "top_px": g.top_px,
        "width_px": g.width_px,
        "height_px": g.height_px,
        "left_ratio": g.left_ratio,
        "top_ratio": g.top_ratio,
        "width_ratio": g.width_ratio,
        "height_ratio": g.height_ratio,
    }


def _geometry_valid(g: FieldGeometry) -> bool:
    """Check whether a FieldGeometry has valid numeric finite dimensions."""
    attrs = [
        "left_px", "top_px", "width_px", "height_px",
        "page_left_pt", "page_top_pt", "page_width_pt", "page_height_pt",
        "left_ratio", "top_ratio", "width_ratio", "height_ratio",
        "range_left_pt", "range_top_pt", "range_width_pt", "range_height_pt",
    ]
    for attr in attrs:
        v = getattr(g, attr, None)
        if v is None:
            return False
        if not isinstance(v, (int, float)):
            return False
        if math.isnan(v) or math.isinf(v):
            return False
    if g.width_px <= 0 or g.height_px <= 0:
        return False
    if g.page_width_pt <= 0 or g.page_height_pt <= 0:
        return False
    return True


def generate_overlay(pdf_path, png_path, geometries, out_dir, dpi):
    """Blue field rectangles overlaid on the rendered PDF page."""
    import fitz
    doc = fitz.open(pdf_path)
    page = doc[0]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for g in geometries:
        addr = g.field.addr if g.field else "unknown"
        geo_dict = {
            "page_left_pt": g.page_left_pt,
            "page_top_pt": g.page_top_pt,
            "page_width_pt": g.page_width_pt,
            "page_height_pt": g.page_height_pt,
            "left_px": g.left_px,
            "top_px": g.top_px,
            "width_px": g.width_px,
            "height_px": g.height_px,
            "left_ratio": g.left_ratio,
            "top_ratio": g.top_ratio,
        }

        # rect in PDF points (NOT pixels)
        rect = fitz.Rect(
            g.page_left_pt, g.page_top_pt,
            g.page_left_pt + g.page_width_pt,
            g.page_top_pt + g.page_height_pt,
        )

        ok, _ = validate_rect(rect, label=addr, cell_ref=addr, geometry=geo_dict)
        if not ok:
            continue

        annot = page.add_rect_annot(rect)
        annot.set_border(width=max(1, int(dpi / 150)))
        annot.set_colors(stroke=[0, 0, 1])
        annot.set_opacity(0.5)
        annot.update()

    pix = page.get_pixmap(matrix=mat)
    path = os.path.join(out_dir, "debug_overlay.png")
    pix.save(path)
    doc.close()
    print(f"  Overlay: {path}")
    return path


def generate_grid(png_path, geometries, out_dir, dpi, wb=None):
    """Cell boundary grid overlaid on the rendered PNG."""
    from PIL import Image, ImageDraw

    img = Image.open(png_path).convert("RGBA")
    draw = ImageDraw.Draw(img, "RGBA")

    # Draw cell boundary rectangles with green borders
    for g in geometries:
        addr = g.field.addr if g.field else "unknown"
        if not _geometry_valid(g):
            geo_dict = _geo_to_dict(g)
            validate_rect((g.left_px, g.top_px, g.left_px + g.width_px, g.top_px + g.height_px),
                          label=addr, cell_ref=addr, geometry=geo_dict)
            continue

        x0 = int(g.left_px)
        y0 = int(g.top_px)
        x1 = int(g.left_px + g.width_px)
        y1 = int(g.top_px + g.height_px)

        for i in range(x0, x1 + 1, max(1, (x1 - x0) // 20)):
            draw.point((i, y0), fill=(0, 255, 0, 180))
            draw.point((i, y1), fill=(0, 255, 0, 180))
        for i in range(y0, y1 + 1, max(1, (y1 - y0) // 20)):
            draw.point((x0, i), fill=(0, 255, 0, 180))
            draw.point((x1, i), fill=(0, 255, 0, 180))

        # Rectangle outline
        draw.rectangle([x0, y0, x1, y1], outline=(0, 255, 0, 200), width=2)

        # Label with cell address
        label = addr
        draw.text((x0 + 2, y0 + 2), label, fill=(0, 255, 0, 220))

    path = os.path.join(out_dir, "debug_grid.png")
    img.save(path)
    img.close()
    print(f"  Grid: {path}")
    return path


def generate_pdf_vectors(pdf_path, geometries, out_dir, dpi):
    """PDF vector elements highlighted on the rendered page."""
    import fitz
    doc = fitz.open(pdf_path)
    page = doc[0]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    # Highlight all text blocks in green
    for b in page.get_text("blocks"):
        if b[6] == 0:
            rect = fitz.Rect(b[0], b[1], b[2], b[3])
            ok, _ = validate_rect(rect, label="pdf_text_block", cell_ref=f"text_{b[0]:.1f}_{b[1]:.1f}")
            if not ok:
                continue
            annot = page.add_rect_annot(rect)
            annot.set_border(width=1)
            annot.set_colors(stroke=[0, 1, 0])
            annot.set_opacity(0.3)
            annot.update()

    # Highlight all drawings in blue
    for d in page.get_drawings():
        r = d.get("rect")
        if r:
            rect = r
            ok, _ = validate_rect(rect, label="pdf_drawing", cell_ref=f"drawing_{id(d)}")
            if not ok:
                continue
            annot = page.add_rect_annot(rect)
            annot.set_border(width=1)
            annot.set_colors(stroke=[0, 0, 1])
            annot.set_opacity(0.3)
            annot.update()

    # Highlight all images in red
    for img_info in page.get_images(full=True):
        try:
            ir = page.get_image_bbox(img_info)
            if ir:
                rect = ir
                ok, _ = validate_rect(rect, label="pdf_image", cell_ref=f"img_{id(img_info)}")
                if not ok:
                    continue
                annot = page.add_rect_annot(rect)
                annot.set_border(width=2)
                annot.set_colors(stroke=[1, 0, 0])
                annot.set_opacity(0.4)
                annot.update()
        except Exception:
            pass

    pix = page.get_pixmap(matrix=mat)
    path = os.path.join(out_dir, "debug_pdf_vectors.png")
    pix.save(path)
    doc.close()
    print(f"  PDF vectors: {path}")
    return path


def generate_field_centers(png_path, geometries, out_dir, dpi):
    """Field center markers overlaid on the rendered PNG."""
    from PIL import Image, ImageDraw

    img = Image.open(png_path).convert("RGBA")
    draw = ImageDraw.Draw(img, "RGBA")

    for g in geometries:
        addr = g.field.addr if g.field else "unknown"
        if not _geometry_valid(g):
            geo_dict = _geo_to_dict(g)
            validate_rect((g.left_px, g.top_px, g.left_px + g.width_px, g.top_px + g.height_px),
                          label=addr, cell_ref=addr, geometry=geo_dict)
            continue

        cx = int(g.left_px + g.width_px / 2)
        cy = int(g.top_px + g.height_px / 2)

        # Crosshair at center
        cross_size = max(6, int(min(g.width_px, g.height_px) * 0.15))
        color = (255, 0, 0, 220)

        # Horizontal line
        for dx in range(-cross_size, cross_size + 1):
            px = cx + dx
            if 0 <= px < img.width:
                draw.point((px, cy), fill=color)

        # Vertical line
        for dy in range(-cross_size, cross_size + 1):
            py = cy + dy
            if 0 <= py < img.height:
                draw.point((cx, py), fill=color)

        # Circle around center
        r = max(4, cross_size // 2)
        draw.ellipse(
            [cx - r, cy - r, cx + r, cy + r],
            outline=color,
            width=1,
        )

        # Label
        label = f"{addr}"
        draw.text((cx + r + 3, cy - 8), label, fill=(255, 0, 0, 220))

    path = os.path.join(out_dir, "debug_field_centers.png")
    img.save(path)
    img.close()
    print(f"  Field centers: {path}")
    return path


def generate_errors(png_path, geometries, comparisons, out_dir, dpi):
    """Error heatmap -- highlights fields with larger discrepancies."""
    from PIL import Image, ImageDraw

    img = Image.open(png_path).convert("RGBA")
    draw = ImageDraw.Draw(img, "RGBA")

    # Build lookup by address
    comp_map = {}
    for c in comparisons:
        comp_map[c.addr] = c

    for g in geometries:
        addr = g.field.addr if g.field else "unknown"
        if not _geometry_valid(g):
            geo_dict = _geo_to_dict(g)
            validate_rect((g.left_px, g.top_px, g.left_px + g.width_px, g.top_px + g.height_px),
                          label=addr, cell_ref=addr, geometry=geo_dict)
            continue

        x0 = int(g.left_px)
        y0 = int(g.top_px)
        x1 = int(g.left_px + g.width_px)
        y1 = int(g.top_px + g.height_px)

        comp = comp_map.get(addr)
        if comp and comp.pdf_left_pt is not None:
            # Error = difference between final page position and PDF content position
            h_err = abs(g.page_left_pt - comp.pdf_left_pt)
            v_err = abs(g.page_top_pt - comp.pdf_top_pt)
            max_err = max(h_err, v_err)
        else:
            max_err = 0

        # Color by error severity
        if max_err < 2:
            overlay_color = (0, 255, 0, 60)  # Green = good
        elif max_err < 8:
            overlay_color = (255, 255, 0, 80)  # Yellow = moderate
        elif max_err < 20:
            overlay_color = (255, 165, 0, 100)  # Orange = poor
        else:
            overlay_color = (255, 0, 0, 120)  # Red = bad

        # Fill rectangle with error color
        draw.rectangle([x0, y0, x1, y1], fill=overlay_color)

        # Error text
        label = f"{addr}: {max_err:.1f}pt"
        draw.text((x0 + 2, y0 + 2), label, fill=(0, 0, 0, 220))

        # PDF content bounds in blue (if available)
        if comp and comp.pdf_left_pt is not None:
            pts_to_px = dpi / 72.0
            px0 = int(comp.pdf_left_pt * pts_to_px)
            py0 = int(comp.pdf_top_pt * pts_to_px)
            px1 = int((comp.pdf_left_pt + comp.pdf_w_pt) * pts_to_px) if comp.pdf_w_pt else x1
            py1 = int((comp.pdf_top_pt + comp.pdf_h_pt) * pts_to_px) if comp.pdf_h_pt else y1
            draw.rectangle([px0, py0, px1, py1], outline=(0, 0, 255, 200), width=1)

    path = os.path.join(out_dir, "debug_errors.png")
    img.save(path)
    img.close()
    print(f"  Errors: {path}")
    return path


def generate_all(
    pdf_path: str,
    png_path: str,
    geometries: List[FieldGeometry],
    comparisons,
    layout: LayoutInfo,
    wb,
    out_dir: str,
    dpi: int,
):
    """Generate all debug visualization images."""
    results = {}

    results["overlay"] = generate_overlay(pdf_path, png_path, geometries, out_dir, dpi)
    results["grid"] = generate_grid(png_path, geometries, out_dir, dpi, wb)
    results["pdf_vectors"] = generate_pdf_vectors(pdf_path, geometries, out_dir, dpi)
    results["field_centers"] = generate_field_centers(png_path, geometries, out_dir, dpi)

    if comparisons:
        results["errors"] = generate_errors(png_path, geometries, comparisons, out_dir, dpi)

    # Write validation report
    _write_report(out_dir, total_fields=len(geometries))

    return results
