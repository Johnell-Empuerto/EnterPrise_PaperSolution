import os, json, fitz
from CoordinateEngine.models.field_def import FieldGeometry


def generate_overlay(pdf_path, png_path, geometries, out_dir, dpi):
    doc = fitz.open(pdf_path)
    page = doc[0]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for i, g in enumerate(geometries):
        left = g.page_left_pt * zoom
        top = g.page_top_pt * zoom
        w = g.page_width_pt * zoom
        h = g.page_height_pt * zoom
        rect = fitz.Rect(left, top, left + w, top + h)
        annot = page.add_rect_annot(rect)
        annot.set_border(width=max(1, int(dpi / 150)))
        annot.set_colors(stroke=[0, 0, 1])
        annot.set_opacity(0.5)
        annot.update()

    pix = page.get_pixmap(matrix=mat)
    overlay_path = os.path.join(out_dir, "debug_overlay.png")
    pix.save(overlay_path)
    doc.close()
    print(f"  Overlay: {overlay_path}")
    return overlay_path


def generate_dump(geometries, layout, workbook, out_dir):
    data = {
        "page": {
            "width_pt": workbook.page_width_pt,
            "height_pt": workbook.page_height_pt,
            "printable_width_pt": layout.printable_w_pt,
            "printable_height_pt": layout.printable_h_pt,
            "margin_left_pt": workbook.margin_left_pt,
            "margin_right_pt": workbook.margin_right_pt,
            "margin_top_pt": workbook.margin_top_pt,
            "margin_bottom_pt": workbook.margin_bottom_pt,
        },
        "layout": {
            "origin_x_pt": layout.origin_x_pt,
            "origin_y_pt": layout.origin_y_pt,
            "scale_w": layout.scale_w,
            "scale_h": layout.scale_h,
            "eff_w_pt": layout.eff_w_pt,
            "eff_h_pt": layout.eff_h_pt,
            "paw_pt": layout.paw_pt,
            "pah_pt": layout.pah_pt,
            "centered_h": layout.centered_h,
            "centered_v": layout.centered_v,
        },
        "fields": []
    }

    for g in geometries:
        data["fields"].append({
            "addr": g.field.addr,
            "merge": g.field.merge_ref if g.field.is_merge else None,
            "worksheet_pt": {
                "range_left": round(g.range_left_pt, 4),
                "range_top": round(g.range_top_pt, 4),
                "range_width": round(g.range_width_pt, 4),
                "range_height": round(g.range_height_pt, 4),
            },
            "page_pt": {
                "left": round(g.page_left_pt, 4),
                "top": round(g.page_top_pt, 4),
                "width": round(g.page_width_pt, 4),
                "height": round(g.page_height_pt, 4),
            },
            "pixel": {
                "left": g.left_px,
                "top": g.top_px,
                "width": g.width_px,
                "height": g.height_px,
            },
            "ratio": {
                "left": g.left_ratio,
                "top": g.top_ratio,
                "width": g.width_ratio,
                "height": g.height_ratio,
            },
        })

    path = os.path.join(out_dir, "coordinate_dump.json")
    with open(path, "w") as f:
        json.dump(data, f, indent=2)
    print(f"  Dump: {path}")
    return path


def generate_report(geometries, layout, workbook, out_dir):
    lines = []
    lines.append("# Coordinate Calculation Report")
    lines.append("")
    lines.append("## Page Setup")
    lines.append(f"- Page: {workbook.page_width_pt:.0f} x {workbook.page_height_pt:.0f} pt")
    lines.append(f"- Margins: L={workbook.margin_left_pt:.2f} R={workbook.margin_right_pt:.2f} T={workbook.margin_top_pt:.2f} B={workbook.margin_bottom_pt:.2f}")
    lines.append(f"- Printable: {layout.printable_w_pt:.2f} x {layout.printable_h_pt:.2f} pt")
    lines.append(f"- Centered: H={layout.centered_h} V={layout.centered_v}")
    lines.append("")
    lines.append("## Layout")
    lines.append(f"- PA Width (PAW): {layout.paw_pt:.4f} pt")
    lines.append(f"- PA Height (PAH): {layout.pah_pt:.4f} pt")
    lines.append(f"- Effective Width (effW): {layout.eff_w_pt:.4f} pt (from PDF content bounds)")
    lines.append(f"- Effective Height (effH): {layout.eff_h_pt:.4f} pt (from PDF content bounds)")
    lines.append(f"- Scale X: {layout.scale_w:.6f}")
    lines.append(f"- Scale Y: {layout.scale_h:.6f}")
    lines.append(f"- Origin X: {layout.origin_x_pt:.4f} pt")
    lines.append(f"- Origin Y: {layout.origin_y_pt:.4f} pt")
    lines.append("")
    lines.append("## Formula")
    lines.append("`")
    lines.append("page_left = origin_x + Range.Left * scale")
    lines.append("page_top  = origin_y + Range.Top  * scale")
    lines.append("pixel     = page * (DPI / 72)")
    lines.append("ratio     = page / page_dimension")
    lines.append("`")
    lines.append("")

    for g in geometries:
        lines.append(f"### {g.field.addr}")
        if g.field.is_merge:
            lines.append(f"- Merge: {g.field.merge_ref}")
        lines.append(f"- Range: L={g.range_left_pt:.2f} T={g.range_top_pt:.2f} W={g.range_width_pt:.2f} H={g.range_height_pt:.2f} pt")
        lines.append(f"- Page:  L={g.page_left_pt:.2f} T={g.page_top_pt:.2f} W={g.page_width_pt:.2f} H={g.page_height_pt:.2f} pt")
        lines.append(f"- Pixel: L={g.left_px:.1f} T={g.top_px:.1f} W={g.width_px:.1f} H={g.height_px:.1f} px")
        lines.append(f"- Ratio: L={g.left_ratio:.6f} T={g.top_ratio:.6f} W={g.width_ratio:.6f} H={g.height_ratio:.6f}")
        lines.append("")

    path = os.path.join(out_dir, "debug_report.md")
    with open(path, "w") as f:
        f.write("\n".join(lines))
    print(f"  Report: {path}")
    return path
