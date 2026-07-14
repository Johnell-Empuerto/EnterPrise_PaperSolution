"""
Per-field geometry trace report generator.

For every field, produces a complete derivation path showing how each
coordinate was computed — from workbook XML → range position →
page position → pixel → ratio — with every intermediate value documented.
"""

from typing import List
from CoordinateEngine.models.field_def import FieldGeometry, FieldDef, WorkbookInfo, LayoutInfo


def trace_field(field: FieldDef, geo: FieldGeometry, wb: WorkbookInfo, layout: LayoutInfo, dpi: int) -> dict:
    """
    Trace the full derivation of a single field's geometry.
    Returns a dict with every intermediate step labeled.
    """
    pts_to_px = dpi / 72.0

    steps = []

    step_id = 0

    def add_step(desc: str, formula: str, value, unit: str, source: str):
        nonlocal step_id
        step_id += 1
        steps.append({
            "step": step_id,
            "description": desc,
            "formula": formula,
            "value": round(value, 6) if isinstance(value, float) else value,
            "unit": unit,
            "source": source,
        })

    # 1. Column/Row info
    add_step(
        f"Field {field.addr}",
        f"col={field.col}, row={field.row}, col_end={field.col_end}, row_end={field.row_end}",
        f"{field.addr}",
        "cell",
        "workbook XML"
    )

    # 2. Column widths
    col_widths = {}
    for c in range(field.col, field.col_end + 1):
        cw = wb.col_widths_pt.get(c, None)
        if cw is None:
            cw = (wb.default_col_width * 7.33 + 5.0) * 72.0 / 96.0
        col_widths[c] = cw
        add_step(
            f"Col {c} width",
            f"col_widths_pt[{c}]" + (f" = {cw:.4f} pt" if cw else f" = default = {cw:.4f} pt"),
            cw,
            "pt",
            "workbook XML"
        )

    # 3. Row heights
    row_heights = {}
    for r in range(field.row, field.row_end + 1):
        rh = wb.row_heights_pt.get(r, wb.default_row_height)
        row_heights[r] = rh
        add_step(
            f"Row {r} height",
            f"row_heights_pt[{r}]" + (f" = {rh:.4f}" if rh else f" = default = {rh:.4f}"),
            rh,
            "pt",
            "workbook XML"
        )

    # 4. Range left: sum of previous columns
    prev_cols = []
    for c in range(1, field.col):
        cw = wb.col_widths_pt.get(c, None)
        if cw is None:
            cw = (wb.default_col_width * 7.33 + 5.0) * 72.0 / 96.0
        prev_cols.append(cw)

    range_left_sum = sum(prev_cols)
    add_step(
        "Range left = Σ col widths before field",
        f"Σ cols 1..{field.col - 1}" + (f" = {' + '.join(f'{w:.4f}' for w in prev_cols)}" if prev_cols else " = 0 (first column)"),
        range_left_sum,
        "pt",
        "workbook XML → computed"
    )

    # 5. Range top: sum of previous rows
    prev_rows = []
    for r in range(1, field.row):
        rh = wb.row_heights_pt.get(r, wb.default_row_height)
        prev_rows.append(rh)

    range_top_sum = sum(prev_rows)
    add_step(
        "Range top = Σ row heights before field",
        f"Σ rows 1..{field.row - 1}" + (f" = {' + '.join(f'{w:.4f}' for w in prev_rows)}" if prev_rows else " = 0 (first row)"),
        range_top_sum,
        "pt",
        "workbook XML → computed"
    )

    # 6. Range width: sum of field columns
    field_cols = list(col_widths.values())
    range_w_sum = sum(field_cols)
    add_step(
        "Range width = Σ col widths of field",
        f"Σ cols {field.col}..{field.col_end}" + (f" = {' + '.join(f'{w:.4f}' for w in field_cols)}" if field_cols else ""),
        range_w_sum,
        "pt",
        "workbook XML → computed"
    )

    # 7. Range height: sum of field rows
    field_rows = list(row_heights.values())
    range_h_sum = sum(field_rows)
    add_step(
        "Range height = Σ row heights of field",
        f"Σ rows {field.row}..{field.row_end}" + (f" = {' + '.join(f'{w:.4f}' for w in field_rows)}" if field_rows else ""),
        range_h_sum,
        "pt",
        "workbook XML → computed"
    )

    # 8. Page setup & layout
    add_step(
        "Page width",
        f"paper_size → page_width_pt",
        wb.page_width_pt,
        "pt",
        "Excel Page Setup"
    )
    add_step(
        "Page height",
        f"paper_size → page_height_pt",
        wb.page_height_pt,
        "pt",
        "Excel Page Setup"
    )
    add_step(
        "Printable width = page - Lmargin - Rmargin",
        f"{wb.page_width_pt:.2f} - {wb.margin_left_pt:.2f} - {wb.margin_right_pt:.2f}",
        layout.printable_w_pt,
        "pt",
        "computed from page setup"
    )
    add_step(
        "Printable height = page - Tmargin - Bmargin",
        f"{wb.page_height_pt:.2f} - {wb.margin_top_pt:.2f} - {wb.margin_bottom_pt:.2f}",
        layout.printable_h_pt,
        "pt",
        "computed from page setup"
    )

    # 9. PA dimensions
    add_step(
        "PA Width (PAW) = Σ col widths of print area",
        f"= {layout.paw_pt:.4f}",
        layout.paw_pt,
        "pt",
        "workbook XML"
    )
    add_step(
        "PA Height (PAH) = Σ row heights of print area",
        f"= {layout.pah_pt:.4f}",
        layout.pah_pt,
        "pt",
        "workbook XML"
    )

    # 10. PDF content analysis
    add_step(
        "Effective width (effW) from PDF content analysis",
        f"centered_h={layout.centered_h} → effW={layout.eff_w_pt:.4f}",
        layout.eff_w_pt,
        "pt",
        "PDF content bounds"
    )
    add_step(
        "Effective height (effH) from PDF content analysis",
        f"centered_v={layout.centered_v} → effH={layout.eff_h_pt:.4f}",
        layout.eff_h_pt,
        "pt",
        "PDF content bounds"
    )

    # 11. Scale
    add_step(
        "Scale X = effW / PAW",
        f"{layout.eff_w_pt:.6f} / {layout.paw_pt:.6f}",
        layout.scale_w,
        "",
        "computed"
    )
    add_step(
        "Scale Y = effH / PAH",
        f"{layout.eff_h_pt:.6f} / {layout.pah_pt:.6f}",
        layout.scale_h,
        "",
        "computed"
    )

    # 12. Origin
    if layout.centered_h:
        origin_formula = f"Lmargin + (printable_w - effW)/2 = {wb.margin_left_pt:.2f} + ({layout.printable_w_pt:.2f} - {layout.eff_w_pt:.2f})/2"
    else:
        origin_formula = f"Lmargin = {wb.margin_left_pt:.2f}"

    add_step(
        "Origin X = " + ("Lmargin + (printable_w - effW)/2" if layout.centered_h else "Lmargin"),
        origin_formula,
        layout.origin_x_pt,
        "pt",
        "computed"
    )

    if layout.centered_v:
        origin_formula = f"Tmargin + (printable_h - effH)/2 = {wb.margin_top_pt:.2f} + ({layout.printable_h_pt:.2f} - {layout.eff_h_pt:.2f})/2"
    else:
        origin_formula = f"Tmargin = {wb.margin_top_pt:.2f}"

    add_step(
        "Origin Y = " + ("Tmargin + (printable_h - effH)/2" if layout.centered_v else "Tmargin"),
        origin_formula,
        layout.origin_y_pt,
        "pt",
        "computed"
    )

    # 13. Page position
    page_left_formula = f"origin_x + range_left * scale_w = {layout.origin_x_pt:.4f} + {geo.range_left_pt:.4f} * {layout.scale_w:.6f}"
    add_step(
        "Page left = origin_x + range_left × scale_w",
        page_left_formula,
        geo.page_left_pt,
        "pt",
        "computed"
    )

    page_top_formula = f"origin_y + range_top * scale_h = {layout.origin_y_pt:.4f} + {geo.range_top_pt:.4f} * {layout.scale_h:.6f}"
    add_step(
        "Page top = origin_y + range_top × scale_h",
        page_top_formula,
        geo.page_top_pt,
        "pt",
        "computed"
    )

    add_step(
        "Page width = range_width × scale_w",
        f"{geo.range_width_pt:.4f} × {layout.scale_w:.6f}",
        geo.page_width_pt,
        "pt",
        "computed"
    )

    add_step(
        "Page height = range_height × scale_h",
        f"{geo.range_height_pt:.4f} × {layout.scale_h:.6f}",
        geo.page_height_pt,
        "pt",
        "computed"
    )

    # 14. Pixel coordinates
    add_step(
        "Pixel left = page_left × (dpi/72)",
        f"{geo.page_left_pt:.4f} × {pts_to_px:.6f}",
        geo.left_px,
        "px",
        "computed"
    )
    add_step(
        "Pixel top = page_top × (dpi/72)",
        f"{geo.page_top_pt:.4f} × {pts_to_px:.6f}",
        geo.top_px,
        "px",
        "computed"
    )
    add_step(
        "Pixel width = page_width × (dpi/72)",
        f"{geo.page_width_pt:.4f} × {pts_to_px:.6f}",
        geo.width_px,
        "px",
        "computed"
    )
    add_step(
        "Pixel height = page_height × (dpi/72)",
        f"{geo.page_height_pt:.4f} × {pts_to_px:.6f}",
        geo.height_px,
        "px",
        "computed"
    )

    # 15. Ratios
    add_step(
        "Left ratio = pixel_left / page_width_px",
        f"{geo.left_px:.2f} / {geo.left_px / geo.left_ratio:.2f}" if geo.left_ratio > 0 else f"{geo.left_px:.2f} / page_width",
        geo.left_ratio,
        "",
        "computed"
    )
    add_step(
        "Top ratio = pixel_top / page_height_px",
        f"{geo.top_px:.2f} / page_height",
        geo.top_ratio,
        "",
        "computed"
    )

    return {
        "field": field.addr,
        "is_merged": field.is_merge,
        "merge_range": field.merge_ref if field.is_merge else None,
        "worksheet_position": {
            "left_pt": geo.range_left_pt,
            "top_pt": geo.range_top_pt,
            "width_pt": geo.range_width_pt,
            "height_pt": geo.range_height_pt,
        },
        "page_position": {
            "left_pt": geo.page_left_pt,
            "top_pt": geo.page_top_pt,
            "width_pt": geo.page_width_pt,
            "height_pt": geo.page_height_pt,
        },
        "pixel_position": {
            "left_px": geo.left_px,
            "top_px": geo.top_px,
            "width_px": geo.width_px,
            "height_px": geo.height_px,
        },
        "ratios": {
            "left": geo.left_ratio,
            "top": geo.top_ratio,
            "width": geo.width_ratio,
            "height": geo.height_ratio,
        },
        "derivation_steps": steps,
    }


def generate_geometry_trace(
    fields: List[FieldDef],
    geometries: List[FieldGeometry],
    wb: WorkbookInfo,
    layout: LayoutInfo,
    dpi: int,
    out_dir: str,
) -> str:
    """
    Generate a complete geometry_trace.md report.
    """
    traces = [
        trace_field(f, g, wb, layout, dpi)
        for f, g in zip(fields, geometries)
    ]

    lines = []
    lines.append("# Geometry Trace Report")
    lines.append("")
    lines.append("## Summary")
    lines.append(f"- Workbook: {wb.name}")
    lines.append(f"- Page: {wb.page_width_pt:.0f} × {wb.page_height_pt:.0f} pt")
    lines.append(f"- Fields: {len(traces)}")
    lines.append(f"- DPI: {dpi}")
    lines.append("")
    lines.append("## Formula")
    lines.append("```")
    lines.append("range_left   = Σ column_widths(1 .. col-1)")
    lines.append("range_top    = Σ row_heights(1 .. row-1)")
    lines.append("range_width  = Σ column_widths(col .. col_end)")
    lines.append("range_height = Σ row_heights(row .. row_end)")
    lines.append("")
    lines.append("page_left    = origin_x + range_left  × scale_w")
    lines.append("page_top     = origin_y + range_top   × scale_h")
    lines.append("page_width   = range_width  × scale_w")
    lines.append("page_height  = range_height × scale_h")
    lines.append("")
    lines.append("pixel_left   = page_left   × (DPI / 72)")
    lines.append("pixel_top    = page_top    × (DPI / 72)")
    lines.append("pixel_width  = page_width  × (DPI / 72)")
    lines.append("pixel_height = page_height × (DPI / 72)")
    lines.append("")
    lines.append("left_ratio   = pixel_left   / page_width_px")
    lines.append("top_ratio    = pixel_top    / page_height_px")
    lines.append("width_ratio  = pixel_width  / page_width_px")
    lines.append("height_ratio = pixel_height / page_height_px")
    lines.append("```")
    lines.append("")

    for t in traces:
        lines.append(f"## {t['field']}")
        if t['is_merged']:
            lines.append(f"**Merged:** {t['merge_range']}")
        lines.append("")
        lines.append("### Worksheet Position")
        lines.append(f"| Property | Value (pt) | Source |")
        lines.append(f"|----------|------------|--------|")
        ws = t['worksheet_position']
        lines.append(f"| Range Left | {ws['left_pt']:.4f} | Workbook XML → column width sum |")
        lines.append(f"| Range Top | {ws['top_pt']:.4f} | Workbook XML → row height sum |")
        lines.append(f"| Range Width | {ws['width_pt']:.4f} | Workbook XML → column width sum |")
        lines.append(f"| Range Height | {ws['height_pt']:.4f} | Workbook XML → row height sum |")
        lines.append("")
        lines.append("### Page Position")
        lines.append(f"| Property | Value (pt) | Formula |")
        lines.append(f"|----------|------------|---------|")
        pp = t['page_position']
        lines.append(f"| Page Left | {pp['left_pt']:.4f} | origin_x + range_left × scale_w |")
        lines.append(f"| Page Top | {pp['top_pt']:.4f} | origin_y + range_top × scale_h |")
        lines.append(f"| Page Width | {pp['width_pt']:.4f} | range_width × scale_w |")
        lines.append(f"| Page Height | {pp['height_pt']:.4f} | range_height × scale_h |")
        lines.append("")
        lines.append("### Pixel Position")
        lines.append(f"| Property | Value (px) | Formula |")
        lines.append(f"|----------|------------|---------|")
        px = t['pixel_position']
        lines.append(f"| Pixel Left | {px['left_px']:.2f} | page_left × (DPI/72) |")
        lines.append(f"| Pixel Top | {px['top_px']:.2f} | page_top × (DPI/72) |")
        lines.append(f"| Pixel Width | {px['width_px']:.2f} | page_width × (DPI/72) |")
        lines.append(f"| Pixel Height | {px['height_px']:.2f} | page_height × (DPI/72) |")
        lines.append("")
        lines.append("### Ratios")
        r = t['ratios']
        lines.append(f"| Property | Value |")
        lines.append(f"|----------|-------|")
        lines.append(f"| Left Ratio | {r['left']:.6f} |")
        lines.append(f"| Top Ratio | {r['top']:.6f} |")
        lines.append(f"| Width Ratio | {r['width']:.6f} |")
        lines.append(f"| Height Ratio | {r['height']:.6f} |")
        lines.append("")
        lines.append("### Derivation Steps")
        lines.append("| # | Description | Formula | Value | Unit | Source |")
        lines.append("|---|-------------|---------|-------|------|--------|")
        for s in t['derivation_steps']:
            lines.append(f"| {s['step']} | {s['description']} | {s['formula']} | {s['value']} | {s['unit']} | {s['source']} |")
        lines.append("")

    import os
    path = os.path.join(out_dir, "geometry_trace.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"  Geometry trace: {path}")
    return path
