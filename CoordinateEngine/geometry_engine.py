from CoordinateEngine.models.field_def import FieldDef, WorkbookInfo, ContentBounds, LayoutInfo, FieldGeometry
from CoordinateEngine.pdf_reader import detect_centering, compute_effective_dims
from CoordinateEngine.utils.file_utils import char_to_pt


def build_layout(workbook, content, dpi):
    printable_w = workbook.page_width_pt - workbook.margin_left_pt - workbook.margin_right_pt
    printable_h = workbook.page_height_pt - workbook.margin_top_pt - workbook.margin_bottom_pt

    pa = _resolve_pa(workbook)
    paw = _sum_cols(workbook, 1, pa[2])
    pah = _sum_rows(workbook, 1, pa[3])

    c_h, c_v = detect_centering(
        content, workbook.margin_left_pt, workbook.margin_top_pt,
        workbook.margin_right_pt, workbook.margin_bottom_pt
    )
    eff_w, eff_h = compute_effective_dims(
        content, workbook.margin_left_pt, workbook.margin_top_pt,
        workbook.margin_right_pt, workbook.margin_bottom_pt, c_h, c_v
    )

    scale_w = eff_w / paw if paw > 0 else 1.0
    scale_h = eff_h / pah if pah > 0 else 1.0

    ox = workbook.margin_left_pt
    oy = workbook.margin_top_pt
    if c_h:
        ox += (printable_w - eff_w) / 2.0
    if c_v:
        oy += (printable_h - eff_h) / 2.0

    return LayoutInfo(
        origin_x_pt=ox, origin_y_pt=oy,
        scale_w=scale_w, scale_h=scale_h,
        eff_w_pt=eff_w, eff_h_pt=eff_h,
        paw_pt=paw, pah_pt=pah,
        printable_w_pt=printable_w, printable_h_pt=printable_h,
        centered_h=c_h, centered_v=c_v,
    )


def compute_field_geometry(field, layout, workbook, dpi):
    pts_to_px = dpi / 72.0

    range_left = _sum_cols(workbook, 1, field.col - 1) if field.col > 1 else 0.0
    range_top = _sum_rows(workbook, 1, field.row - 1) if field.row > 1 else 0.0
    range_w = _sum_cols(workbook, field.col, field.col_end)
    range_h = _sum_rows(workbook, field.row, field.row_end)

    pl = layout.origin_x_pt + range_left * layout.scale_w
    pt = layout.origin_y_pt + range_top * layout.scale_h
    pw = range_w * layout.scale_w
    ph = range_h * layout.scale_h

    return FieldGeometry(
        field=field,
        range_left_pt=range_left, range_top_pt=range_top,
        range_width_pt=range_w, range_height_pt=range_h,
        page_left_pt=pl, page_top_pt=pt,
        page_width_pt=pw, page_height_pt=ph,
        left_px=round(pl * pts_to_px, 2),
        top_px=round(pt * pts_to_px, 2),
        width_px=round(pw * pts_to_px, 2),
        height_px=round(ph * pts_to_px, 2),
        left_ratio=round(pl / workbook.page_width_pt, 6) if workbook.page_width_pt > 0 else 0.0,
        top_ratio=round(pt / workbook.page_height_pt, 6) if workbook.page_height_pt > 0 else 0.0,
        width_ratio=round(pw / workbook.page_width_pt, 6) if workbook.page_width_pt > 0 else 0.0,
        height_ratio=round(ph / workbook.page_height_pt, 6) if workbook.page_height_pt > 0 else 0.0,
    )


def _resolve_pa(wb):
    if wb.print_area_addr:
        from CoordinateEngine.utils.file_utils import parse_range
        p = parse_range(wb.print_area_addr)
        if p:
            return p
    # No print area configured — derive from available column/row keys
    max_col = 1
    for key in wb.col_widths_pt:
        if isinstance(key, int):
            if key > max_col:
                max_col = key
    max_row = 1
    for key in wb.row_heights_pt:
        if isinstance(key, int):
            if key > max_row:
                max_row = key
    return 1, 1, max_col, max_row


def _col_pt(wb, c):
    raw = wb.col_widths_pt.get(c)
    if raw is not None:
        return raw
    return char_to_pt(wb.default_col_width)


def _row_pt(wb, r):
    return wb.row_heights_pt.get(r, wb.default_row_height)


def _sum_cols(wb, start, end):
    return sum(_col_pt(wb, c) for c in range(start, end + 1))


def _sum_rows(wb, start, end):
    return sum(_row_pt(wb, r) for r in range(start, end + 1))
