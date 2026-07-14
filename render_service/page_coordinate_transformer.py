
"""Deterministic page-coordinate transformer matching LibreOffice''s printed page algorithm."""

import math
from dataclasses import dataclass
from typing import Optional


@dataclass
class PageLayout:
    origin_x_pt: float
    origin_y_pt: float
    scale_w: float
    scale_h: float
    content_width_pt: float
    content_height_pt: float
    printable_width_pt: float
    printable_height_pt: float
    page_width_pt: float
    page_height_pt: float
    dpi: int = 300

    @property
    def origin_x_px(self):
        return self.origin_x_pt * self.dpi / 72.0

    @property
    def origin_y_px(self):
        return self.origin_y_pt * self.dpi / 72.0

    @property
    def scale_w_px(self):
        return self.scale_w * self.dpi / 72.0

    @property
    def scale_h_px(self):
        return self.scale_h * self.dpi / 72.0


def _sum_cols(wb, start, end, sheet_index=0):
    return sum(_col_pt(wb, c, sheet_index) for c in range(start, end + 1))


def _col_pt(wb, c, sheet_index=0):
    raw = wb.col_widths_pt.get((c, sheet_index), None)
    if raw is None:
        raw = wb.col_widths_pt.get(c, None)
    if raw is None:
        raw = wb.col_widths_pt.get((c, 0), None)
    if raw is not None:
        return raw
    from CoordinateEngine.utils.file_utils import char_to_pt
    return char_to_pt(wb.default_col_width)


def _sum_rows(wb, start, end, sheet_index=0):
    return sum(_row_pt(wb, r, sheet_index) for r in range(start, end + 1))


def _row_pt(wb, r, sheet_index=0):
    rv = wb.row_heights_pt.get((r, sheet_index), None)
    if rv is None:
        rv = wb.row_heights_pt.get(r, None)
    if rv is None:
        rv = wb.row_heights_pt.get((r, 0), None)
    if rv is not None:
        return rv
    return wb.default_row_height


_INCH_TO_PT = 72.0


_INCH_THRESHOLD = 20.0  # max reasonable margin in inches is ~2" = 144pt; anything < 20pt is likely inches

def _margin_pt(value: Optional[float], default_pt: float) -> float:
    """Convert margin to points.  XML stores margins in inches; the workbook_reader
    stores them raw (in inches) in fields named *_pt.  Detect by size."""
    if value is not None:
        if value < _INCH_THRESHOLD:
            return value * _INCH_TO_PT  # stored as inches
        return value  # already in points
    return default_pt


def resolve_print_area(wb, sheet_index=0):
    """Resolve the print area for a given sheet."""
    from CoordinateEngine.utils.file_utils import parse_range
    if wb.print_area_addr:
        pa = parse_range(wb.print_area_addr)
        if pa:
            return pa
    max_col = 1
    for key in wb.col_widths_pt:
        if isinstance(key, int):
            max_col = max(max_col, key)
    max_row = 1
    for key in wb.row_heights_pt:
        if isinstance(key, int):
            max_row = max(max_row, key)
    return 1, 1, max_col, max_row


def compute_page_layout(wb, sheet_index=0, dpi=300):
    """Compute the deterministic page layout matching LibreOffice''s algorithm."""
    sheet = wb.sheets[sheet_index] if sheet_index < len(wb.sheets) else None

    # Page dimensions (already in points)
    page_w = sheet.page_width_pt if sheet else wb.page_width_pt
    page_h = sheet.page_height_pt if sheet else wb.page_height_pt

    # Margins: XML stores in inches, convert to points
    if sheet:
        margin_left = _margin_pt(sheet.margin_left_pt if sheet.margin_left_pt != 51.024 else None, 51.024)
        margin_right = _margin_pt(sheet.margin_right_pt if sheet.margin_right_pt != 51.024 else None, 51.024)
        margin_top = _margin_pt(sheet.margin_top_pt if sheet.margin_top_pt != 53.858 else None, 53.858)
        margin_bottom = _margin_pt(sheet.margin_bottom_pt if sheet.margin_bottom_pt != 53.858 else None, 53.858)
    else:
        margin_left = _margin_pt(None, 51.024)
        margin_right = _margin_pt(None, 51.024)
        margin_top = _margin_pt(None, 53.858)
        margin_bottom = _margin_pt(None, 53.858)

    printable_w = page_w - margin_left - margin_right
    printable_h = page_h - margin_top - margin_bottom

    # Print area
    c1, r1, c2, r2 = resolve_print_area(wb, sheet_index)
    content_w = _sum_cols(wb, c1, c2, sheet_index)
    content_h = _sum_rows(wb, r1, r2, sheet_index)

    # Determine scaling
    fit_w = sheet.fit_to_pages_wide if sheet else 0
    fit_h = sheet.fit_to_pages_tall if sheet else 0
    zoom = sheet.zoom if sheet else 0

    if fit_w > 0 and fit_h > 0:
        scale_w = printable_w / (content_w / fit_w) if content_w > 0 else 1.0
        scale_h = printable_h / (content_h / fit_h) if content_h > 0 else 1.0
        scale = min(scale_w, scale_h)
    elif zoom > 0:
        scale = zoom / 100.0
    else:
        scale_w = printable_w / content_w if content_w > 0 else 1.0
        scale_h = printable_h / content_h if content_h > 0 else 1.0
        scale = min(scale_w, scale_h)

    # Apply centering
    scaled_content_w = content_w * scale
    scaled_content_h = content_h * scale

    center_h = sheet.center_horizontally if sheet else False
    center_v = sheet.center_vertically if sheet else False

    ox = margin_left
    oy = margin_top
    if center_h:
        ox += (printable_w - scaled_content_w) / 2.0
    if center_v:
        oy += (printable_h - scaled_content_h) / 2.0

    return PageLayout(
        origin_x_pt=ox,
        origin_y_pt=oy,
        scale_w=scale,
        scale_h=scale,
        content_width_pt=content_w,
        content_height_pt=content_h,
        printable_width_pt=printable_w,
        printable_height_pt=printable_h,
        page_width_pt=page_w,
        page_height_pt=page_h,
        dpi=dpi,
    )


def cell_range_to_page_rect(layout: PageLayout, wb, col, row, col_end, row_end, sheet_index=0):
    """Convert an Excel cell range to page coordinates in points.
    Matches the algorithm from CoordinateEngine.geometry_engine.compute_field_geometry."""
    range_left = _sum_cols(wb, 1, col - 1, sheet_index) if col > 1 else 0.0
    range_top = _sum_rows(wb, 1, row - 1, sheet_index) if row > 1 else 0.0
    range_w = _sum_cols(wb, col, col_end, sheet_index)
    range_h = _sum_rows(wb, row, row_end, sheet_index)

    pl = layout.origin_x_pt + range_left * layout.scale_w
    pt = layout.origin_y_pt + range_top * layout.scale_h
    pw = range_w * layout.scale_w
    ph = range_h * layout.scale_h

    pts_to_px = layout.dpi / 72.0
    return {
        "page_left_pt": pl,
        "page_top_pt": pt,
        "page_width_pt": pw,
        "page_height_pt": ph,
        "left_px": round(pl * pts_to_px, 1),
        "top_px": round(pt * pts_to_px, 1),
        "width_px": round(pw * pts_to_px, 1),
        "height_px": round(ph * pts_to_px, 1),
    }
