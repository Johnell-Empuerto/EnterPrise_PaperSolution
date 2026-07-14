"""
Multi-source geometry cross-validation.

Compares field geometry from all three sources:
  - Workbook XML (column widths, row heights)
  - PDF vector/text content near expected field positions
  - PNG pixel content near expected field positions

Produces per-field difference reports and aggregate accuracy statistics.
"""

import math
from dataclasses import dataclass
from typing import List, Optional
from CoordinateEngine.models.field_def import FieldGeometry, FieldDef, ContentBounds


@dataclass
class FieldSourceComparison:
    addr: str
    workbook_left_pt: float
    workbook_top_pt: float
    workbook_w_pt: float
    workbook_h_pt: float
    pdf_left_pt: Optional[float] = None
    pdf_top_pt: Optional[float] = None
    pdf_w_pt: Optional[float] = None
    pdf_h_pt: Optional[float] = None
    png_left_pt: Optional[float] = None
    png_top_pt: Optional[float] = None
    png_w_pt: Optional[float] = None
    png_h_pt: Optional[float] = None
    final_left_px: float = 0.0
    final_top_px: float = 0.0
    final_w_px: float = 0.0
    final_h_px: float = 0.0
    workbook_confidence: float = 0.0
    pdf_confidence: float = 0.0
    png_confidence: float = 0.0


@dataclass
class ValidationStats:
    mean_error: float = 0.0
    median_error: float = 0.0
    max_error: float = 0.0
    rms_error: float = 0.0
    std_dev: float = 0.0
    field_count: int = 0
    errors_px: List[float] = None


def compare_sources(
    pdf_path: str,
    png_path: str,
    fields: List[FieldDef],
    geometries: List[FieldGeometry],
    workbook,
    layout,
    dpi: int,
) -> List[FieldSourceComparison]:
    """
    For each field, compare geometry from workbook XML, PDF vectors, and PNG pixels.
    """
    import fitz
    from PIL import Image
    from CoordinateEngine.utils.file_utils import col_letter

    pts_to_px = dpi / 72.0

    doc = fitz.open(pdf_path) if pdf_path else None
    png_img = Image.open(png_path).convert("RGBA") if png_path else None

    comparisons = []

    for g in geometries:
        f = g.field
        addr = f.addr

        # ── Workbook-derived position ──────────────────────────────
        wb_left = g.range_left_pt
        wb_top = g.range_top_pt
        wb_w = g.range_width_pt
        wb_h = g.range_height_pt

        # ── PDF local content bounds near field ────────────────────
        pdf_left = None
        pdf_top = None
        pdf_w = None
        pdf_h = None

        if doc:
            page = doc[0]
            # Search area: expected field region ± margin
            margin = max(20, wb_w * 2, wb_h * 2)
            search_x0 = max(0, wb_left - margin)
            search_y0 = max(0, wb_top - margin)
            search_x1 = min(
                workbook.page_width_pt, wb_left + wb_w + margin
            )
            search_y1 = min(
                workbook.page_height_pt, wb_top + wb_h + margin
            )

            local_bounds = []
            for b in page.get_text("blocks"):
                if b[6] == 0:
                    bx0, by0, bx1, by1 = b[0], b[1], b[2], b[3]
                    if (
                        bx0 >= search_x0
                        and bx1 <= search_x1
                        and by0 >= search_y0
                        and by1 <= search_y1
                    ):
                        local_bounds.append((bx0, by0, bx1, by1))

            for d in page.get_drawings():
                r = d.get("rect")
                if r:
                    if (
                        r.x0 >= search_x0
                        and r.x1 <= search_x1
                        and r.y0 >= search_y0
                        and r.y1 <= search_y1
                    ):
                        local_bounds.append((r.x0, r.y0, r.x1, r.y1))

            if local_bounds:
                pdf_left = min(b[0] for b in local_bounds)
                pdf_top = min(b[1] for b in local_bounds)
                pdf_w = max(b[2] for b in local_bounds) - pdf_left
                pdf_h = max(b[3] for b in local_bounds) - pdf_top

        # ── PNG local content bounds near field ────────────────────
        png_left = None
        png_top = None
        png_w = None
        png_h = None

        if png_img:
            px_scale = workbook.page_width_pt / png_img.width
            py_scale = workbook.page_height_pt / png_img.height
            # Convert workbook pt to pixel coords (actual rendered scale)
            sx_pt_to_px = g.left_px / g.page_left_pt if g.page_left_pt > 0 else pts_to_px
            sy_pt_to_px = g.top_px / g.page_top_pt if g.page_top_pt > 0 else pts_to_px

            # Search area in pixels
            center_x_px = int(g.left_px + g.width_px / 2)
            center_y_px = int(g.top_px + g.height_px / 2)
            margin_px = max(20, int(g.width_px * 2), int(g.height_px * 2))

            sx0 = max(0, center_x_px - margin_px)
            sy0 = max(0, center_y_px - margin_px)
            sx1 = min(png_img.width - 1, center_x_px + margin_px)
            sy1 = min(png_img.height - 1, center_y_px + margin_px)

            p = png_img.load()
            found_px = []
            step = max(1, (sx1 - sx0) // 200, (sy1 - sy0) // 200)

            for py in range(sy0, sy1 + 1, step):
                for px in range(sx0, sx1 + 1, step):
                    pr, pg, pb, pa = p[px, py]
                    if pa > 128 and (pr < 240 or pg < 240 or pb < 240):
                        found_px.append((px, py))

            if found_px:
                png_min_x = min(pt[0] for pt in found_px)
                png_min_y = min(pt[1] for pt in found_px)
                png_max_x = max(pt[0] for pt in found_px)
                png_max_y = max(pt[1] for pt in found_px)

                png_left = png_min_x / sx_pt_to_px
                png_top = png_min_y / sy_pt_to_px
                png_w = (png_max_x - png_min_x) / sx_pt_to_px
                png_h = (png_max_y - png_min_y) / sy_pt_to_px

        # ── Confidence scoring (agreement between sources) ─────────
        wb_conf = 0.95
        pdf_conf = 0.5
        png_conf = 0.5

        if pdf_left is not None:
            h_diff = abs(g.page_left_pt - pdf_left)
            if h_diff < 5:
                wb_conf = min(1.0, wb_conf + 0.03)
                pdf_conf = 0.90
            elif h_diff < 20:
                pdf_conf = 0.70
            else:
                pdf_conf = 0.40

        if png_left is not None:
            h_diff = abs(g.page_left_pt - png_left)
            if h_diff < 5:
                wb_conf = min(1.0, wb_conf + 0.02)
                png_conf = 0.90
            elif h_diff < 20:
                png_conf = 0.70
            else:
                png_conf = 0.40

        comparisons.append(
            FieldSourceComparison(
                addr=addr,
                workbook_left_pt=wb_left,
                workbook_top_pt=wb_top,
                workbook_w_pt=wb_w,
                workbook_h_pt=wb_h,
                pdf_left_pt=(
                    round(pdf_left, 4) if pdf_left is not None else None
                ),
                pdf_top_pt=(
                    round(pdf_top, 4) if pdf_top is not None else None
                ),
                pdf_w_pt=round(pdf_w, 4) if pdf_w is not None else None,
                pdf_h_pt=round(pdf_h, 4) if pdf_h is not None else None,
                png_left_pt=(
                    round(png_left, 4) if png_left is not None else None
                ),
                png_top_pt=(
                    round(png_top, 4) if png_top is not None else None
                ),
                png_w_pt=round(png_w, 4) if png_w is not None else None,
                png_h_pt=round(png_h, 4) if png_h is not None else None,
                final_left_px=g.left_px,
                final_top_px=g.top_px,
                final_w_px=g.width_px,
                final_h_px=g.height_px,
                workbook_confidence=round(wb_conf, 4),
                pdf_confidence=round(pdf_conf, 4),
                png_confidence=round(png_conf, 4),
            )
        )

    if doc:
        doc.close()
    if png_img:
        png_img.close()

    return comparisons


def compute_accuracy_stats(
    comparisons: List[FieldSourceComparison],
) -> ValidationStats:
    errors = []
    pts_to_px = 300.0 / 72.0
    for c in comparisons:
        # Use the PDF or PNG content position as ground truth,
        # compare against the engine's computed page position.
        if c.pdf_left_pt is not None and c.pdf_top_pt is not None:
            ref_x = c.pdf_left_pt + c.pdf_w_pt / 2 if c.pdf_w_pt else c.pdf_left_pt
            ref_y = c.pdf_top_pt + c.pdf_h_pt / 2 if c.pdf_h_pt else c.pdf_top_pt
        elif c.png_left_pt is not None and c.png_top_pt is not None:
            ref_x = c.png_left_pt + c.png_w_pt / 2 if c.png_w_pt else c.png_left_pt
            ref_y = c.png_top_pt + c.png_h_pt / 2 if c.png_h_pt else c.png_top_pt
        else:
            continue

        # Engine's final page position (convert from px back to pts)
        eng_x = c.final_left_px / pts_to_px + c.final_w_px / pts_to_px / 2
        eng_y = c.final_top_px / pts_to_px + c.final_h_px / pts_to_px / 2

        err = math.sqrt((eng_x - ref_x) ** 2 + (eng_y - ref_y) ** 2)
        errors.append(err)

    if not errors:
        return ValidationStats(field_count=len(comparisons))

    sorted_errs = sorted(errors)
    n = len(sorted_errs)
    mean = sum(sorted_errs) / n
    median = sorted_errs[n // 2] if n % 2 == 1 else (
        sorted_errs[n // 2 - 1] + sorted_errs[n // 2]
    ) / 2.0
    max_err = max(sorted_errs)
    rms = math.sqrt(sum(e * e for e in sorted_errs) / n)
    variance = sum((e - mean) ** 2 for e in sorted_errs) / n
    std_dev = math.sqrt(variance)

    return ValidationStats(
        mean_error=round(mean, 4),
        median_error=round(median, 4),
        max_error=round(max_err, 4),
        rms_error=round(rms, 4),
        std_dev=round(std_dev, 4),
        field_count=n,
        errors_px=[round(e, 4) for e in sorted_errs],
    )


def compute_confidence_scores(
    comparisons: List[FieldSourceComparison],
    overall_geo_error: float,
) -> dict:
    wb_scores = [c.workbook_confidence for c in comparisons]
    pdf_scores = [c.pdf_confidence for c in comparisons]
    png_scores = [c.png_confidence for c in comparisons]

    def avg(scores):
        return sum(scores) / len(scores) if scores else 0.0

    overall = (
        avg(wb_scores) * 0.4
        + avg(pdf_scores) * 0.35
        + avg(png_scores) * 0.25
    )

    if overall_geo_error > 10:
        overall *= max(0.5, 1.0 - overall_geo_error / 100.0)

    return {
        "workbook_geometry": round(avg(wb_scores) * 100, 1),
        "pdf_geometry": round(avg(pdf_scores) * 100, 1),
        "png_geometry": round(avg(png_scores) * 100, 1),
        "overall_confidence": round(min(overall * 100, 100.0), 1),
    }

