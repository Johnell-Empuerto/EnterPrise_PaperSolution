"""
Engine confidence scoring.

Computes confidence metrics based on agreement between multiple geometry
sources (workbook XML, PDF, PNG) and internal consistency checks.

Scores are diagnostics only — not included in the frontend runtime.json schema.
"""

from dataclasses import dataclass
from typing import List, Optional
from CoordinateEngine.models.field_def import FieldGeometry, LayoutInfo, WorkbookInfo


@dataclass
class ConfidenceReport:
    workbook_geometry: float  # 0-100
    pdf_geometry: float  # 0-100
    png_geometry: float  # 0-100
    overall_confidence: float  # 0-100
    warnings: List[str]
    strengths: List[str]


def score(
    geometries: List[FieldGeometry],
    layout: LayoutInfo,
    wb: WorkbookInfo,
    comparisons,
    dpi: int,
) -> ConfidenceReport:
    """
    Compute confidence scores by analyzing geometry agreement across sources.

    Parameters
    ----------
    geometries : list of computed field geometries
    layout : layout info from geometry engine
    wb : workbook info
    comparisons : cross-validator results (list of FieldSourceComparison)
    dpi : rendering DPI

    Returns
    -------
    ConfidenceReport with per-source and overall scores.
    """
    warnings = []
    strengths = []

    # ── Workbook confidence ──────────────────────────────────────
    wb_issues = 0
    wb_checks = 0

    # Check page dimensions are reasonable
    wb_checks += 1
    if wb.page_width_pt < 100 or wb.page_height_pt < 100:
        wb_issues += 1
        warnings.append("Page dimensions are unusually small")

    # Check margins are reasonable
    wb_checks += 1
    if wb.margin_left_pt < 0 or wb.margin_right_pt < 0 or wb.margin_top_pt < 0 or wb.margin_bottom_pt < 0:
        wb_issues += 1
        warnings.append("Negative margin values detected")

    if wb.margin_left_pt + wb.margin_right_pt >= wb.page_width_pt * 0.8:
        wb_issues += 1
        warnings.append("Margins consume most of page width")

    # Check column/row data
    wb_checks += 1
    if not wb.col_widths_pt:
        wb_issues += 1
        warnings.append("No explicit column widths in workbook XML")

    wb_checks += 1
    if wb.default_row_height <= 0:
        wb_issues += 1
        warnings.append("Default row height is zero or negative")

    # Check print area
    if wb.print_area_addr:
        wb_checks += 1
        if layout.paw_pt <= 0 or layout.pah_pt <= 0:
            wb_issues += 1
            warnings.append("Print area dimensions are zero or negative")
    else:
        wb_checks += 1
        wb_issues += 1
        warnings.append("No print area configured in workbook")

    wb_score = max(0, 100 - (wb_issues / max(wb_checks, 1)) * 100)

    if wb_score >= 80:
        strengths.append("Workbook XML provides complete geometry data")

    # ── PDF confidence ───────────────────────────────────────────
    pdf_issues = 0
    pdf_checks = 0

    if layout.eff_w_pt > 0 and layout.eff_h_pt > 0:
        pdf_checks += 1
        if layout.eff_w_pt < 10 or layout.eff_h_pt < 10:
            pdf_issues += 1
            warnings.append("PDF effective dimensions are very small")
    else:
        pdf_checks += 1
        pdf_issues += 1
        warnings.append("PDF content bounds could not be determined")

    # Check if PDF content fills the printable area reasonably
    if layout.printable_w_pt > 0 and layout.eff_w_pt > 0:
        pdf_checks += 1
        ratio = layout.eff_w_pt / layout.printable_w_pt
        if ratio < 0.1:
            pdf_issues += 1
            warnings.append(f"PDF content width ({layout.eff_w_pt:.1f}pt) is very small relative to printable area ({layout.printable_w_pt:.1f}pt)")
        elif ratio > 0.95:
            pdf_checks += 1  # bonus: content nearly fills page
            strengths.append("PDF content nearly fills printable width")

    pdf_score = max(0, 100 - (pdf_issues / max(pdf_checks, 1)) * 100)

    if pdf_score >= 80:
        strengths.append("PDF content bounds clearly defined")

    # ── PNG confidence ───────────────────────────────────────────
    png_issues = 0
    png_checks = 0

    # Cross-validation with PDF if comparisons available
    pdf_png_agreement = 0
    total_comparisons = 0

    if comparisons:
        for c in comparisons:
            if c.pdf_left_pt is not None and c.png_left_pt is not None:
                total_comparisons += 1
                diff = abs(c.pdf_left_pt - c.png_left_pt)
                if diff < 5:
                    pdf_png_agreement += 1

        if total_comparisons > 0:
            png_checks += 1
            agree_pct = pdf_png_agreement / total_comparisons
            if agree_pct < 0.5:
                png_issues += 1
                warnings.append(f"PNG/PDF content position agreement is low ({agree_pct*100:.0f}%)")
            else:
                strengths.append(f"PNG/PDF content agreement: {agree_pct*100:.0f}%")
        else:
            png_checks += 1
            png_issues += 1
            warnings.append("Could not compare PNG and PDF per-field content")

    png_score = max(0, 100 - (png_issues / max(png_checks, 1)) * 100)

    # ── Overall confidence ───────────────────────────────────────
    overall = wb_score * 0.35 + pdf_score * 0.35 + png_score * 0.30

    # Adjust down if there are many warnings
    if len(warnings) > 3:
        overall *= max(0.6, 1.0 - (len(warnings) - 3) * 0.05)

    # Adjust down if layouts have weird scaling
    if layout.scale_w < 0.5 or layout.scale_w > 2.0:
        overall *= 0.85
        warnings.append(f"Unusual horizontal scale factor: {layout.scale_w:.4f}")

    if layout.scale_h < 0.5 or layout.scale_h > 2.0:
        overall *= 0.85
        warnings.append(f"Unusual vertical scale factor: {layout.scale_h:.4f}")

    return ConfidenceReport(
        workbook_geometry=round(wb_score, 1),
        pdf_geometry=round(pdf_score, 1),
        png_geometry=round(png_score, 1),
        overall_confidence=round(min(overall, 100.0), 1),
        warnings=warnings,
        strengths=strengths,
    )

