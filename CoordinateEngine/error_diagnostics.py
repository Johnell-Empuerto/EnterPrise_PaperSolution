"""
Error diagnostics — root cause analysis for geometry misalignment.

Automatically identifies the likely cause when overlay coordinates
do not match the rendered PDF/PNG content.

Diagnostic categories:
  - Workbook geometry mismatch
  - PDF rendering mismatch
  - PNG detection mismatch
  - Merge range issue
  - Print Area issue
  - Page Setup mismatch
  - Scaling mismatch
  - Margin mismatch
  - Centering mismatch
"""

from dataclasses import dataclass
from typing import List, Optional
from CoordinateEngine.models.field_def import FieldGeometry, LayoutInfo, WorkbookInfo


@dataclass
class DiagnosticResult:
    has_issues: bool
    severity: str  # "none", "low", "medium", "high"
    categories: List[dict]  # {"category": str, "severity": str, "detail": str, "fields_affected": List[str]}
    summary: str


def diagnose(
    geometries: List[FieldGeometry],
    layout: LayoutInfo,
    wb: WorkbookInfo,
    comparisons,
    dpi: int,
) -> DiagnosticResult:
    """
    Analyze geometry and cross-validation data to identify misalignment causes.
    """
    categories = []
    all_issues = []
    severity = "none"

    # ── Check 1: Workbook geometry mismatch ──────────────────────
    if not wb.col_widths_pt:
        categories.append({
            "category": "Workbook geometry mismatch",
            "severity": "high",
            "detail": "No explicit column widths in workbook XML. Engine uses default width which may not match Excel's rendered layout.",
            "fields_affected": [g.field.addr for g in geometries],
        })
        all_issues.append("high")

    default_col_used = any(
        g.field.col not in wb.col_widths_pt for g in geometries
    )
    if default_col_used:
        categories.append({
            "category": "Workbook geometry mismatch",
            "severity": "medium",
            "detail": "Some columns use default width. Excel may render at different width than our calculation.",
            "fields_affected": [
                g.field.addr
                for g in geometries
                if g.field.col not in wb.col_widths_pt
            ],
        })
        all_issues.append("medium")

    # ── Check 2: Page Setup mismatch ─────────────────────────────
    if wb.margin_left_pt + wb.margin_right_pt >= wb.page_width_pt * 0.5:
        categories.append({
            "category": "Margin mismatch",
            "severity": "medium",
            "detail": f"Margins (L={wb.margin_left_pt:.1f} + R={wb.margin_right_pt:.1f}) consume >50% of page width ({wb.page_width_pt:.1f}pt). Content will be compressed.",
            "fields_affected": [g.field.addr for g in geometries],
        })
        all_issues.append("medium")

    if wb.margin_top_pt + wb.margin_bottom_pt >= wb.page_height_pt * 0.5:
        categories.append({
            "category": "Margin mismatch",
            "severity": "medium",
            "detail": f"Top+bottom margins consume >50% of page height.",
            "fields_affected": [g.field.addr for g in geometries],
        })
        all_issues.append("medium")

    # ── Check 3: Scaling mismatch ────────────────────────────────
    if layout.scale_w < 0.8 or layout.scale_w > 1.2:
        categories.append({
            "category": "Scaling mismatch",
            "severity": "high" if (layout.scale_w < 0.5 or layout.scale_w > 1.5) else "medium",
            "detail": f"Horizontal scale factor {layout.scale_w:.4f} deviates significantly from 1.0. Effective width ({layout.eff_w_pt:.1f}pt) differs from PA width ({layout.paw_pt:.1f}pt).",
            "fields_affected": [g.field.addr for g in geometries],
        })
        all_issues.append("high" if layout.scale_w < 0.5 else "medium")

    if layout.scale_h < 0.8 or layout.scale_h > 1.2:
        categories.append({
            "category": "Scaling mismatch",
            "severity": "high" if (layout.scale_h < 0.5 or layout.scale_h > 1.5) else "medium",
            "detail": f"Vertical scale factor {layout.scale_h:.4f} deviates significantly from 1.0.",
            "fields_affected": [g.field.addr for g in geometries],
        })
        all_issues.append("high" if layout.scale_h < 0.5 else "medium")

    # ── Check 4: Centering mismatch ──────────────────────────────
    if layout.centered_h or layout.centered_v:
        left_gap = None
        right_gap = None
        top_gap = None
        bottom_gap = None

        if comparisons and len(comparisons) > 0:
            c = comparisons[0]
            if c.pdf_left_pt is not None:
                left_gap = c.pdf_left_pt - wb.margin_left_pt
                right_gap = (wb.page_width_pt - wb.margin_right_pt) - (
                    c.pdf_left_pt + (c.pdf_w_pt or 0)
                )

        if left_gap is not None and right_gap is not None:
            gap_diff = abs(left_gap - right_gap)
            if gap_diff > 10:
                categories.append({
                    "category": "Centering mismatch",
                    "severity": "medium",
                    "detail": f"Horizontal centering detected (auto) but gap difference is {gap_diff:.1f}pt (left_gap={left_gap:.1f}, right_gap={right_gap:.1f}). PDF content is not symmetrically centered.",
                    "fields_affected": [g.field.addr for g in geometries],
                })
                all_issues.append("medium")

    # ── Check 5: Merge range issues ──────────────────────────────
    merges_with_no_content = []
    for g in geometries:
        if g.field.is_merge and not g.field.value:
            merges_with_no_content.append(g.field.addr)

    if merges_with_no_content:
        categories.append({
            "category": "Merge range issue",
            "severity": "low",
            "detail": f"Merged cells with no content: {', '.join(merges_with_no_content)}. These fields may not render anything in Excel.",
            "fields_affected": merges_with_no_content,
        })
        all_issues.append("low")

    # ── Check 6: Print Area issues ───────────────────────────────
    if not wb.print_area_addr:
        categories.append({
            "category": "Print Area issue",
            "severity": "high",
            "detail": "No print area configured in workbook. Engine uses default (first row, first column) which may be incorrect.",
            "fields_affected": [g.field.addr for g in geometries],
        })
        all_issues.append("high")

    # ── Overall assessment ──────────────────────────────────────
    if "high" in all_issues:
        severity = "high"
    elif "medium" in all_issues:
        severity = "medium"
    elif "low" in all_issues:
        severity = "low"

    high_count = all_issues.count("high")
    if high_count >= 2:
        summary = f"Multiple critical issues detected ({high_count} high severity). Alignment may be significantly off."
    elif severity == "high":
        summary = "At least one critical issue detected. Coordinate accuracy may be compromised."
    elif severity == "medium":
        summary = "Moderate issues detected. Coordinate accuracy may have minor deviations."
    elif severity == "low":
        summary = "Minor issues detected. Coordinates should be mostly accurate."
    else:
        summary = "No issues detected. Coordinate engine is operating normally."

    return DiagnosticResult(
        has_issues=severity != "none",
        severity=severity,
        categories=categories,
        summary=summary,
    )


def generate_diagnostic_report(
    diagnostic: DiagnosticResult,
    out_dir: str,
) -> str:
    """
    Generate a diagnostic_report.md file.
    """
    lines = []
    lines.append("# Diagnostic Report")
    lines.append("")
    lines.append(f"**Status:** {'Issues Found' if diagnostic.has_issues else 'Clean'}")
    lines.append(f"**Severity:** {diagnostic.severity}")
    lines.append("")
    lines.append("## Summary")
    lines.append(f"{diagnostic.summary}")
    lines.append("")
    lines.append("## Issues")
    lines.append("")

    if not diagnostic.categories:
        lines.append("No issues detected.")
    else:
        lines.append("| Category | Severity | Detail | Fields Affected |")
        lines.append("|----------|----------|--------|-----------------|")
        for cat in diagnostic.categories:
            detail = cat["detail"]
            if len(detail) > 100:
                detail = detail[:97] + "..."
            fields = ", ".join(cat["fields_affected"])
            if len(fields) > 60:
                fields = fields[:57] + "..."
            lines.append(
                f"| {cat['category']} | {cat['severity']} | {detail} | {fields} |"
            )

    lines.append("")
    lines.append("## Recommendations")
    lines.append("")

    if "Print Area issue" in [c['category'] for c in diagnostic.categories]:
        lines.append("- Set a print area in Excel (Page Layout > Print Area > Set Print Area)")
    if "Margin mismatch" in [c['category'] for c in diagnostic.categories]:
        lines.append("- Reduce page margins in Excel Page Setup")
    if "Scaling mismatch" in [c['category'] for c in diagnostic.categories]:
        lines.append("- Check FitToPages or Zoom settings in Excel Page Setup")
    if "Centering mismatch" in [c['category'] for c in diagnostic.categories]:
        lines.append("- Check Center on Page settings in Excel Page Setup")
    if "Workbook geometry mismatch" in [c['category'] for c in diagnostic.categories]:
        lines.append("- Ensure workbook has explicit column widths for all print area columns")
    if "Merge range issue" in [c['category'] for c in diagnostic.categories]:
        lines.append("- Verify merged cells contain data or ensure merge areas are correct")

    if not any(c['severity'] == 'high' for c in diagnostic.categories):
        lines.append("- No critical issues found. Coordinates should be accurate.")

    import os
    path = os.path.join(out_dir, "diagnostic_report.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"  Diagnostic report: {path}")
    return path
