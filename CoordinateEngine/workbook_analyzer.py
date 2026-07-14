import os
from typing import List
from CoordinateEngine.workbook_reader import read_workbook
from CoordinateEngine.models.field_def import WorkbookInfo, CellFormat, SheetInfo, MergeRange


def analyze_workbook(filepath: str, debug_dir: str = None) -> str:
    """Generate a comprehensive workbook_analysis.md from an XLSX file."""
    info = read_workbook(filepath)

    lines = []
    lines.append("# Workbook Analysis\n")
    lines.append(f"- **File:** {os.path.basename(filepath)}")
    lines.append(f"- **Path:** {filepath}")
    lines.append(f"- **Workbook Name:** {info.name}\n")

    # --- Sheets ---
    lines.append("## Sheets\n")
    lines.append(f"| # | Name | Visible | Page Size | Orientation | Center H | Center V |")
    lines.append(f"|---|------|---------|-----------|-------------|----------|----------|")
    for s in info.sheets:
        vis = "Yes" if s.visible else "No (hidden)"
        size = f"{s.page_width_pt:.1f} x {s.page_height_pt:.1f} pt"
        lines.append(f"| {s.index + 1} | {s.name} | {vis} | {size} | {s.orientation} | {s.center_horizontally} | {s.center_vertically} |")
    lines.append("")

    # --- Named Ranges ---
    lines.append("## Named Ranges\n")
    if info.named_ranges:
        lines.append(f"| # | Name | Refers To | Scope | Cell Ref |")
        lines.append(f"|---|------|-----------|-------|----------|")
        for i, nr in enumerate(info.named_ranges):
            lines.append(f"| {i + 1} | {nr.name} | {nr.refers_to} | {nr.scope} | {nr.cell_ref or '-'} |")
    else:
        lines.append("No named ranges defined.\n")
    lines.append("")

    # --- Dimensions ---
    lines.append("## Dimensions\n")
    lines.append(f"- **Columns found:** {len(info.col_widths_pt)}")
    lines.append(f"- **Rows found:** {len(info.row_heights_pt)}")
    lines.append(f"- **Default column width:** {info.default_col_width} pt")
    lines.append(f"- **Default row height:** {info.default_row_height} pt")
    lines.append(f"- **Page width:** {info.page_width_pt:.1f} pt")
    lines.append(f"- **Page height:** {info.page_height_pt:.1f} pt")
    lines.append(f"- **Margins:** L={info.margin_left_pt:.2f} R={info.margin_right_pt:.2f} T={info.margin_top_pt:.2f} B={info.margin_bottom_pt:.2f}")
    lines.append(f"- **Print area:** {info.print_area_addr or 'None'}")
    if info.print_titles:
        pt = info.print_titles
        lines.append(f"- **Print titles (rows):** {pt.rows_to_repeat_at_top or 'None'}")
        lines.append(f"- **Print titles (cols):** {pt.cols_to_repeat_at_left or 'None'}")
    lines.append("")

    # --- Hidden Rows and Columns ---
    lines.append("## Hidden Rows & Columns\n")
    hidden_cols = {(c, s) for c, s in info.hidden_cols}
    hidden_rows = {(r, s) for r, s in info.hidden_rows}
    if hidden_cols:
        by_sheet_hc = {}
        for c, s in hidden_cols:
            by_sheet_hc.setdefault(s, []).append(c)
        for s_idx, cols in sorted(by_sheet_hc.items()):
            sheet_name = info.sheets[s_idx].name if s_idx < len(info.sheets) else f"Sheet{s_idx + 1}"
            lines.append(f"- **{sheet_name}:** {len(cols)} hidden columns: {sorted(cols)}")
    else:
        lines.append("- No hidden columns detected.\n")
    if hidden_rows:
        by_sheet_hr = {}
        for r, s in hidden_rows:
            by_sheet_hr.setdefault(s, []).append(r)
        for s_idx, rows in sorted(by_sheet_hr.items()):
            sheet_name = info.sheets[s_idx].name if s_idx < len(info.sheets) else f"Sheet{s_idx + 1}"
            lines.append(f"- **{sheet_name}:** {len(rows)} hidden rows: {sorted(rows)}")
    else:
        lines.append("- No hidden rows detected.\n")
    lines.append("")

    # --- Merged Cells ---
    lines.append("## Merged Cells\n")
    if info.merges:
        by_sheet_merge = {}
        for (c, r, s_idx), val in info.merges.items():
            by_sheet_merge.setdefault(s_idx, []).append(val.ref)
        for s_idx, refs in sorted(by_sheet_merge.items()):
            sheet_name = info.sheets[s_idx].name if s_idx < len(info.sheets) else f"Sheet{s_idx + 1}"
            lines.append(f"- **{sheet_name}:** {len(refs)} merges")
            for ref_str in sorted(refs):
                lines.append(f"  - {ref_str}")
    else:
        lines.append("No merged cells.\n")
    lines.append("")

    # --- Data Validation ---
    lines.append("## Data Validation\n")
    if info.validation_rules:
        lines.append(f"| # | Type | Operator | Formula1 | Formula2 | Allow Blank |")
        lines.append(f"|---|------|----------|----------|----------|-------------|")
        for i, dv in enumerate(info.validation_rules):
            lines.append(f"| {i + 1} | {dv.type} | {dv.operator or '-'} | {dv.formula1 or ''} | {dv.formula2 or ''} | {dv.allow_blank} |")
    else:
        lines.append("No data validation rules.\n")
    lines.append("")

    # --- Protection ---
    lines.append("## Worksheet Protection\n")
    if info.protection and info.protection.sheet_protected:
        p = info.protection
        lines.append(f"- **Protected:** Yes (password hash: {p.sheet_password or 'none'})")
        lines.append(f"- **Select locked cells:** {p.select_locked_cells}")
        lines.append(f"- **Select unlocked cells:** {p.select_unlocked_cells}")
        lines.append(f"- **Format cells:** {p.format_cells}")
        lines.append(f"- **Format columns:** {p.format_columns}")
        lines.append(f"- **Format rows:** {p.format_rows}")
        lines.append(f"- **Insert columns:** {p.insert_columns}")
        lines.append(f"- **Insert rows:** {p.insert_rows}")
        lines.append(f"- **Delete columns:** {p.delete_columns}")
        lines.append(f"- **Delete rows:** {p.delete_rows}")
        lines.append(f"- **Sort:** {p.sort}")
        lines.append(f"- **AutoFilter:** {p.auto_filter}")
    else:
        lines.append("Worksheet is not protected.\n")
    lines.append("")

    # --- Hyperlinks ---
    lines.append("## Hyperlinks\n")
    hyperlinks = getattr(info, '_hyperlinks', {})
    if hyperlinks:
        for ref, target in hyperlinks.items():
            lines.append(f"- {ref} -> {target}")
    else:
        lines.append("No hyperlinks.\n")
    lines.append("")

    # --- Comments ---
    lines.append("## Comments & Notes\n")
    comments = getattr(info, '_comments', {})
    if comments:
        for ref, text in comments.items():
            lines.append(f"- {ref}: _{text}_")
    else:
        lines.append("No comments.\n")
    lines.append("")

    # --- Drawings (Images, Shapes, Charts) ---
    lines.append("## Drawings\n")
    if info.drawings:
        lines.append(f"| # | Type | Name | Description | Anchor | Size (pt) |")
        lines.append(f"|---|------|------|-------------|--------|-----------|")
        for i, d in enumerate(info.drawings):
            anchor = f"{d.anchor_col},{d.anchor_row}" if (d.anchor_col is not None and d.anchor_row is not None) else "-"
            size = f"{d.width_pt:.1f} x {d.height_pt:.1f}" if d.width_pt else "-"
            img_info = f" ({d.image_type})" if d.image_type else ""
            lines.append(f"| {i + 1} | {d.drawing_type}{img_info} | {d.name or '-'} | {d.description or '-'} | {anchor} | {size} |")
    else:
        lines.append("No drawings (images/shapes/charts).\n")
    lines.append("")

    # --- Controls (Form Controls / ActiveX) ---
    lines.append("## Controls\n")
    if info.controls:
        lines.append(f"| # | Type | ActiveX | Name | Cell Ref | Linked Cell |")
        lines.append(f"|---|------|---------|------|----------|-------------|")
        for i, c in enumerate(info.controls):
            ax = "Yes" if c.is_activex else "No"
            lines.append(f"| {i + 1} | {c.ctrl_type} | {ax} | {c.name or '-'} | {c.cell_ref or '-'} | {c. linked_cell or '-'} |")
    else:
        lines.append("No form controls (checkboxes, buttons, etc.).\n")
    lines.append("")

    # --- Formatting (Summary) ---
    lines.append("## Cell Formatting Summary\n")
    if info.formats:
        lines.append(f"- **Total cell formats (xf entries):** {len(info.formats)}")
        font_names = set()
        font_sizes = set()
        has_bold = 0
        has_color = 0
        for fmt in info.formats.values():
            if fmt.font:
                if fmt.font.name:
                    font_names.add(fmt.font.name)
                if fmt.font.size:
                    font_sizes.add(fmt.font.size)
                if fmt.font.bold:
                    has_bold += 1
                if fmt.font.color_argb:
                    has_color += 1
        if font_names:
            lines.append(f"- **Fonts used:** {', '.join(sorted(font_names))}")
        if font_sizes:
            lines.append(f"- **Font sizes:** {', '.join(str(s) for s in sorted(font_sizes))}")
        lines.append(f"- **Bold format count:** {has_bold}")
        lines.append(f"- **Custom color count:** {has_color}")
        alignments = set()
        for fmt in info.formats.values():
            if fmt.alignment and fmt.alignment.horizontal:
                alignments.add(fmt.alignment.horizontal)
        if alignments:
            lines.append(f"- **Alignments:** {', '.join(sorted(alignments))}")
    else:
        lines.append("No format information available.\n")
    lines.append("")

    # --- Formulas ---
    lines.append("## Formula Analysis\n")
    formula_count = 0
    for sheet_data_entry in _collect_sheet_cells(info):
        if sheet_data_entry.get('formula'):
            formula_count += 1
    lines.append(f"- **Total formulas detected:** {formula_count}")
    lines.append("(Full per-cell formula details available in debug overlay.)\n")
    lines.append("")

    # --- Printer Settings ---
    lines.append("## Printer Settings\n")
    if info.printer_settings_path:
        lines.append(f"- **Printer settings file:** {info.printer_settings_path}")
    lines.append(f"- **Paper size:** {info.sheets[0].paper_size if info.sheets else 'Letter'}")
    lines.append(f"- **Orientation:** {info.sheets[0].orientation if info.sheets else 'portrait'}")
    lines.append("")

    # --- Summary Stats ---
    lines.append("## Summary Statistics\n")
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| Total sheets | {len(info.sheets)} |")
    lines.append(f"| Visible sheets | {sum(1 for s in info.sheets if s.visible)} |")
    lines.append(f"| Named ranges | {len(info.named_ranges)} |")
    lines.append(f"| Merged cell regions | {len(info.merges)} |")
    lines.append(f"| Hidden columns | {len(info.hidden_cols)} |")
    lines.append(f"| Hidden rows | {len(info.hidden_rows)} |")
    lines.append(f"| Data validations | {len(info.validation_rules)} |")
    lines.append(f"| Drawings (images/shapes) | {len(info.drawings)} |")
    lines.append(f"| Form controls | {len(info.controls)} |")
    lines.append(f"| Hyperlinks | {len(getattr(info, '_hyperlinks', {}))} |")
    lines.append(f"| Comments | {len(getattr(info, '_comments', {}))} |")
    lines.append(f"| Cell formats | {len(info.formats)} |")
    lines.append(f"| Protected | {info.protection is not None and info.protection.sheet_protected} |")
    lines.append("")

    analysis = '\n'.join(lines)

    if debug_dir:
        os.makedirs(debug_dir, exist_ok=True)
        out_path = os.path.join(debug_dir, 'workbook_analysis.md')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(analysis)
        return out_path
    return analysis


def _collect_sheet_cells(info: WorkbookInfo) -> list:
    """Helper to iterate cells; returns list of dicts. Currently returns empty until full cell map is parsed."""
    return []


def run_analysis(excel_path: str, debug_dir: str = None) -> str:
    """CLI entry point for workbook analysis."""
    return analyze_workbook(excel_path, debug_dir)


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print("Usage: python -m CoordinateEngine.workbook_analyzer <path_to.xlsx>")
        sys.exit(1)
    result = run_analysis(path, debug_dir='.')
    print(result)