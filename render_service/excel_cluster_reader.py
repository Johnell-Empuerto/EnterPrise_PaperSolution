"""
ExcelClusterReader — reads cluster metadata from the hidden '_Fields' sheet
inside a ConMas-generated workbook and produces FieldDefinition objects
identical to those returned from the database.

Storage format (confirmed by investigation):
  - Hidden worksheet named '_Fields'
  - Columns: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes
  - Each row = one form field

Fallback: if _Fields sheet is missing/empty, reads cell comments.

Ratio computation uses the EXISTING proven coordinate pipeline:
  compute_page_layout() + cell_range_to_page_rect()
to match the database-stored ratios exactly.

Output: List[FieldDef] — matching the database schema identically.
"""

import os
import sys
from pathlib import Path
from typing import Optional

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

from CoordinateEngine.models.field_def import FieldDef
from CoordinateEngine.workbook_reader import read_workbook
from render_service.page_coordinate_transformer import compute_page_layout, cell_range_to_page_rect
from CoordinateEngine.utils.file_utils import parse_range


# ── column indices in the _Fields sheet ──────────────────────────────────
COL_ADDRESS    = 1   # A
COL_FIELD_ID   = 2   # B
COL_FIELD_NAME = 3   # C
COL_FIELD_TYPE = 4   # D
COL_SHEET_NAME = 5   # E
COL_CREATED    = 6   # F
COL_NOTES      = 7   # G


# ── known ConMas field type identifiers ───────────────────────────────────
KNOWN_FIELD_TYPES = frozenset({
    "KeyboardText",
    "Handwriting",
    "Machine",
    "Machine_Output",
    "InputNumeric",
    "Number",
    "Integer",
    "Date",
    "DateTime",
    "CheckBox",
    "RadioButton",
    "DropDown",
    "ComboBox",
    "ListBox",
    "Label",
    "Email",
    "Phone",
    "Signature",
    "Image",
    "Barcode",
    "TextArea",
    "Password",
    "ZipCode",
    "Memo",
    "QRCode",
})


def read_fields(xlsx_path: str) -> list[FieldDef]:
    """Read cluster/field definitions from the workbook.

    Priority:
      1. Hidden '_Fields' sheet (primary — created by ConMas Designer)
      2. Cell comments (fallback — if _Fields does not exist or is empty)

    Uses the proven coordinate pipeline for ratio computation so the
    resulting FieldDef objects match the database schema exactly.

    Phase X.38: When the global RenderQueue has a persistent Excel instance,
    this function uses it (no new Excel created).  Otherwise falls back to
    creating its own Excel.Application (original behavior).

    Returns:
        List[FieldDef] compatible with the existing renderer.
    """
    import pythoncom
    import win32com.client
    from render_service.render_queue import get_queue_excel

    # ── Phase X.38: Use queue's persistent Excel if available ──────
    queue_excel = get_queue_excel()
    if queue_excel is not None:
        wb = queue_excel.Workbooks.Open(os.path.abspath(xlsx_path))
        try:
            fields = _read_fields_sheet(wb, xlsx_path)
            if fields:
                return fields
            fields = _read_comments(wb, xlsx_path)
            if fields:
                return fields
            return []
        finally:
            try:
                wb.Close(False)
            except Exception:
                pass

    # ── Original behavior: create own Excel ────────────────────────
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
            fields = _read_fields_sheet(wb, xlsx_path)
            if fields:
                return fields
            fields = _read_comments(wb, xlsx_path)
            if fields:
                return fields
            return []
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()


def _read_fields_sheet(wb_com, xlsx_path: str) -> list[FieldDef]:
    """Read cluster definitions from hidden _Fields worksheet.

    Uses Excel COM to read the sheet, then feeds cell addresses through
    the proven coordinate pipeline for ratio computation.
    """
    # Locate _Fields sheet
    sheet = None
    for i in range(1, wb_com.Sheets.Count + 1):
        if wb_com.Sheets(i).Name == "_Fields":
            sheet = wb_com.Sheets(i)
            break
    if sheet is None:
        return []

    last_row = sheet.UsedRange.Rows.Count
    if last_row < 2:
        return []

    # Read the workbook via our proven parser (for coordinate pipeline)
    wb_parsed = read_workbook(xlsx_path, verbose=False)
    sheet_index = 0

    # Determine which visible sheet to use as the data sheet
    data_sheet_name = None
    data_sheet_index = 0
    data_sheet_com = None  # Excel COM worksheet object
    for i in range(1, wb_com.Sheets.Count + 1):
        ws = wb_com.Sheets(i)
        if ws.Visible == -1 and ws.Name != "_Fields":
            data_sheet_name = ws.Name
            data_sheet_com = ws
            # Find matching index in parsed workbook
            for j, s in enumerate(wb_parsed.sheets):
                if s.name == data_sheet_name:
                    data_sheet_index = j
                    break
            break

    # Compute layout using the proven pipeline
    layout = compute_page_layout(wb_parsed, sheet_index=data_sheet_index, dpi=300)
    page_w_pt = layout.page_width_pt
    page_h_pt = layout.page_height_pt

    fields = []
    for row in range(2, last_row + 1):
        cell_addr = _get_cell_str(sheet, row, COL_ADDRESS)
        field_id  = _get_cell_str(sheet, row, COL_FIELD_ID)
        field_name = _get_cell_str(sheet, row, COL_FIELD_NAME)
        field_type = _get_cell_str(sheet, row, COL_FIELD_TYPE)
        sheet_name = _get_cell_str(sheet, row, COL_SHEET_NAME)

        if not cell_addr:
            continue

        # Parse the cell address
        parsed = parse_range(cell_addr)
        if parsed is None:
            # Try to read range address from Excel COM
            try:
                if data_sheet_com is not None:
                    rng = data_sheet_com.Range(cell_addr)
                    if rng.MergeCells:
                        ma = rng.MergeArea
                        cell_addr = ma.Address(False, False)
                    parsed = parse_range(cell_addr)
            except Exception:
                pass
            if parsed is None:
                continue

        col, r, col_end, row_end = parsed

        # ── Use Excel COM Range dimensions (as user specified) ───────
        # Read actual rendered cell position from Excel COM
        try:
            if data_sheet_com is not None:
                rng = data_sheet_com.Range(cell_addr)
                range_left_pt   = float(rng.Left)
                range_top_pt    = float(rng.Top)
                range_width_pt  = float(rng.Width)
                range_height_pt = float(rng.Height)
            else:
                raise Exception("No data sheet")
        except Exception:
            # Fallback: use coordinate pipeline from column widths
            rect = cell_range_to_page_rect(layout, wb_parsed,
                                            col, r, col_end, row_end,
                                            sheet_index=data_sheet_index)
            range_left_pt   = rect["page_left_pt"]
            range_top_pt    = rect["page_top_pt"]
            range_width_pt  = rect["page_width_pt"]
            range_height_pt = rect["page_height_pt"]
            # Skip layout origin below since we already have page coords
            page_left_pt   = range_left_pt
            page_top_pt    = range_top_pt
            page_width_pt  = range_width_pt
            page_height_pt = range_height_pt
        else:
            # Convert COM range positions to page coordinates via layout
            # Range.Left/Top are from worksheet origin; layout has margins/centering
            page_left_pt   = layout.origin_x_pt + range_left_pt * layout.scale_w
            page_top_pt    = layout.origin_y_pt + range_top_pt * layout.scale_h
            page_width_pt  = range_width_pt  * layout.scale_w
            page_height_pt = range_height_pt * layout.scale_h

        # ── Convert page coordinates to ratios ───────────────────────
        ratio_left   = page_left_pt   / page_w_pt
        ratio_top    = page_top_pt    / page_h_pt
        ratio_right  = (page_left_pt + page_width_pt)  / page_w_pt
        ratio_bottom = (page_top_pt  + page_height_pt) / page_h_pt

        # Map the ConMas field type to our data type
        data_type = _conmas_type_to_renderer_type(field_type)

        fd = FieldDef(
            addr=cell_addr,
            col=col,
            row=r,
            col_end=col_end,
            row_end=row_end,
            sheet_index=data_sheet_index,
            value=None,
            is_merge=(col != col_end or r != row_end),
            data_type=data_type,
            required=False,
            ratio_left=ratio_left,
            ratio_top=ratio_top,
            ratio_right=ratio_right,
            ratio_bottom=ratio_bottom,
        )
        fields.append(fd)

    return fields


def _parse_comment_text(comment_text: str) -> tuple:
    """Parse a ConMas cell comment into (field_name, field_type, parameters).

    ConMas comments store field metadata in two possible formats:

    Old format (field type first):
        KeyboardText
        0
        0

    New format (field name first, field type second):
        samples
        KeyboardText
        0
        0

    The parser detects the format by checking whether the first line
    is a recognised field type identifier.

    Returns:
        (field_name: str, field_type: str, parameters: list[str])
    """
    # Normalise line endings and split into non-empty lines
    lines = comment_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = [l.strip() for l in lines if l.strip()]

    if not lines:
        return ("", "text", [])

    first = lines[0]

    # Determine format:
    #   If first line is a known field type → old format (type first, no name)
    #   Otherwise → new format (name first, type second)
    if first in KNOWN_FIELD_TYPES:
        # Old format: first line IS the field type, generate a default name
        field_type = first
        field_name = ""  # caller must generate a default name
        parameters = lines[1:]
    else:
        # New format: first line is the field name, second is the field type
        field_name = first
        field_type = lines[1] if len(lines) > 1 else "text"
        parameters = lines[2:]

    return (field_name, field_type, parameters)


def _generate_default_field_name(sheet_index: int, row: int, col: int) -> str:
    """Generate a default field name from sheet and cell position.

    The ConMas Designer auto-generates names like "S1C1" (Sheet 1, Cell 1)
    when the comment only contains the field type (old format).

    We reuse the cell address (e.g. "A1", "B3") since that is the most
    natural positional identifier used throughout this project.

    Example: field at Sheet 0, Row 1, Col 1 → "Field_A1"
    """
    return f"Field_{_col_to_letter(col)}{row}"


def _col_to_letter(col: int) -> str:
    """Convert a 1-based column number to an Excel column letter.

    Examples:
        1  → "A"
        2  → "B"
        27 → "AA"
        28 → "AB"
    """
    letter = ""
    c = col
    while c > 0:
        c -= 1
        letter = chr(ord("A") + (c % 26)) + letter
        c //= 26
    return letter


def _read_comments(wb_com, xlsx_path: str) -> list[FieldDef]:
    """Fallback: read cluster definitions from cell comments.

    Uses per-cell Range().Comment scanning instead of the ws.Comments
    collection, because win32com's Comments collection iteration is
    unreliable with some Excel COM versions.

    This matches the ConMas Designer approach of scanning each cell
    in the used range for comments (via Infragistics HasComment()).
    """
    # Read the workbook via our proven parser
    wb_parsed = read_workbook(xlsx_path, verbose=False)
    layout = compute_page_layout(wb_parsed, sheet_index=0, dpi=300)
    page_w_pt = layout.page_width_pt
    page_h_pt = layout.page_height_pt

    fields = []

    for i in range(1, wb_com.Sheets.Count + 1):
        ws = wb_com.Sheets(i)
        if ws.Visible != -1:  # skip hidden sheets
            continue
        if ws.Name == "_Fields":
            continue

        # Quick check: does this sheet have any comments at all?
        try:
            comment_count = ws.Comments.Count
        except Exception:
            comment_count = 0

        if comment_count == 0:
            continue

        sheet_index = i - 1
        # Update layout for this sheet
        try:
            layout = compute_page_layout(wb_parsed, sheet_index=sheet_index, dpi=300)
            page_w_pt = layout.page_width_pt
            page_h_pt = layout.page_height_pt
        except Exception:
            pass

        # ── Per-cell comment scan ─────────────────────────────────────
        # The ws.Comments collection Count works, but iterating/accessing
        # individual comments often fails with win32com.
        # Instead we scan every cell in the used range using Range().Comment
        # which reliably returns the Comment COM object when present.
        #
        # Note: use absolute row/col (UsedRange.Row, UsedRange.Column)
        # not 1-based offsets, because the UsedRange may not start at A1.
        try:
            used = ws.UsedRange
            first_row = used.Row
            first_col = used.Column
            row_count = used.Rows.Count
            col_count = used.Columns.Count
        except Exception:
            continue

        for row_offset in range(row_count):
            for col_offset in range(col_count):
                r = first_row + row_offset
                c = first_col + col_offset
                # Build cell address like "A1", "C3", "AA12"
                addr_str = f"{_col_to_letter(c)}{r}"

                try:
                    cell_range = ws.Range(addr_str)
                    comment = cell_range.Comment
                except Exception:
                    comment = None

                if comment is None:
                    continue

                # Check for merged cells
                try:
                    if cell_range.MergeCells:
                        merge_area = cell_range.MergeArea
                        addr_str = merge_area.Address(False, False)
                except Exception:
                    pass

                parsed = parse_range(addr_str)
                if parsed is None:
                    continue

                col_idx, row_idx, col_end, row_end = parsed

                rect = cell_range_to_page_rect(layout, wb_parsed,
                                               col_idx, row_idx, col_end, row_end,
                                               sheet_index=sheet_index)

                page_left_pt   = rect["page_left_pt"]
                page_top_pt    = rect["page_top_pt"]
                page_width_pt  = rect["page_width_pt"]
                page_height_pt = rect["page_height_pt"]

                ratio_left   = page_left_pt   / page_w_pt
                ratio_top    = page_top_pt    / page_h_pt
                ratio_right  = (page_left_pt + page_width_pt)  / page_w_pt
                ratio_bottom = (page_top_pt  + page_height_pt) / page_h_pt

                comment_text = ""
                try:
                    comment_text = str(comment.Text())
                except Exception:
                    pass

                # ── Parse comment text using the two-format parser ──
                field_name, field_type_str, _params = _parse_comment_text(comment_text)

                data_type = _conmas_type_to_renderer_type(field_type_str)

                # Generate default field name if the comment didn't provide one
                if not field_name:
                    field_name = _generate_default_field_name(sheet_index, row_idx, col_idx)

                fd = FieldDef(
                    addr=addr_str,
                    col=col_idx,
                    row=row_idx,
                    col_end=col_end,
                    row_end=row_end,
                    sheet_index=sheet_index,
                    value=None,
                    is_merge=(col_idx != col_end or row_idx != row_end),
                    data_type=data_type,
                    required=False,
                    ratio_left=ratio_left,
                    ratio_top=ratio_top,
                    ratio_right=ratio_right,
                    ratio_bottom=ratio_bottom,
                )
                fields.append(fd)

    return fields


def _get_cell_str(sheet, row: int, col: int) -> str:
    """Read a cell value as a string from a COM worksheet."""
    try:
        val = sheet.Cells(row, col).Value
        if val is None:
            return ""
        return str(val).strip()
    except Exception:
        return ""


def _conmas_type_to_renderer_type(conmas_type: str) -> str:
    """Map ConMas field type strings to renderer data types."""
    mapping = {
        "KeyboardText":   "text",
        "Handwriting":    "signature",
        "Text":           "text",
        "Machine":        "text",
        "Machine_Output": "text",
        "InputNumeric":   "number",
        "Number":         "number",
        "Integer":        "number",
        "Date":           "date",
        "DateTime":       "date",
        "CheckBox":       "checkbox",
        "RadioButton":    "radio",
        "DropDown":       "dropdown",
        "ComboBox":       "dropdown",
        "ListBox":        "list",
        "Label":          "label",
        "Email":          "email",
        "Phone":          "phone",
        "ZipCode":        "zip",
        "Password":       "password",
        "Memo":           "textarea",
        "TextArea":       "textarea",
        "Signature":      "signature",
        "Image":          "image",
        "Barcode":        "barcode",
        "QRCode":         "qrcode",
    }
    return mapping.get(conmas_type, "text")
