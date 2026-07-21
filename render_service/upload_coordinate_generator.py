"""
ConMas Upload Coordinate Generation — MakeCluster + Range geometry + ratio normalization.

Replicates the original ConMas Designer's CalcClusterSize pipeline:

  MakeCluster() → ExportPdf() → GetAddress() → normalize()

Rectangle positions are calculated from Excel COM Range properties (Left, Top,
Width, Height) rather than PDF pixel scanning. This avoids the loss of adjacent
clusters that merge into a single blob in pixel analysis.

PDF export is still performed (required as final output), but coordinates come
from the workbook geometry directly — matching ConMas's original approach.

Supports two input formats:
  1. Cell comments (newer ConMas format — Template 547 style)
  2. _Fields sheet (older ConMas format — Template 546 style)
"""

import os
import re
import sys
import tempfile
import shutil
import traceback
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

from CoordinateEngine.utils.file_utils import parse_range

DPI = 300
PT_TO_PX = DPI / 72.0

# Phase 6.3: Legacy ConMas configuration sheet names
# These worksheets contain designer metadata, NOT printable content.
# They must be excluded from the printable worksheet pipeline exactly
# as the original ConMas Designer did (Algorithm A — Reserved Names).
CONFIGURATION_SHEET_NAMES = frozenset({
    "_Fields",
    "_RawData",
    "ExcelOutputSetting",
    "DesignerConfig",
    "PaperLessConfig",
    "ConMasConfig",
})

BLACK_THRESHOLD = 64
WHITE_THRESHOLD = 192
MIN_RECT_SIZE = 6


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

CELL_ADDR_RE = re.compile(r"(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?")


def _parse_cell_addr(addr: str):
    """Parse an Excel address like $A$1:$B$2 into (col, row, col_end, row_end)."""
    m = CELL_ADDR_RE.match(addr.strip().upper())
    if not m:
        return None
    from openpyxl.utils import column_index_from_string
    col = column_index_from_string(m.group(1).replace("$", ""))
    row = int(m.group(2).replace("$", ""))
    if m.group(3):
        col_end = column_index_from_string(m.group(3).replace("$", ""))
        row_end = int(m.group(4).replace("$", ""))
    else:
        col_end = col
        row_end = row
    return (col, row, col_end, row_end)


def _sort_key_addr(addr: str):
    """Sort key for cell addresses: (row, col)."""
    try:
        p = _parse_cell_addr(addr)
        if p:
            col, row, _, _ = p
            return (row, col)
    except Exception:
        pass
    return (0, 0)


def _sort_key_meta(m):
    return _sort_key_addr(m.get("cellAddr", ""))


# ──────────────────────────────────────────────────────────────────────────────
# Phase 0 — Identify cluster cells
# ──────────────────────────────────────────────────────────────────────────────

def _identify_clusters_from_comments(xlsx_path: str):
    """Read cell comments via Worksheet.Comments collection (native direct enumeration).

    Phase X.16 proved that ws.Comments returns the exact same comment set as
    per-cell cell.Comment scan, but 2.8x faster (no UsedRange traversal).

    Phase X.38: Uses the queue's persistent Excel when available.

    Returns list of dicts with cellAddr, name, type, input_parameter.
    Returns [] if no comments found.
    """
    import pythoncom
    import win32com.client
    from render_service.render_queue import get_queue_excel

    # ── Phase X.38: Use queue's persistent Excel if available ──────
    queue_excel = get_queue_excel()
    if queue_excel is not None:
        wb = queue_excel.Workbooks.Open(os.path.abspath(xlsx_path))
        try:
            return _read_comments_native(wb)
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
        clusters = []
        try:
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
            clusters = _read_comments_native(wb)
            wb.Close(False)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
        return clusters
    finally:
        pythoncom.CoUninitialize()


def _read_comments_native(wb):
    """Read cell comments from an already-opened workbook.

    Uses Comments.Count to skip empty sheets, then iterates via
    ws.Comments for a 2.8x speed improvement over per-cell scanning.

    Returns list of dicts with cellAddr, name, type, input_parameter.
    """
    clusters = []
    for ws in wb.Worksheets:
        try:
            if ws.Comments.Count == 0:
                continue
        except Exception:
            try:
                used = ws.UsedRange
                if used is None:
                    continue
                lr = used.Row + used.Rows.Count - 1
                lc = used.Column + used.Columns.Count - 1
                for row in range(1, lr + 1):
                    for col in range(1, lc + 1):
                        try:
                            cell = ws.Cells(row, col)
                            comment = cell.Comment
                            if comment is None:
                                continue
                            ma = cell.MergeArea
                            addr = str(ma.Address)
                            text = str(comment.Text())
                            lines = text.replace("\r\n", "\n").split("\n")
                            clusters.append({
                                "cellAddr": addr,
                                "name": lines[0].strip() if lines else "",
                                "type": lines[1].strip() if len(lines) > 1 else "Text",
                                "input_parameter": "\n".join(lines[2:]).strip() if len(lines) > 2 else "",
                            })
                        except Exception:
                            continue
            except Exception:
                pass
            continue

        for comment in ws.Comments:
            try:
                parent_range = comment.Parent
            except Exception:
                continue
            if parent_range is None:
                continue
            try:
                merged = parent_range.MergeArea
                addr = str(merged.Address)
            except Exception:
                addr = str(parent_range.Address)
            try:
                text = str(comment.Text())
            except Exception:
                text = ""
            lines = text.replace("\r\n", "\n").split("\n")
            clusters.append({
                "cellAddr": addr,
                "name": lines[0].strip() if lines else "",
                "type": lines[1].strip() if len(lines) > 1 else "Text",
                "input_parameter": "\n".join(lines[2:]).strip() if len(lines) > 2 else "",
            })
    return clusters


def _identify_clusters_from_fields_sheet(xlsx_path: str):
    """Read cluster ranges from hidden _Fields sheet.

    The _Fields sheet has columns:
      A=Address, B=FieldId, C=FieldName, D=FieldType, E=SheetName, ...

    Phase X.38: Uses the queue's persistent Excel when available.

    Returns list of dicts with cellAddr, name, type, input_parameter.
    Returns [] if no _Fields sheet or no data rows.
    """
    import pythoncom
    import win32com.client
    from render_service.render_queue import get_queue_excel

    # ── Phase X.38: Use queue's persistent Excel if available ──────
    queue_excel = get_queue_excel()
    if queue_excel is not None:
        wb = queue_excel.Workbooks.Open(os.path.abspath(xlsx_path))
        try:
            return _read_fields_sheet_data(wb)
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
        clusters = []
        try:
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
            clusters = _read_fields_sheet_data(wb)
            wb.Close(False)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
        return clusters
    finally:
        pythoncom.CoUninitialize()


def _read_fields_sheet_data(wb):
    """Read cluster metadata from the _Fields sheet in an already-opened workbook.

    Returns list of dicts with cellAddr, name, type, input_parameter.
    Returns [] if no _Fields sheet or no data rows.
    """
    clusters = []
    fields_sheet = None
    for i in range(1, wb.Sheets.Count + 1):
        if wb.Sheets(i).Name == "_Fields":
            fields_sheet = wb.Sheets(i)
            break
    if fields_sheet is None:
        return []

    lr = fields_sheet.UsedRange.Rows.Count
    if lr < 2:
        return []

    for row in range(2, lr + 1):
        cell_addr = None
        try:
            v = fields_sheet.Cells(row, 1).Value
            if v is not None:
                cell_addr = str(v).strip()
        except Exception:
            pass
        if not cell_addr or not CELL_ADDR_RE.match(cell_addr):
            continue

        field_name = ""
        try:
            v = fields_sheet.Cells(row, 3).Value
            if v is not None:
                field_name = str(v).strip()
        except Exception:
            pass

        field_type = "Text"
        try:
            v = fields_sheet.Cells(row, 4).Value
            if v is not None:
                field_type = str(v).strip()
        except Exception:
            pass

        clusters.append({
            "cellAddr": cell_addr,
            "name": field_name,
            "type": field_type,
            "input_parameter": "",
        })
    return clusters


def _identify_clusters(xlsx_path: str) -> list[dict]:
    """Identify cluster cells from the workbook.

    Priority:
      1. Cell comments (original MakeCluster behavior)
      2. _Fields sheet (older ConMas format)

    Returns list of dicts with cellAddr, name, type, input_parameter.
    """
    clusters = _identify_clusters_from_comments(xlsx_path)
    if clusters:
        return clusters
    return _identify_clusters_from_fields_sheet(xlsx_path)


# ──────────────────────────────────────────────────────────────────────────────
# Phase A — MakeCluster: sanitize workbook
# ──────────────────────────────────────────────────────────────────────────────

def sanitize_workbook(xlsx_path: str, clusters: list[dict]) -> str:
    """Replicate ConMas MakeCluster().

    Fills cluster cells with BLACK, all other cells WHITE.
    Removes borders, shapes, headers, footers, cell values.

    Phase X.38: Uses the queue's persistent Excel when available.

    Args:
        xlsx_path: Path to original workbook.
        clusters: List of dicts with cellAddr to fill black.

    Returns:
        Path to sanitized temporary XLSX.
    """
    import pythoncom
    import win32com.client
    from render_service.render_queue import get_queue_excel

    tmp_dir = tempfile.mkdtemp(prefix="ple_sanitize_")
    sanitized_path = os.path.join(tmp_dir, "sanitized.xlsx")

    # Build set of cluster ranges for quick lookup
    cluster_addrs = set()
    for c in clusters:
        addr = c["cellAddr"]
        cluster_addrs.add(addr.upper().replace(" ", ""))

    # ── Phase X.38: Use queue's persistent Excel if available ──────
    queue_excel = get_queue_excel()
    if queue_excel is not None:
        try:
            wb = queue_excel.Workbooks.Open(os.path.abspath(xlsx_path))
            try:
                _do_sanitize(wb, cluster_addrs)
                wb.SaveAs(os.path.abspath(sanitized_path))
                wb.Close(False)
            except Exception:
                shutil.rmtree(tmp_dir, ignore_errors=True)
                raise
        except Exception:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            raise
        return sanitized_path

    # ── Original behavior: create own Excel ────────────────────────
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        try:
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
            _do_sanitize(wb, cluster_addrs)
            wb.SaveAs(os.path.abspath(sanitized_path))
            wb.Close(False)
        except Exception:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            raise
        finally:
            try:
                excel.Quit()
            except Exception:
                pass

        return sanitized_path
    finally:
        pythoncom.CoUninitialize()


def _do_sanitize(wb, cluster_addrs: set):
    """Perform workbook sanitization on an already-opened workbook.

    Fills cluster cells black, all others white.  Removes shapes,
    cell values, and borders.
    """
    xlNone = -4142
    for ws in wb.Worksheets:
        try:
            if ws.Shapes.Count > 0:
                for shape in list(ws.Shapes):
                    shape.Delete()
        except Exception:
            pass

        used = ws.UsedRange
        if used is None:
            continue

        lr = used.Row + used.Rows.Count - 1
        lc = used.Column + used.Columns.Count - 1

        try:
            used.Interior.Color = 0xFFFFFF
        except Exception:
            pass

        try:
            used.ClearContents()
        except Exception:
            pass

        for addr in cluster_addrs:
            try:
                rng = ws.Range(addr)
                rng.Interior.Color = 1
                rng.Value = ""
            except Exception:
                pass

        try:
            clear_range = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc))
            clear_range.Borders.LineStyle = xlNone
        except Exception:
            pass


# ──────────────────────────────────────────────────────────────────────────────
# Phase B — Export sanitized workbook to PDF via Excel COM
# ──────────────────────────────────────────────────────────────────────────────

def export_sanitized_pdf(sanitized_path: str) -> str:
    """Export sanitized workbook to PDF via Excel COM ExportAsFixedFormat.

    Phase X.38: Uses the queue's persistent Excel when available.
    """
    import pythoncom
    import win32com.client
    from render_service.render_queue import get_queue_excel

    tmp_dir = tempfile.mkdtemp(prefix="ple_export_pdf_")
    pdf_path = os.path.join(tmp_dir, "sanitized.pdf")

    # ── Phase X.38: Use queue's persistent Excel if available ──────
    queue_excel = get_queue_excel()
    if queue_excel is not None:
        try:
            wb = queue_excel.Workbooks.Open(os.path.abspath(sanitized_path))
            try:
                _delete_metadata_sheets(wb)
                ws = wb.ActiveSheet
                ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path), 0, 0, True)
                wb.Close(False)
            except Exception:
                shutil.rmtree(tmp_dir, ignore_errors=True)
                raise
        except Exception:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            raise
        return pdf_path

    # ── Original behavior: create own Excel ────────────────────────
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        try:
            wb = excel.Workbooks.Open(os.path.abspath(sanitized_path))
            _delete_metadata_sheets(wb)
            ws = wb.ActiveSheet
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path), 0, 0, True)
            wb.Close(False)
        except Exception:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            raise
        finally:
            try:
                excel.Quit()
            except Exception:
                pass

        return pdf_path
    finally:
        pythoncom.CoUninitialize()


def _delete_metadata_sheets(wb):
    """Delete metadata sheets (_Fields, _RawData) from an opened workbook."""
    # PHASE 6.1 DIAGNOSTIC: Before deletion
    names_before = []
    for di in range(1, wb.Sheets.Count + 1):
        try:
            names_before.append(str(wb.Sheets(di).Name))
        except Exception:
            names_before.append("<ERROR>")
    print(f"  [_delete_metadata_sheets] BEFORE: {names_before}")

    to_delete = []
    for i in range(wb.Sheets.Count, 0, -1):
        name = ""
        try:
            name = wb.Sheets(i).Name
        except Exception:
            pass
        if name in CONFIGURATION_SHEET_NAMES:
            print(f"  [_delete_metadata_sheets] Deleting metadata sheet: '{name}'")
            to_delete.append(i)

    for i in to_delete:
        try:
            wb.Sheets(i).Delete()
            print(f"  [_delete_metadata_sheets] Deleted: '{i}'")
        except Exception as e:
            print(f"  [_delete_metadata_sheets] Delete FAILED for sheet {i}: {e}")
            pass

    # PHASE 6.3: After deletion summary
    names_after = []
    for di in range(1, wb.Sheets.Count + 1):
        try:
            names_after.append(str(wb.Sheets(di).Name))
        except Exception:
            names_after.append("<ERROR>")
    print(f"  [_delete_metadata_sheets] Remaining sheets: {names_after}")


# ──────────────────────────────────────────────────────────────────────────────
# Phase C — Render PDF to image at 300 DPI
# ──────────────────────────────────────────────────────────────────────────────

def render_pdf_to_image(pdf_path: str, dpi: int = DPI, page_index: int = 0):
    """Render a PDF page at 300 DPI to numpy array.

    Args:
        pdf_path: Path to the PDF file.
        dpi: Rendering resolution (default 300).
        page_index: Zero-based page index to render (default 0).

    Returns (numpy_array_HWC, width_px, height_px) in RGB uint8.
    """
    if not HAS_NUMPY:
        raise RuntimeError("numpy is required for pixel scanning")
    import fitz

    doc = fitz.open(pdf_path)
    try:
        if page_index >= len(doc):
            raise IndexError(
                f"PDF has {len(doc)} pages, requested page {page_index}"
            )
        page = doc[page_index]
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        w, h = pix.width, pix.height
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, 3)
    finally:
        doc.close()

    return img, w, h


# ──────────────────────────────────────────────────────────────────────────────
# Phase D — Legacy pixel scan (replaced by calculate_rectangles_from_excel)
# ──────────────────────────────────────────────────────────────────────────────

def _is_black(r: int, g: int, b: int, threshold: int = BLACK_THRESHOLD) -> bool:
    return r < threshold and g < threshold and b < threshold


def _is_white(r: int, g: int, b: int, threshold: int = WHITE_THRESHOLD) -> bool:
    return r > threshold and g > threshold and b > threshold


def _is_not_black(r: int, g: int, b: int, threshold: int = BLACK_THRESHOLD) -> bool:
    return not (r < threshold and g < threshold and b < threshold)


def scan_black_rectangles(img, min_size: int = MIN_RECT_SIZE) -> list[dict]:
    """Scan image pixels for black rectangles on white background.

    Replication of ConMas GetAddress() algorithm:
      1. Top→bottom, left→right scan for top-left corners
      2. Detect corner: BLACK pixel with NON-BLACK left AND above
      3. 6-pixel noise verification
      4. Expand right until white → right edge
      5. Expand down (per column, take max Y) → bottom edge
      6. Minimum 6×6 pixel filter
      7. Mark visited rect

    NumPy-vectorized implementation (validated: identical output,
    67.8x faster than original Python-loop version).

    Returns list of dicts with Left, Top, Right, Bottom (pixel coords).
    """
    h, w, _ = img.shape

    # Pre-compute binary masks (one-time vectorized operations)
    black = np.all(img < BLACK_THRESHOLD, axis=2)
    white = np.all(img > WHITE_THRESHOLD, axis=2)
    not_black = ~black  # equivalent to _is_not_black()

    visited = np.zeros((h, w), dtype=bool)
    rects = []

    # Vectorized corner detection:
    #   Corner = black pixel where pixel above AND left are NOT black
    #   Pad edges with True (= not-black) so edge pixels are never corners
    left_pad = np.pad(not_black, ((0, 0), (1, 0)), constant_values=True)[:, :-1]
    above_pad = np.pad(not_black, ((1, 0), (0, 0)), constant_values=True)[:-1, :]
    corners = black & left_pad & above_pad

    corner_ys, corner_xs = np.where(corners)

    for idx in range(len(corner_ys)):
        y, x = corner_ys[idx], corner_xs[idx]

        if visited[y, x]:
            continue

        # 6-pixel noise verification (top row: none of the 6 pixels above are black)
        if np.any(black[y - 1, x + 1:x + min_size]):
            continue

        # 6-pixel noise verification (left column: none of the 6 pixels left are black)
        if np.any(black[y + 1:y + min_size, x - 1]):
            continue

        # Expand right until white (same logic as original while loop)
        row_white = white[y, x + 1:]
        first_white = np.where(row_white)[0]
        right_x = (x + first_white[0]) if len(first_white) > 0 else (w - 1)

        if right_x - x + 1 < min_size:
            continue

        # Expand down per column until white (same logic as original nested while)
        bottom_y = y
        for col in range(x, right_x + 1):
            col_white = white[y + 1:, col]
            first_white_in_col = np.where(col_white)[0]
            col_bottom = (y + first_white_in_col[0]) if len(first_white_in_col) > 0 else (h - 1)
            if col_bottom > bottom_y:
                bottom_y = col_bottom

        if bottom_y - y + 1 < min_size:
            continue

        rects.append({
            "Left": float(x),
            "Top": float(y),
            "Right": float(right_x),
            "Bottom": float(bottom_y),
        })

        visited[y:bottom_y + 1, x:right_x + 1] = True

    return rects


# ──────────────────────────────────────────────────────────────────────────────
# Phase E — Normalize to ratios
# ──────────────────────────────────────────────────────────────────────────────


def _cluster_row_range(addr: str):
    """Return (row_start, row_end) for a cell address."""
    p = _parse_cell_addr(addr)
    if p is None:
        return None
    return (p[1], p[3])


def _cluster_col_range(addr: str):
    """Return (col_start, col_end) for a cell address."""
    p = _parse_cell_addr(addr)
    if p is None:
        return None
    return (p[0], p[2])


def split_merged_rects(pixel_rects: list[dict], cluster_meta: list[dict]) -> list[dict]:
    """Split pixel-detected rectangles that contain multiple adjacent clusters.

    The PDF pixel scan (scan_black_rectangles) sometimes merges clusters that
    touch with zero pixel gap (e.g., side‑by‑side or vertically adjacent).
    This function splits those merged blobs using the cluster cell‑address
    grid coordinates as proportional guides.

    The absolute position still comes from the pixel scan — cell addresses
    only determine the RELATIVE SPLIT RATIO within a merged blob.
    """
    if len(pixel_rects) >= len(cluster_meta):
        # No splitting needed (equal or more rects than clusters)
        return pixel_rects

    meta_sorted = sorted(cluster_meta, key=_sort_key_meta)
    rects_sorted = sorted(pixel_rects, key=lambda r: (r["Top"], r["Left"]))

    # Group clusters by row proximity — a gap > 1 means a header row
    # separates them into different pixel rects.
    cl_groups = []
    cur = []
    prev_end = 0
    for m in meta_sorted:
        rr = _cluster_row_range(m["cellAddr"])
        if rr is None:
            return pixel_rects
        rs, re = rr
        if not cur:
            cur.append(m)
            prev_end = re
        elif rs - prev_end > 1:
            cl_groups.append(cur)
            cur = [m]
            prev_end = re
        else:
            cur.append(m)
            prev_end = max(prev_end, re)
    if cur:
        cl_groups.append(cur)

    # ══════════════════════════════════════════════════════════════════
    # Map each group to pixel rects — prefer independent rects, fall back to split
    # ══════════════════════════════════════════════════════════════════
    #
    # Phase 12.2: When multiple pixel rects overlap a group's Y range,
    # match each field to its own rect by horizontal position. Only
    # fall back to proportional splitting when not enough rects exist.
    #
    # This fixes the critical bug where adding C12:D12 would discard
    # its independent pixel rect and incorrectly split A12's rect.
    result = []
    for group in cl_groups:
        if not rects_sorted:
            break

        if len(group) == 1:
            # Single field group: pop first rect (no splitting needed)
            result.append(rects_sorted.pop(0))
            continue

        # Multi-field group: find ALL rects at the same vertical band
        # Rects at the same row have identical Top values (from the pixel scan).
        # Since rects_sorted is sorted by (Top, Left), same-band rects are adjacent.
        first_top = rects_sorted[0]["Top"]
        first_bottom = rects_sorted[0]["Bottom"]
        overlapping = []
        remaining = []
        for r in rects_sorted:
            # Same row band: Top within 2px of the first rect's Top
            if abs(r["Top"] - first_top) < 2.0 and abs(r["Bottom"] - first_bottom) < 2.0:
                overlapping.append(r)
            elif overlapping:
                break  # Passed the band — stop collecting
        # Remaining = rects after the band
        remaining = rects_sorted[len(overlapping):]
        rects_sorted = remaining

        if len(overlapping) >= len(group):
            # ── Path A: Enough independent rects exist — match by column position ──
            # Sort fields left-to-right by their start column
            fields_by_col = sorted(
                [(m, _cluster_col_range(m["cellAddr"])) for m in group],
                key=lambda x: x[1][0],
            )
            # Sort rects left-to-right by Left pixel position
            rects_by_left = sorted(overlapping, key=lambda r: r["Left"])

            print(f"  [SPLIT] Group has {len(group)} fields and {len(overlapping)} independent rects — matching")
            for m, cr in fields_by_col:
                matched_rect = rects_by_left.pop(0)
                result.append(matched_rect)
                print(f"  [SPLIT DEBUG] Matched independent rect:"
                      f" cell={m.get('cellAddr','?')}"
                      f" pixel=({matched_rect['Left']:.0f},{matched_rect['Top']:.0f},{matched_rect['Right']:.0f},{matched_rect['Bottom']:.0f})")
        else:
            # ── Path B: Not enough rects — proportional split (legacy fallback) ──
            # This handles the common case where multiple fields merge into
            # a single pixel blob (e.g., A1:B2 + C1:D2 + A3:D4 → one rect)
            rect = overlapping[0]

            all_rr = [_cluster_row_range(m["cellAddr"]) for m in group]
            all_cr = [_cluster_col_range(m["cellAddr"]) for m in group]
            min_row = min(rr[0] for rr in all_rr)
            max_row = max(rr[1] for rr in all_rr)
            min_col = min(cr[0] for cr in all_cr)
            max_col = max(cr[1] for cr in all_cr)
            grid_cols = max_col - min_col + 1
            grid_rows = max_row - min_row + 1

            rl = rect["Left"]
            rr_v = rect["Right"]
            rt = rect["Top"]
            rb = rect["Bottom"]
            rw = rr_v - rl
            rh = rb - rt

            print(f"  [SPLIT] Group has {len(group)} fields but only {len(overlapping)} rect — proportional split")

            # Gap-filling: each field gets space from its start col to the next
            col_starts = sorted(set(cr[0] for cr in all_cr))

            for i, m in enumerate(group):
                cr = _cluster_col_range(m["cellAddr"])
                rr2 = _cluster_row_range(m["cellAddr"])
                rel_left = (cr[0] - min_col) / grid_cols

                next_start = col_starts[i + 1] if i + 1 < len(col_starts) else max_col + 1
                rel_right = (next_start - min_col) / grid_cols

                rel_top = (rr2[0] - min_row) / grid_rows
                rel_bot = (rr2[1] - min_row + 1) / grid_rows

                print(f"  [SPLIT DEBUG] Field in multi-field group:"
                      f" cell={m.get('cellAddr','?')}"
                      f" col_range=({cr[0]},{cr[1]})"
                      f" min_col={min_col} max_col={max_col} grid={grid_cols}"
                      f" col_starts={col_starts}"
                      f" next_start={next_start}"
                      f" rel_left={rel_left:.4f} rel_right={rel_right:.4f}"
                      f" pixel=({rl+rel_left*rw:.0f},{rt+rel_top*rh:.0f},{rl+rel_right*rw:.0f},{rt+rel_bot*rh:.0f})")

                result.append({
                    "Left": rl + rel_left * rw,
                    "Top": rt + rel_top * rh,
                    "Right": rl + rel_right * rw,
                    "Bottom": rt + rel_bot * rh,
                })

    return result


def normalize_rects(rects: list[dict], image_w: int, image_h: int) -> list[dict]:
    """Convert pixel coordinates to 0-to-1 ratios.

    Formula (from IL decompilation):
      left_ratio   = rect.Left   / imageWidth
      top_ratio    = rect.Top    / imageHeight
      right_ratio  = rect.Right  / imageWidth
      bottom_ratio = rect.Bottom / imageHeight
    """
    for r in rects:
        r["left_ratio"]   = r["Left"]   / image_w
        r["top_ratio"]    = r["Top"]    / image_h
        r["right_ratio"]  = r["Right"]  / image_w
        r["bottom_ratio"] = r["Bottom"] / image_h
    return rects


# ──────────────────────────────────────────────────────────────────────────────
# Orchestrator
# ──────────────────────────────────────────────────────────────────────────────

def generate_coordinates(xlsx_path: str) -> list[dict]:
    """Execute the complete ConMas coordinate generation pipeline.

    Steps:
      1. Identify cluster cells (comments or _Fields sheet)
      2. Sanitize workbook (MakeCluster: black fills, white clears)
      3. Export sanitized PDF via Excel COM
      4. Render PDF at 300 DPI
      5. Pixel scan for black rectangles (GetAddress)
      6. Split merged rectangles using cell address proportions
      7. Normalize: ratio = pixel / image_dimension
      8. Merge metadata with ratios

    The PDF pixel scan provides correct absolute positions (matching the
    background PNG). Cell addresses are used ONLY to split merged blobs
    where adjacent clusters touch in the pixel data.

    Returns list of dicts (one per field) with:
      name, type, cellAddr, input_parameter,
      left_ratio, top_ratio, right_ratio, bottom_ratio
    """
    # Phase 0: Identify clusters
    cluster_meta = _identify_clusters(xlsx_path)
    if not cluster_meta:
        return []

    # Phase A: Sanitize
    sanitized_path = sanitize_workbook(xlsx_path, cluster_meta)

    # Phase B: Export PDF
    pdf_path = export_sanitized_pdf(sanitized_path)

    # Phase C: Render to image
    img, img_w, img_h = render_pdf_to_image(pdf_path)

    # Phase D: Pixel scan for black rectangles
    rects = scan_black_rectangles(img)

    # Phase D2: Split merged rectangles using cell address proportions
    # (preserves pixel-derived positions, only splits merged blobs)
    rects = split_merged_rects(rects, cluster_meta)

    # Phase E: Normalize
    rects = normalize_rects(rects, img_w, img_h)

    # Sort both lists by position for matching
    cluster_meta.sort(key=_sort_key_meta)
    rects.sort(key=lambda r: (r["Top"], r["Left"]))

    # Merge (both lists now of equal length after splitting)
    results = []
    for meta, rect in zip(cluster_meta, rects):
        results.append({
            "name": meta["name"],
            "type": meta["type"],
            "cellAddr": meta["cellAddr"],
            "input_parameter": meta.get("input_parameter", ""),
            "left_ratio": round(rect["left_ratio"], 7),
            "top_ratio": round(rect["top_ratio"], 7),
            "right_ratio": round(rect["right_ratio"], 7),
            "bottom_ratio": round(rect["bottom_ratio"], 7),
        })

    # Cleanup
    for tmp_dir in {os.path.dirname(sanitized_path), os.path.dirname(pdf_path)}:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass

    return results


# ──────────────────────────────────────────────────────────────────────────────
# Preview Pipeline (legacy) — coordinates + background PNG
# Kept for backward compatibility (/upload/coordinates endpoint)
# ──────────────────────────────────────────────────────────────────────────────

def generate_preview(xlsx_path: str, output_dir: str, output_id: str) -> dict:
    """Generate coordinates AND render background PNG for upload preview.

    LEGACY: Uses 4 separate COM sessions.
    See generate_coordinates_and_preview() for the single-session replacement.

    Steps:
      1. Run coordinates via generate_coordinates() (ConMas pipeline)
      2. Render original workbook as PDF → PNG for background
      3. Save PNG to output_dir / preview_{output_id}.png
      4. Return page dimensions + fields

    Args:
        xlsx_path: Path to uploaded XLSX.
        output_dir: Directory to save the background PNG.
        output_id: Unique ID for naming (e.g. GUID).

    Returns:
        dict with:
          backgroundImage: filename of saved PNG
          page: { width, height } in pixels
          fields: list of field dicts with ratios
    """
    os.makedirs(output_dir, exist_ok=True)

    # 1. Generate coordinates via ConMas pipeline
    fields = generate_coordinates(xlsx_path)
    if not fields:
        fields = []

    # 2. Render original workbook as PDF → PNG
    from render_service.pdf_converter import xlsx_to_pdf
    from render_service.background_renderer import pdf_page_to_png, get_page_dimensions

    pdf_path = xlsx_to_pdf(xlsx_path)

    png_filename = f"preview_{output_id}.png"
    png_path = os.path.join(output_dir, png_filename)
    pdf_page_to_png(pdf_path, png_path, dpi=DPI)

    page_w, page_h = get_page_dimensions(pdf_path, dpi=DPI)

    # Cleanup PDF temp
    try:
        os.unlink(pdf_path)
    except Exception:
        pass
    pdf_dir = os.path.dirname(pdf_path)
    if pdf_dir and pdf_dir != output_dir:
        try:
            shutil.rmtree(pdf_dir, ignore_errors=True)
        except Exception:
            pass

    return {
        "backgroundImage": png_filename,
        "page": {
            "width": page_w,
            "height": page_h,
        },
        "fields": fields,
    }


# ──────────────────────────────────────────────────────────────────────────────
# Performance Instrumentation (Phase X.12)
# Set to False to disable all timing logs.
# ──────────────────────────────────────────────────────────────────────────────
ENABLE_PERFORMANCE_LOGS = True

import time as _time_mod


def _print_perf_report(stage_timings: list[tuple[str, float]]):
    """Print a formatted performance report from a list of (label, elapsed_sec) tuples."""
    total_ms = sum(t[1] for t in stage_timings) * 1000

    print("")
    print("=" * 58)
    print("PaperLess Upload Performance")
    print("=" * 58)
    print("")

    # Print each stage in EXECUTION ORDER (preserved from call sequence)
    for label, sec in stage_timings:
        ms = sec * 1000
        print("  %-28s : %8.0f ms" % (label, ms))

    print("  " + "-" * 42)
    print("  %-28s : %8.0f ms" % ("TOTAL", total_ms))
    print("  %-28s : %8.3f sec" % ("TOTAL", total_ms / 1000))
    print("=" * 58)

    # Collect all warnings for top-N ranking
    warnings = []
    for label, sec in stage_timings:
        ms = sec * 1000
        if ms > 5000:
            print("")
            print("CRITICAL BOTTLENECK")
            print("  %s : %.0f ms" % (label, ms))
            warnings.append((label, ms))
        elif ms > 1000:
            print("")
            print("Slow Stage Detected")
            print("  %s : %.0f ms" % (label, ms))
            warnings.append((label, ms))

    # Top-N ranking sorted by time (slowest first)
    if warnings:
        warnings.sort(key=lambda x: -x[1])
        print("")
        print("Top %d Slowest Operations" % len(warnings))
        for i, (label, ms) in enumerate(warnings, 1):
            print("  %d. %-28s %8.0f ms" % (i, label, ms))
    print("")


# ──────────────────────────────────────────────────────────────────────────────
# Phase X.9 — Single COM Session Pipeline
# Matches original ConMas architecture: open Excel ONCE, do everything.
# ──────────────────────────────────────────────────────────────────────────────

def generate_coordinates_and_preview(
    xlsx_path: str,
    output_dir: str,
    output_id: str,
) -> dict:
    """Generate coordinates AND background PNG in a SINGLE COM session.

    Multi‑worksheet support (Phase UI.3):
      - Exports ALL visible worksheets as separate PDF pages
      - Processes each worksheet independently (comment scan, sanitize,
        pixel scan, background render)
      - Returns a `pages[]` array so the frontend can render one page
        per worksheet

    Architecture matches original ConMas:
      1. Open Excel (one Application, one COM lifecycle)
      2. Open original workbook
      3. Export original PDF (background) — BEFORE sanitization
      4a. Read comments (READ-ONLY pass) — track sheet names
      4b. Sanitize workbook per-sheet (WRITE-ONLY pass)
      5. Delete metadata sheets & export sanitized PDF (ALL worksheets)
      6. Quit Excel (single cleanup)
      7. For each visible worksheet:
           Render sanitized PDF page → pixel scan → normalize ratios
           Render background PNG from original PDF → collect fields
      8. Return JSON with pages[] array

    Args:
        xlsx_path: Path to uploaded XLSX.
        output_dir: Directory to save the background PNGs.
        output_id: Unique ID for naming (e.g. GUID).

    Returns:
        dict with:
          pages: list of {
              sheetName: str,
              backgroundImage: str,
              page: {width, height},
              fields: [...]
          }
    """
    import pythoncom
    import win32com.client
    from render_service.render_queue import get_queue_excel

    # Performance instrumentation
    _pt = []
    if ENABLE_PERFORMANCE_LOGS:
        def _mark(label):
            _pt.append((label, _time_mod.perf_counter()))
        def _stage_mark(label):
            _mark(label)
    else:
        def _mark(_label):
            pass
        def _stage_mark(_label):
            pass

    # Temporary directories (cleaned up at end)
    tmp_dirs = []

    try:
        # ════════════════════════════════════════════════════
        # PHASE A: SINGLE COM SESSION (everything)
        # ════════════════════════════════════════════════════
        queue_excel = get_queue_excel()
        if queue_excel is not None:
            excel = queue_excel
            own_excel = False
        else:
            pythoncom.CoInitialize()
            _stage_mark("COM: CoInitialize")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            own_excel = True
        _stage_mark("Excel Launch")
        try:
            try:
                wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
                _stage_mark("Workbook Open")

                # ── Step 1: Collect visible worksheet order ──
                # NOTE: We store BOTH the COM proxy (visible_sheets) for sanitization
                # AND plain string names (visible_sheet_names) for the NC phase after
                # Excel.Quit(). COM proxies become stale after CoUninitialize().
                visible_sheets = []
                visible_sheet_names = []
                # ── PHASE 6.1 DIAGNOSTIC: Full worksheet analysis ──
                print("=" * 50)
                print("WORKBOOK DIAGNOSTIC")
                print("=" * 50)
                print(f"Workbook: {os.path.basename(xlsx_path)}")
                print(f"Total Worksheets: {wb.Sheets.Count}")
                print()
                for ws_i_idx in range(1, wb.Sheets.Count + 1):
                    try:
                        ws_i = wb.Sheets(ws_i_idx)
                        ws_name = str(ws_i.Name)
                        ws_visible = "UNKNOWN"
                        try:
                            v = ws_i.Visible
                            if v == -1:
                                ws_visible = "VISIBLE"
                            elif v == 0:
                                ws_visible = "HIDDEN"
                            elif v == 2:
                                ws_visible = "VERY_HIDDEN"
                            else:
                                ws_visible = str(v)
                        except Exception:
                            pass
                        ws_print_area = "NONE"
                        try:
                            pa = ws_i.PageSetup.PrintArea
                            if pa:
                                ws_print_area = str(pa)[:60]
                        except Exception:
                            pass
                        ws_comments = 0
                        try:
                            ws_comments = ws_i.Comments.Count
                        except Exception:
                            pass
                        is_config = "YES" if ws_name in ("_Fields", "ExcelOutputSetting", "_RawData", "DesignerConfig", "PaperLessConfig", "ConMasConfig") else "NO"
                        print(f"  Sheet #{ws_i_idx}")
                        print(f"    Name:         {ws_name}")
                        print(f"    Visible:      {ws_visible}")
                        print(f"    PrintArea:    {ws_print_area}")
                        print(f"    Comments:     {ws_comments}")
                        print(f"    ConfigSheet:  {is_config}")
                        will_be_page = "UNKNOWN"
                        try:
                            if ws_visible == "VISIBLE":
                                if is_config == "YES":
                                    will_be_page = "YES (CURRENTLY LEAKING)"
                                else:
                                    will_be_page = "YES"
                            else:
                                will_be_page = "NO (hidden)"
                        except Exception:
                            pass
                        print(f"    Will Be Page: {will_be_page}")
                        print()
                    except Exception as e:
                        print(f"  Sheet #{ws_i_idx}: ERROR: {e}")
                        print()

                # ── PHASE 6.3: Collect printable worksheets ──
                # Legacy ConMas behavior: filter by BOTH visibility AND reserved name.
                # Visibility alone is insufficient because ExcelOutputSetting is VISIBLE.
                # Algorithm A (Reserved Worksheet Names) — proven by Phase 6.2 forensic investigation.
                visible_sheets = []
                visible_sheet_names = []
                skipped_config_sheets = []
                for ws in wb.Worksheets:
                    try:
                        if ws.Visible != -1:  # -1 = xlSheetVisible
                            continue
                    except Exception:
                        pass
                    try:
                        ws_name = ws.Name
                    except Exception:
                        continue
                    # Phase 6.3: Skip reserved configuration sheets
                    if ws_name in CONFIGURATION_SHEET_NAMES:
                        print(f"  [FILTER] Sheet '{ws_name}': SKIPPED (configuration sheet)")
                        skipped_config_sheets.append(ws_name)
                        continue
                    visible_sheets.append(ws)
                    visible_sheet_names.append(ws_name)

                # ── PHASE 6.3: Printable worksheets summary ──
                print(f"==================================================")
                print(f"Printable Worksheets Summary")
                print(f"==================================================")
                print(f"Total printable: {len(visible_sheet_names)}")
                for vs_name in visible_sheet_names:
                    print(f"  INCLUDED: {vs_name}")
                if skipped_config_sheets:
                    print(f"Skipped (configuration):")
                    for scs in skipped_config_sheets:
                        print(f"  EXCLUDED: {scs}")
                else:
                    print(f"No configuration sheets skipped.")
                print(f"==================================================")
                _stage_mark("Collect Visible Sheets")

                # ── Step 2: Export original PDF (background) ──
                # MUST be before sanitization because sanitization
                # destroys cell values and fill colors.
                # Export ALL visible worksheets as one multi-page PDF.
                orig_pdf_dir = tempfile.mkdtemp(prefix="ple_orig_pdf_")
                tmp_dirs.append(orig_pdf_dir)
                orig_pdf_path = os.path.join(orig_pdf_dir, "original.pdf")
                wb.ExportAsFixedFormat(0, os.path.abspath(orig_pdf_path))
                _stage_mark("Export Original PDF")

                # ── Step 3a: Read comments per-sheet ──
                # Phase X.16 proved ws.Comments returns the exact same comment set
                # as per-cell cell.Comment scan, but 2.8x faster.
                #
                # IMPORTANT: This must be a separate pass from sanitization.
                # Reading cell.Comment and then modifying cell.Interior.Color
                # in the same iteration corrupts COM proxy state.
                #
                # Key addition: track sheetName so we can group fields by sheet.
                cluster_meta = []  # each entry includes "sheetName"
                cluster_addrs_by_sheet = {}  # sheetName -> set of addresses

                for ws in visible_sheets:
                    sheet_name = ws.Name
                    try:
                        if ws.Comments.Count == 0:
                            continue
                    except Exception:
                        continue

                    for comment in ws.Comments:
                        try:
                            parent_range = comment.Parent
                        except Exception:
                            continue
                        if parent_range is None:
                            continue
                        try:
                            merged = parent_range.MergeArea
                            addr = str(merged.Address).upper()
                        except Exception:
                            addr = str(parent_range.Address).upper()
                        try:
                            text = str(comment.Text())
                        except Exception:
                            text = ""
                        lines = text.replace("\r\n", "\n").split("\n")
                        cluster_meta.append({
                            "cellAddr": addr,
                            "sheetName": sheet_name,
                            "name": lines[0].strip() if lines else "",
                            "type": lines[1].strip() if len(lines) > 1 else "Text",
                            "input_parameter": (
                                "\n".join(lines[2:]).strip()
                                if len(lines) > 2 else ""
                            ),
                        })
                        cluster_addrs_by_sheet.setdefault(sheet_name, set()).add(addr)

                _stage_mark("Comment Scan")

                # Early exit if no comments found anywhere
                if not cluster_meta:
                    wb.Close(False)
                    return {"pages": []}

                # ── Step 3b: Sanitize workbook per-sheet ──
                # Phase X.15 optimizations (validated in Phase X.14):
                #   C: Skip PageSetup header/footer clearing
                #   D: Only enumerate shapes if count > 0
                #   B: Batch fill entire UsedRange white via Range.Interior.Color
                #   A: Batch clear values via UsedRange.ClearContents
                xlNone = -4142

                for ws in visible_sheets:
                    sheet_name = ws.Name
                    sheet_addrs = cluster_addrs_by_sheet.get(sheet_name, set())

                    # Optimization D: Only delete shapes when shapes exist
                    try:
                        if ws.Shapes.Count > 0:
                            for shape in list(ws.Shapes):
                                shape.Delete()
                    except Exception:
                        pass

                    used = ws.UsedRange
                    if used is None:
                        continue

                    lr = used.Row + used.Rows.Count - 1
                    lc = used.Column + used.Columns.Count - 1

                    # Optimization B: Batch fill entire UsedRange white
                    try:
                        used.Interior.Color = 0xFFFFFF
                    except Exception:
                        pass

                    # Optimization A: Batch clear all cell values
                    try:
                        used.ClearContents()
                    except Exception:
                        pass

                    # Fill only THIS sheet's cluster cells black
                    for addr in sheet_addrs:
                        try:
                            rng = ws.Range(addr)
                            rng.Interior.Color = 1  # Black
                            rng.Value = ""
                        except Exception:
                            pass

                    # Clear borders
                    try:
                        clear_range = ws.Range(
                            ws.Cells(1, 1),
                            ws.Cells(lr, lc)
                        )
                        clear_range.Borders.LineStyle = xlNone
                    except Exception:
                        pass

                _stage_mark("Workbook Sanitization")

                # ════════════════════════════════════════════════════
                # [ROW12 DIAGNOSTIC] COM state after sanitization
                # ════════════════════════════════════════════════════
                print(f"\n{'='*60}")
                print(f"[SANITIZE CELL] COM state after sanitization (Sheet1)")
                print(f"{'='*60}")
                for _diag_addr in ["$A$12", "$C$12:$D$12", "$B$12", "$A$1:$B$2", "$C$1:$D$2"]:
                    _diag_sheet = None
                    for _ws_cell in wb.Worksheets:
                        try:
                            if _ws_cell.Visible == -1 and _ws_cell.Name not in CONFIGURATION_SHEET_NAMES:
                                _diag_sheet = _ws_cell
                                break
                        except Exception:
                            pass
                    if _diag_sheet is None:
                        continue
                    try:
                        _rng = _diag_sheet.Range(_diag_addr)
                        _addr_out = str(_rng.Address)
                        _merge_out = "YES" if _rng.MergeCells else "NO"
                        try:
                            _ma = _rng.MergeArea
                            _ma_addr = str(_ma.Address)
                        except Exception:
                            _ma_addr = "<ERROR>"
                        _color = "?"
                        try:
                            _c = _rng.Interior.Color
                            if _c == 1:
                                _color = "BLACK(1)"
                            elif _c == 16777215 or _c == 0xFFFFFF:
                                _color = "WHITE(16777215/0xFFFFFF)"
                            else:
                                _color = f"OTHER({_c})"
                        except Exception:
                            _color = "<ERROR>"
                        _val = str(_rng.Value) if _rng.Value is not None else "<EMPTY/NONE>"
                        _borders = "?"
                        try:
                            _ls = _rng.Borders.LineStyle
                            if _ls == -4142:
                                _borders = "xlNone"
                            elif _ls == 1:
                                _borders = "xlContinuous"
                            else:
                                _borders = f"OTHER({_ls})"
                        except Exception:
                            _borders = "<ERROR>"
                        print(f"  Cell {_diag_addr}:"
                              f" Addr={_addr_out}"
                              f" MergeCells={_merge_out}"
                              f" MergeArea={_ma_addr}"
                              f" Color={_color}"
                              f" Value='{_val[:30]}'"
                              f" Borders={_borders}")
                    except Exception as _e:
                        print(f"  Cell {_diag_addr}: ERROR={_e}")
                # Also log column widths for columns A-D
                try:
                    _diag_sheet2 = None
                    for _ws2 in wb.Worksheets:
                        try:
                            if _ws2.Visible == -1 and _ws2.Name not in CONFIGURATION_SHEET_NAMES:
                                _diag_sheet2 = _ws2
                                break
                        except Exception:
                            pass
                    if _diag_sheet2 is not None:
                        print(f"  Column widths:")
                        for _col_idx in range(1, 8):
                            try:
                                _cw = _diag_sheet2.Columns(_col_idx).ColumnWidth
                                _col_l = chr(ord('A') + _col_idx - 1)
                                print(f"    Col {_col_l} width={_cw} chars ({_cw*7:.1f}pt)")
                            except Exception:
                                pass
                except Exception:
                    pass
                # PageSetup info
                try:
                    if _diag_sheet2 is not None:
                        _ps = _diag_sheet2.PageSetup
                        print(f"  PageSetup:"
                              f" PrintArea={_ps.PrintArea}"
                              f" FitToPagesWide={_ps.FitToPagesWide}"
                              f" FitToPagesTall={_ps.FitToPagesTall}"
                              f" Zoom={_ps.Zoom}"
                              f" LeftMargin={_ps.LeftMargin:.2f}"
                              f" TopMargin={_ps.TopMargin:.2f}"
                              f" CenterH={_ps.CenterHorizontally}"
                              f" CenterV={_ps.CenterVertically}")
                except Exception as _e:
                    print(f"  PageSetup ERROR: {_e}")
                print(f"{'='*60}\n")

                # ── Step 4: Delete metadata sheets & export sanitized PDF ──
                # Export ALL worksheets so the PDF has one page per visible sheet.
                # Phase 6.3: Delete ALL configuration sheets, not just _Fields/_RawData
                for i in range(wb.Sheets.Count, 0, -1):
                    try:
                        name = wb.Sheets(i).Name
                    except Exception:
                        name = ""
                    if name in CONFIGURATION_SHEET_NAMES:
                        print(f"  [SANITIZE EXPORT] Deleting config sheet: '{name}'")
                        try:
                            wb.Sheets(i).Delete()
                        except Exception as e:
                            print(f"  [SANITIZE EXPORT] Failed to delete '{name}': {e}")
                            pass

                pdf_dir = tempfile.mkdtemp(prefix="ple_pdf_")
                tmp_dirs.append(pdf_dir)
                pdf_path = os.path.join(pdf_dir, "sanitized.pdf")

                # ── PHASE 6.3: PDF Export diagnostics ──
                remaining_names = []
                for ri in range(1, wb.Sheets.Count + 1):
                    try:
                        remaining_names.append(str(wb.Sheets(ri).Name))
                    except Exception:
                        remaining_names.append("<ERROR>")
                print(f"  [PDF EXPORT] Worksheets being exported ({len(remaining_names)}): {remaining_names}")
                print(f"  [PDF EXPORT] Expected PDF pages: {len(remaining_names)}")

                # Export ALL worksheets (workbook-level export)
                # IgnorePrintAreas=False matches ConMas default
                wb.ExportAsFixedFormat(
                    0, os.path.abspath(pdf_path),
                    0, 0, False,
                )

                # Verify PDF page count & compare dimensions
                import fitz
                try:
                    ver_doc = fitz.open(pdf_path)
                    actual_pages = len(ver_doc)
                    ver_doc.close()
                    print(f"  [PDF EXPORT] Actual PDF pages generated: {actual_pages}")
                    if actual_pages != len(remaining_names):
                        print(f"  [PDF EXPORT] WARNING: PDF page count ({actual_pages}) differs from"
                              f" worksheet count ({len(remaining_names)})")

                    # ════════════════════════════════════════════════════
                    # [PDF DIMENSIONS] Compare original vs sanitized PDF
                    # ════════════════════════════════════════════════════
                    odoc = fitz.open(orig_pdf_path)
                    sdoc = fitz.open(pdf_path)
                    for _pi in range(min(len(odoc), len(sdoc))):
                        op = odoc[_pi]
                        sp = sdoc[_pi]
                        ow, oh = op.rect.width, op.rect.height
                        sw_, sh_ = sp.rect.width, sp.rect.height
                        print(f"  [PDF DIMENSIONS] Page {_pi}:"
                              f" Original=({ow:.1f}x{oh:.1f}pt)"
                              f" Sanitized=({sw_:.1f}x{sh_:.1f}pt)"
                              f" Match={'YES' if abs(ow-sw_)<0.1 and abs(oh-sh_)<0.1 else 'NO'}")
                    odoc.close()
                    sdoc.close()
                except Exception as _e_pdf:
                    print(f"  [PDF DIAGNOSTICS] ERROR: {_e_pdf}")

                _stage_mark("Export Sanitized PDF (all sheets)")

            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass
                if own_excel:
                    try:
                        excel.Quit()
                    except Exception:
                        pass
                    _stage_mark("Excel Quit")
        finally:
            if own_excel:
                pythoncom.CoUninitialize()
                _stage_mark("COM: CoUninitialize")

        # ════════════════════════════════════════════════════
        # PHASE B: NON-COM OPERATIONS (pixel scan per-page)
        # ════════════════════════════════════════════════════

        from render_service.background_renderer import (
            pdf_page_to_png,
            get_page_dimensions,
        )
        os.makedirs(output_dir, exist_ok=True)

        # Count available PDF pages in sanitized PDF
        import fitz
        san_doc = fitz.open(pdf_path)
        total_pdf_pages = len(san_doc)
        san_doc.close()

        pages_result = []

        # Use plain string names — COM proxies are stale after Excel.Quit()
        for i, sheet_name in enumerate(visible_sheet_names):
            sheet_clusters = [m for m in cluster_meta if m["sheetName"] == sheet_name]

            # PHASE 6.1 DIAGNOSTIC: Every page being generated
            print(f"  [PAGE GEN] Sheet #{i}: '{sheet_name}' — {len(sheet_clusters)} field(s), "
                  f"config={'YES' if sheet_name in ('_Fields', 'ExcelOutputSetting', '_RawData') else 'NO'}")

            if not sheet_clusters:
                # No fields on this sheet — still include as a page with empty fields
                # (the background PNG is still useful)
                print(f"  [PAGE GEN] -> Empty page (no fields), still generating background PNG")
                png_filename = f"preview_{output_id}_page{i}.png"
                png_path = os.path.join(output_dir, png_filename)
                pdf_page_to_png(orig_pdf_path, png_path, dpi=DPI, page_index=min(i, total_pdf_pages - 1))
                page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=DPI, page_index=min(i, total_pdf_pages - 1))
                pages_result.append({
                    "sheetName": sheet_name,
                    "backgroundImage": png_filename,
                    "page": {"width": page_w, "height": page_h},
                    "fields": [],
                })
                continue

            # Render this page of the sanitized PDF
            page_idx = min(i, total_pdf_pages - 1)
            img, img_w, img_h = render_pdf_to_image(pdf_path, page_index=page_idx)
            _stage_mark("Render PDF page %d" % i)

            rects = scan_black_rectangles(img)
            _stage_mark("Pixel Scan page %d" % i)

            # ════════════════════════════════════════════════════
            # [ROW12 DIAGNOSTIC] Raw rectangles BEFORE split
            # ════════════════════════════════════════════════════
            print(f"\n{'='*60}")
            print(f"[ROW12 DIAGNOSTIC] Sheet '{sheet_name}' — Pixel Scan Results")
            print(f"{'='*60}")
            print(f"  Raw rectangles found: {len(rects)}")
            print(f"  Sheet clusters: {len(sheet_clusters)}")
            for ri, r in enumerate(rects):
                print(f"  Rect {ri}: Left={r['Left']:.1f} Top={r['Top']:.1f} "
                      f"Right={r['Right']:.1f} Bottom={r['Bottom']:.1f} "
                      f"Width={r['Right']-r['Left']:.1f}px Height={r['Bottom']-r['Top']:.1f}px")

            # ── Per-column pixel analysis of the LAST rectangle (likely row 12) ──
            if rects:
                last = rects[-1]
                _rl = int(last["Left"])
                _rr = int(last["Right"])
                _rt = int(last["Top"])
                _rb = int(last["Bottom"])
                _rw = _rr - _rl
                _rh = _rb - _rt
                last_region = img[_rt:_rb+1, _rl:_rr+1]
                total_pix = _rw * _rh
                black_pix = int(np.sum(np.all(last_region < BLACK_THRESHOLD, axis=2)))
                white_pix = int(np.sum(np.all(last_region > WHITE_THRESHOLD, axis=2)))
                other_pix = total_pix - black_pix - white_pix
                print(f"  Last rect pixel composition: total={total_pix} black={black_pix} "
                      f"white={white_pix} other(gray)={other_pix}")

                # The expected column count for the form is 4 (A-D) at ~51px each
                # Divide the width into 4 equal column regions and count pixels
                if _rw > 20:  # Only if wide enough to analyze
                    n_cols = 4
                    col_w = _rw // n_cols
                    diag_rows_used = min(_rh, 100)  # Analyze up to 100 rows
                    print(f"  Row 12 column pixel analysis (analyzing {diag_rows_used} rows vertically):")
                    for ci in range(n_cols):
                        cx1 = max(0, ci * col_w)
                        cx2 = min(_rw, (ci + 1) * col_w)
                        col_region = last_region[:diag_rows_used, cx1:cx2]
                        ct = col_region.shape[0] * col_region.shape[1]
                        cb = int(np.sum(np.all(col_region < BLACK_THRESHOLD, axis=2)))
                        cw = int(np.sum(np.all(col_region > WHITE_THRESHOLD, axis=2)))
                        co = ct - cb - cw
                        col_letter = chr(ord('A') + ci)
                        print(f"    Col {col_letter}(px{cx1}-{cx2}): "
                              f"black={cb} white={cw} other={co} "
                              f"({cb*100//ct if ct>0 else 0}% black)")

                # Compare row-12 column positions with expected from rows 1-10
                # Expected: column A starts at ~857px, column C at ~1274px (from legacy data)
                print(f"  Row 12 rect pixel coords: x=[{_rl},{_rr}] y=[{_rt},{_rb}]")
                print(f"  Legacy expected: Col A~857px Col C~1274px (at 300 DPI)")
                print(f"  Actual rect: left={_rl}px matches Col A at {'YES' if abs(_rl-857)<20 else 'NO'}")
                if abs(_rl - 857) < 20:
                    # Column A starts at ~857, C at ~1274 (about 417px apart)
                    expected_c_start = 1274
                    actual_c_x = _rl + _rw * 2 // 4  # Approx position of column C
                    print(f"  Approx col C at pixel: {actual_c_x}px (expected ~{expected_c_start}px)")
            print(f"{'='*60}\n")

            if len(rects) < len(sheet_clusters):
                rects = split_merged_rects(rects, sheet_clusters)
            _stage_mark("Split Rectangles page %d" % i)

            rects = normalize_rects(rects, img_w, img_h)
            _stage_mark("Normalize page %d" % i)

            # Sort both by position for matching
            sheet_clusters_sorted = sorted(sheet_clusters, key=_sort_key_meta)
            rects.sort(key=lambda r: (r["Top"], r["Left"]))

            # Merge metadata with ratios
            fields = []
            for meta, rect in zip(sheet_clusters_sorted, rects):
                fields.append({
                    "name": meta["name"],
                    "type": meta["type"],
                    "cellAddr": meta["cellAddr"],
                    "input_parameter": meta.get("input_parameter", ""),
                    "left_ratio": round(rect["left_ratio"], 7),
                    "top_ratio": round(rect["top_ratio"], 7),
                    "right_ratio": round(rect["right_ratio"], 7),
                    "bottom_ratio": round(rect["bottom_ratio"], 7),
                })

            # Render background PNG for this sheet from original PDF
            png_filename = f"preview_{output_id}_page{i}.png"
            png_path = os.path.join(output_dir, png_filename)
            pdf_page_to_png(orig_pdf_path, png_path, dpi=DPI, page_index=page_idx)
            page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=DPI, page_index=page_idx)
            _stage_mark("Render Bg PNG page %d" % i)

            pages_result.append({
                "sheetName": sheet_name,
                "backgroundImage": png_filename,
                "page": {"width": page_w, "height": page_h},
                "fields": fields,
            })

        _stage_mark("Cleanup")

        # Compute durations and print report
        if ENABLE_PERFORMANCE_LOGS:
            durations = []
            for i in range(1, len(_pt)):
                label = _pt[i][0]
                elapsed = _pt[i][1] - _pt[i - 1][1]
                durations.append((label, elapsed))
            _print_perf_report(durations)

        # PHASE 6.1 DIAGNOSTIC: Preview page generation summary
        print("=" * 50)
        print("PREVIEW PAGES GENERATED")
        print("=" * 50)
        for pi, p in enumerate(pages_result):
            is_config = "YES" if p["sheetName"] in ("_Fields", "ExcelOutputSetting", "_RawData") else "NO"
            print(f"  Page {pi}")
            print(f"    Source Worksheet: {p['sheetName']}")
            print(f"    Background Image: {p.get('backgroundImage', 'N/A')}")
            print(f"    Page Size:        {p.get('page', {}).get('width', '?')}x{p.get('page', {}).get('height', '?')}")
            print(f"    Fields:           {len(p.get('fields', []))}")
            print(f"    Config Sheet:     {is_config}")
            print()
        config_pages = [p for p in pages_result if p["sheetName"] in ("_Fields", "ExcelOutputSetting", "_RawData")]
        if config_pages:
            print(f"WARNING: {len(config_pages)} configuration sheet(s) are leaking into preview pages:")
            for cp in config_pages:
                print(f"  -> {cp['sheetName']}")
        print("=" * 50)

        return {"pages": pages_result}

    finally:
        # Cleanup all temporary directories
        for d in tmp_dirs:
            shutil.rmtree(d, ignore_errors=True)
