"""
ConMas Upload Coordinate Generation — MakeCluster + pixel scan + ratio normalization.

Replicates the original ConMas Designer's CalcClusterSize pipeline EXACTLY:

  MakeCluster() → ExportPdf() → CreateImageFromPdf() → GetAddress() → normalize()

No worksheet geometry, no column widths, no margins, no calibration.
Pure pixel scanning of a sanitized workbook rendered at 300 DPI.

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
    """Scan every cell via COM for cell comments (original MakeCluster).

    Returns list of dicts with cellAddr, name, type, input_parameter.
    Returns [] if no comments found.
    """
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    clusters = []
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        for ws in wb.Worksheets:
            used = ws.UsedRange
            if used is None:
                continue
            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1
            for row in range(1, lr + 1):
                for col in range(1, lc + 1):
                    cell = ws.Cells(row, col)
                    try:
                        comment = cell.Comment
                    except Exception:
                        comment = None
                    if comment is None:
                        continue
                    try:
                        ma = cell.MergeArea
                    except Exception:
                        ma = cell
                    try:
                        addr = str(ma.Address)
                    except Exception:
                        addr = str(cell.Address)
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
        wb.Close(False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
    return clusters


def _identify_clusters_from_fields_sheet(xlsx_path: str):
    """Read cluster ranges from hidden _Fields sheet.

    The _Fields sheet has columns:
      A=Address, B=FieldId, C=FieldName, D=FieldType, E=SheetName, ...

    Returns list of dicts with cellAddr, name, type, input_parameter.
    Returns [] if no _Fields sheet or no data rows.
    """
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    clusters = []
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
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
            if not cell_addr or not re.search(r"\$[A-Z]+\$\d+", cell_addr):
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
        wb.Close(False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
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

    Args:
        xlsx_path: Path to original workbook.
        clusters: List of dicts with cellAddr to fill black.

    Returns:
        Path to sanitized temporary XLSX.
    """
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False

    tmp_dir = tempfile.mkdtemp(prefix="ple_sanitize_")
    sanitized_path = os.path.join(tmp_dir, "sanitized.xlsx")

    # Build set of cluster ranges for quick lookup
    cluster_addrs = set()
    sheet_clusters = {}
    for c in clusters:
        addr = c["cellAddr"]
        cluster_addrs.add(addr.upper().replace(" ", ""))
        sheet_name = c.get("sheetName", "")
        if sheet_name not in sheet_clusters:
            sheet_clusters[sheet_name] = []
        sheet_clusters[sheet_name].append(addr)

    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

        for ws in wb.Worksheets:
            ws.PageSetup.CenterHeader = ""
            ws.PageSetup.CenterFooter = ""

            try:
                for shape in list(ws.Shapes):
                    shape.Delete()
            except Exception:
                pass

            used = ws.UsedRange
            if used is None:
                continue

            lr = used.Row + used.Rows.Count - 1
            lc = used.Column + used.Columns.Count - 1

            xlNone = -4142

            for row in range(1, lr + 1):
                for col in range(1, lc + 1):
                    cell = ws.Cells(row, col)

                    try:
                        ma = cell.MergeArea
                    except Exception:
                        ma = None

                    if ma is not None:
                        try:
                            ma_addr = str(ma.Address).upper()
                        except Exception:
                            ma_addr = ""
                        is_cluster = ma_addr in cluster_addrs
                    else:
                        try:
                            cell_addr = str(cell.Address).upper()
                        except Exception:
                            cell_addr = ""
                        is_cluster = cell_addr in cluster_addrs

                    if is_cluster:
                        fill_cell = ma if ma is not None else cell
                        try:
                            fill_cell.Interior.Color = 1
                        except Exception:
                            pass
                        try:
                            fill_cell.Value = ""
                        except Exception:
                            pass
                    else:
                        try:
                            cell.Interior.Color = 0xFFFFFF
                        except Exception:
                            pass
                        try:
                            cell.Value = ""
                        except Exception:
                            pass

            try:
                clear_range = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc))
                clear_range.Borders.LineStyle = xlNone
            except Exception:
                pass

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


# ──────────────────────────────────────────────────────────────────────────────
# Phase B — Export sanitized workbook to PDF via Excel COM
# ──────────────────────────────────────────────────────────────────────────────

def export_sanitized_pdf(sanitized_path: str) -> str:
    """Export sanitized workbook to PDF via Excel COM ExportAsFixedFormat."""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False

    tmp_dir = tempfile.mkdtemp(prefix="ple_export_pdf_")
    pdf_path = os.path.join(tmp_dir, "sanitized.pdf")

    try:
        wb = excel.Workbooks.Open(os.path.abspath(sanitized_path))

        # Delete metadata sheets to avoid multi-page PDF
        for i in range(wb.Sheets.Count, 0, -1):
            name = ""
            try:
                name = wb.Sheets(i).Name
            except Exception:
                pass
            if name in ("_Fields", "_RawData"):
                wb.Sheets(i).Delete()

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


# ──────────────────────────────────────────────────────────────────────────────
# Phase C — Render PDF to image at 300 DPI
# ──────────────────────────────────────────────────────────────────────────────

def render_pdf_to_image(pdf_path: str, dpi: int = DPI):
    """Render PDF page 0 at 300 DPI to numpy array.

    Returns (numpy_array_HWC, width_px, height_px) in RGB uint8.
    """
    if not HAS_NUMPY:
        raise RuntimeError("numpy is required for pixel scanning")
    import fitz

    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        w, h = pix.width, pix.height
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, 3)
    finally:
        doc.close()

    return img, w, h


# ──────────────────────────────────────────────────────────────────────────────
# Phase D — Pixel scan for black rectangles (GetAddress)
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
      1. Top→bottom, left→right pixel scan
      2. Detect top-left corner: BLACK pixel with NON-BLACK left AND above
      3. 6-pixel noise verification
      4. Expand right until not-black → right edge
      5. Expand down (per column, take max Y) → bottom edge
      6. Minimum 6×6 pixel filter
      7. Mark visited rect

    Returns list of dicts with Left, Top, Right, Bottom (pixel coords).
    """
    h, w, _ = img.shape
    visited = np.zeros((h, w), dtype=bool)
    rects = []

    for y in range(1, h - 1):
        for x in range(1, w - 1):
            if visited[y, x]:
                continue

            r, g, b = img[y, x]
            if not _is_black(r, g, b):
                continue

            r_left, g_left, b_left = img[y, x - 1]
            r_up, g_up, b_up = img[y - 1, x]
            if not (_is_not_black(r_left, g_left, b_left)
                    and _is_not_black(r_up, g_up, b_up)):
                continue

            noise = False
            for dx in range(1, 7):
                cx = x + dx
                if cx >= w:
                    break
                rr, gg, bb = img[y - 1, cx]
                if _is_black(rr, gg, bb):
                    noise = True
                    break
            if noise:
                continue

            for dy in range(1, 7):
                cy = y + dy
                if cy >= h:
                    break
                rr, gg, bb = img[cy, x - 1]
                if _is_black(rr, gg, bb):
                    noise = True
                    break
            if noise:
                continue

            right_x = x + 1
            while right_x < w:
                rr, gg, bb = img[y, right_x]
                if _is_white(rr, gg, bb):
                    break
                right_x += 1
            right_x -= 1

            if right_x - x + 1 < min_size:
                continue

            bottom_y = y
            for col in range(x, right_x + 1):
                scan_y = y + 1
                while scan_y < h:
                    rr, gg, bb = img[scan_y, col]
                    if _is_white(rr, gg, bb):
                        break
                    scan_y += 1
                bottom_y = max(bottom_y, scan_y - 1)

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
      6. Normalize: ratio = pixel / image_dimension
      7. Merge metadata with ratios

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

    # Phase D: Pixel scan
    rects = scan_black_rectangles(img)

    # Phase E: Normalize
    rects = normalize_rects(rects, img_w, img_h)

    # Sort both lists by position for matching
    cluster_meta.sort(key=_sort_key_meta)
    rects.sort(key=lambda r: (r["Top"], r["Left"]))

    # Merge
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
# Preview Pipeline — coordinates + background PNG
# ──────────────────────────────────────────────────────────────────────────────

def generate_preview(xlsx_path: str, output_dir: str, output_id: str) -> dict:
    """Generate coordinates AND render background PNG for upload preview.

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
