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
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
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
    finally:
        pythoncom.CoUninitialize()


def _identify_clusters_from_fields_sheet(xlsx_path: str):
    """Read cluster ranges from hidden _Fields sheet.

    The _Fields sheet has columns:
      A=Address, B=FieldId, C=FieldName, D=FieldType, E=SheetName, ...

    Returns list of dicts with cellAddr, name, type, input_parameter.
    Returns [] if no _Fields sheet or no data rows.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
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
            wb.Close(False)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
        return clusters
    finally:
        pythoncom.CoUninitialize()


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
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
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

                # Extend bounds to cover all cluster addresses
                # (they may be outside UsedRange if cells are empty)
                import traceback as _tb
                print(f"[SANITIZE] Original UsedRange: lr={lr}, lc={lc}")
                for addr in cluster_addrs:
                    m = CELL_ADDR_RE.match(addr)
                    if m:
                        def _col_idx(s):
                            n = 0
                            for ch in s.replace("$", ""):
                                n = n * 26 + (ord(ch) - 64)
                            return n
                        c1 = _col_idx(m.group(1))
                        c2 = _col_idx(m.group(3)) if m.group(3) else c1
                        r1 = int(m.group(2).replace("$", ""))
                        r2 = int(m.group(4).replace("$", "")) if m.group(4) else r1
                        if c2 > lc or r2 > lr:
                            print(f"[SANITIZE] Extending bounds: cluster addr={addr} needs row={r2} col={c2} (current: lr={lr} lc={lc})")
                        lc = max(lc, c1, c2)
                        lr = max(lr, r1, r2)
                print(f"[SANITIZE] Extended bounds: lr={lr}, lc={lc}")

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
    finally:
        pythoncom.CoUninitialize()


# ──────────────────────────────────────────────────────────────────────────────
# Phase B — Export sanitized workbook to PDF via Excel COM
# ──────────────────────────────────────────────────────────────────────────────

def export_sanitized_pdf(sanitized_path: str) -> str:
    """Export sanitized workbook to PDF via Excel COM ExportAsFixedFormat."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
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

            # Preserve original page setup — do NOT modify Zoom, FitToPages, or margins
            # The sanitized PDF must be geometrically identical to the original workbook PDF
            # so that pixel-scanned coordinates align with the background image.
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

    # Map each group to a pixel rect by vertical overlap (order preserved)
    result = []
    for group in cl_groups:
        if not rects_sorted:
            break
        rect = rects_sorted.pop(0)

        if len(group) == 1:
            result.append(rect)
            continue

        # Multiple clusters in this rect — split proportionally
        # Determine bounding grid of this group
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
        # Exclusive range (Right − Left) — adjacent clusters naturally
        # share the boundary pixel, avoiding gaps in the overlay.
        rw = rr_v - rl
        rh = rb - rt

        for m in group:
            cr = _cluster_col_range(m["cellAddr"])
            rr2 = _cluster_row_range(m["cellAddr"])
            rel_left = (cr[0] - min_col) / grid_cols
            rel_right = (cr[1] - min_col + 1) / grid_cols
            rel_top = (rr2[0] - min_row) / grid_rows
            rel_bot = (rr2[1] - min_row + 1) / grid_rows

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
# Phase X.9 — Single COM Session Pipeline
# Matches original ConMas architecture: open Excel ONCE, do everything.
# ──────────────────────────────────────────────────────────────────────────────

def generate_coordinates_and_preview(
    xlsx_path: str,
    output_dir: str,
    output_id: str,
) -> dict:
    """Generate coordinates AND background PNG in a SINGLE COM session.

    Architecture matches original ConMas:
      1. Open Excel (one Application, one COM lifecycle)
      2. Open original workbook
      3. Export original PDF (background) — BEFORE sanitization
      4. Read comments + sanitize workbook in ONE pass
      5. Save sanitized copy
      6. Reopen sanitized copy
      7. Export sanitized PDF (with IgnorePrintAreas=False, matching ConMas default)
      8. Quit Excel (single cleanup)
      9. Render sanitized PDF → pixel scan → normalize ratios
     10. Render background PNG from original PDF
     11. Return JSON (identical schema to generate_preview)

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
    import pythoncom
    import win32com.client

    # Temporary directories (cleaned up at end)
    tmp_dirs = []

    try:
        # ════════════════════════════════════════════════════
        # PHASE 0: Identify clusters (uses proven old function)
        # Uses its OWN COM session — this is a separate scan
        # that matches ConMas MakeCluster() reading comments
        # via Infragistics or COM per-cell.
        # ════════════════════════════════════════════════════
        cluster_meta = _identify_clusters(xlsx_path)
        if not cluster_meta:
            return {
                "backgroundImage": "",
                "page": {"width": 0, "height": 0},
                "fields": [],
            }

        # ════════════════════════════════════════════════════
        # PHASE A: SINGLE COM SESSION (export PDFs + sanitize)
        # ════════════════════════════════════════════════════
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            try:
                wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

                # ── Step 1: Export original PDF (background) ──
                # MUST be before sanitization because sanitization
                # destroys cell values and fill colors.
                orig_pdf_dir = tempfile.mkdtemp(prefix="ple_orig_pdf_")
                tmp_dirs.append(orig_pdf_dir)
                orig_pdf_path = os.path.join(orig_pdf_dir, "original.pdf")
                wb.ExportAsFixedFormat(0, os.path.abspath(orig_pdf_path))

                # ── Step 2: Sanitize workbook (uses known cluster addresses) ──
                cluster_addrs = set()
                for c in cluster_meta:
                    cluster_addrs.add(c["cellAddr"].upper().replace(" ", ""))
                xlNone = -4142

                for ws in wb.Worksheets:
                    # Clear shapes (interference prevention)
                    try:
                        for shape in list(ws.Shapes):
                            shape.Delete()
                    except Exception:
                        pass

                    # Clear headers/footers
                    ws.PageSetup.CenterHeader = ""
                    ws.PageSetup.CenterFooter = ""

                    used = ws.UsedRange
                    if used is None:
                        continue

                    lr = used.Row + used.Rows.Count - 1
                    lc = used.Column + used.Columns.Count - 1

                    # Sanitize cells: BLACK for cluster, WHITE for non-cluster
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
                                    fill_cell.Interior.Color = 1  # Black
                                except Exception:
                                    pass
                                try:
                                    fill_cell.Value = ""
                                except Exception:
                                    pass
                            else:
                                try:
                                    cell.Interior.Color = 0xFFFFFF  # White
                                except Exception:
                                    pass
                                try:
                                    cell.Value = ""
                                except Exception:
                                    pass

                    # Clear borders across the full range
                    try:
                        clear_range = ws.Range(
                            ws.Cells(1, 1),
                            ws.Cells(lr, lc)
                        )
                        clear_range.Borders.LineStyle = xlNone
                    except Exception:
                        pass

                # ── Step 3: Save sanitized copy ──
                sanitized_dir = tempfile.mkdtemp(prefix="ple_san_")
                tmp_dirs.append(sanitized_dir)
                sanitized_path = os.path.join(sanitized_dir, "sanitized.xlsx")
                wb.SaveAs(os.path.abspath(sanitized_path))
                wb.Close(False)

                # ── Step 4: Reopen sanitized copy & export sanitized PDF ──
                wb_san = excel.Workbooks.Open(os.path.abspath(sanitized_path))

                # Delete metadata sheets to avoid multi-page PDF
                for i in range(wb_san.Sheets.Count, 0, -1):
                    try:
                        name = wb_san.Sheets(i).Name
                    except Exception:
                        name = ""
                    if name in ("_Fields", "_RawData"):
                        wb_san.Sheets(i).Delete()

                pdf_dir = tempfile.mkdtemp(prefix="ple_pdf_")
                tmp_dirs.append(pdf_dir)
                pdf_path = os.path.join(pdf_dir, "sanitized.pdf")

                # Export sanitized PDF
                # IgnorePrintAreas=False matches ConMas default behavior
                ws_active = wb_san.ActiveSheet
                ws_active.ExportAsFixedFormat(
                    0, os.path.abspath(pdf_path),
                    0, 0, False,  # IgnorePrintAreas=False = ConMas default
                )
                wb_san.Close(False)

            finally:
                # Always quit Excel to prevent orphaned processes
                try:
                    excel.Quit()
                except Exception:
                    pass
        finally:
            pythoncom.CoUninitialize()

        # ════════════════════════════════════════════════════
        # PHASE B: NON-COM OPERATIONS (pixel scan, render, normalize)
        # ════════════════════════════════════════════════════

        # ── Step 6: Render sanitized PDF → pixel scan → normalize ──
        os.makedirs(output_dir, exist_ok=True)

        img, img_w, img_h = render_pdf_to_image(pdf_path)
        rects = scan_black_rectangles(img)
        if len(rects) < len(cluster_meta):
            rects = split_merged_rects(rects, cluster_meta)
        rects = normalize_rects(rects, img_w, img_h)

        # Sort by position for matching
        cluster_meta.sort(key=_sort_key_meta)
        rects.sort(key=lambda r: (r["Top"], r["Left"]))

        # Merge metadata with ratios
        fields = []
        for meta, rect in zip(cluster_meta, rects):
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

        # ── Step 7: Render background PNG from original PDF ──
        from render_service.background_renderer import (
            pdf_page_to_png,
            get_page_dimensions,
        )

        png_filename = f"preview_{output_id}.png"
        png_path = os.path.join(output_dir, png_filename)
        pdf_page_to_png(orig_pdf_path, png_path, dpi=DPI)
        page_w, page_h = get_page_dimensions(orig_pdf_path, dpi=DPI)

        return {
            "backgroundImage": png_filename,
            "page": {"width": page_w, "height": page_h},
            "fields": fields,
        }

    finally:
        # Cleanup all temporary directories
        for d in tmp_dirs:
            shutil.rmtree(d, ignore_errors=True)
