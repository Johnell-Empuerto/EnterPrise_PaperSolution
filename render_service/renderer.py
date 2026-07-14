import os
import sys
import tempfile
import shutil
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

from CoordinateEngine.workbook_reader import read_workbook

from render_service.pdf_converter import xlsx_to_pdf
from render_service.background_renderer import pdf_page_to_png, get_page_dimensions
from render_service.models import RenderResponse, FieldModel
from render_service.xml_field_provider import get_cluster_fielddefs
from render_service.page_coordinate_transformer import compute_page_layout, cell_range_to_page_rect
from render_service.validate_alignment import generate_validation_overlay, HAS_CV


def _resolve_xlsx(template_id: int) -> str:
    # 1. Named lookups: prefix search in project root
    TEMPLATE_PREFIXES: dict[int, str] = {
        546: "original.xlsx",
        547: "[V3.1_Sample]",
    }
    if template_id in TEMPLATE_PREFIXES:
        prefix = TEMPLATE_PREFIXES[template_id]
        for entry in os.listdir(_PROJECT_ROOT):
            if entry.startswith(prefix) and entry.endswith(".xlsx"):
                return os.path.join(_PROJECT_ROOT, entry)
    # 2. Named lookup: search known subdirectories
    SUBDIRECTORIES = ["Investigation_546", "LegacyInvestigation_546", "LegacyReconstruction_546", "LegacyReconstruction_456"]
    if template_id in TEMPLATE_PREFIXES:
        prefix = TEMPLATE_PREFIXES[template_id]
        for subdir in SUBDIRECTORIES:
            spath = os.path.join(_PROJECT_ROOT, subdir)
            if os.path.isdir(spath):
                for entry in os.listdir(spath):
                    if entry.startswith(prefix) and entry.endswith(".xlsx"):
                        return os.path.join(spath, entry)
    # 3. Fallback: Forms directory (template_id.xlsx)
    excelapi = os.path.join(_PROJECT_ROOT, "ExcelAPI", "ExcelAPI")
    forms_dir = os.path.join(excelapi, "Forms")
    if os.path.isdir(forms_dir):
        for f in os.listdir(forms_dir):
            if f.endswith(".xlsx"):
                stem = os.path.splitext(f)[0]
                if stem == str(template_id):
                    return os.path.join(forms_dir, f)
    raise FileNotFoundError(f"Template {template_id} not found")


def render(template_id: int | None = None,
           xlsx_path: str | None = None,
           output_dir: str | None = None,
           dpi: int = 300) -> RenderResponse:
    """Render form overlays — original DB-backed pipeline (UNCHANGED)."""
    if xlsx_path is None:
        if template_id is None:
            raise ValueError("Either template_id or xlsx_path is required")
        xlsx_path = _resolve_xlsx(template_id)
    elif template_id is None:
        stem = os.path.splitext(os.path.basename(xlsx_path))[0]
        try:
            template_id = int(stem)
        except ValueError:
            template_id = abs(hash(stem)) % 1000
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"XLSX not found: {xlsx_path}")
    if output_dir is None:
        output_dir = os.path.join(_PROJECT_ROOT, "ExcelAPI", "ExcelAPI", "Preview")
    os.makedirs(output_dir, exist_ok=True)

    print(f"[renderer] Reading workbook: {xlsx_path}")
    wb = read_workbook(xlsx_path, verbose=False)

    print(f"[renderer] Reading clusters from DB for template {template_id}...")
    try:
        fields, raw_clusters = get_cluster_fielddefs(template_id)
    except Exception as e:
        print(f"[renderer] DB query failed, falling back to field_detector: {e}")
        from CoordinateEngine.field_detector import detect as detect_fields
        fields = detect_fields(wb, verbose=False)

    print(f"[renderer] Found {len(fields)} fields from ConMas definition")
    return _fields_to_response(fields, wb, xlsx_path, template_id, output_dir, dpi)


def render_with_fields(fields: list,
                       xlsx_path: str,
                       output_dir: str | None = None,
                       dpi: int = 300,
                       label: str | None = None) -> RenderResponse:
    """Render form overlays — uses externally provided field definitions.

    This is the upload path: fields come from ExcelClusterReader instead
    of the database.  The rendering pipeline is identical.
    """
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"XLSX not found: {xlsx_path}")
    if output_dir is None:
        output_dir = os.path.join(_PROJECT_ROOT, "ExcelAPI", "ExcelAPI", "Preview")
    os.makedirs(output_dir, exist_ok=True)

    stem = os.path.splitext(os.path.basename(xlsx_path))[0]
    template_id = abs(hash(stem)) % 10000

    print(f"[renderer] Reading workbook: {xlsx_path}")
    wb = read_workbook(xlsx_path, verbose=False)

    print(f"[renderer] Using {len(fields)} externally provided field definitions")
    return _fields_to_response(fields, wb, xlsx_path, template_id, output_dir, dpi)


def _fields_to_response(fields: list,
                        wb,
                        xlsx_path: str,
                        template_id: int,
                        output_dir: str,
                        dpi: int) -> RenderResponse:
    """Common field processing: generate PNG, compute pixel positions, return response.

    Used by both render() (DB path) and render_with_fields() (upload path).
    """
    # Compute layout (needed as fallback if ratios are missing or invalid)
    sheet_index = next((i for i, s in enumerate(wb.sheets) if s.visible), 0)
    layout = compute_page_layout(wb, sheet_index=sheet_index, dpi=dpi)

    tmp_dir = tempfile.mkdtemp(prefix="ple_render_")
    page_w_px = page_h_px = 0
    try:
        png_filename = f"runtime_{template_id}_page1.png"
        png_path = os.path.join(output_dir, png_filename)

        print(f"[renderer] Converting to PDF via Excel COM (ExportAsFixedFormat)...")
        pdf_path = xlsx_to_pdf(xlsx_path, output_dir=tmp_dir)
        print(f"[renderer] Rendering PDF to PNG: {png_path}")
        pdf_page_to_png(pdf_path, png_path, dpi=dpi, page_index=0)
        try:
            with Image.open(png_path) as pimg:
                page_w_px, page_h_px = pimg.size
        except Exception:
            page_w_px, page_h_px = get_page_dimensions(pdf_path, dpi=dpi)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    if page_w_px == 0 or page_h_px == 0:
        raise RuntimeError(f"Failed to determine page dimensions for template {template_id}")

    field_models = []
    for field in fields:
        has_ratios = all(v is not None for v in (field.ratio_left, field.ratio_top, field.ratio_right, field.ratio_bottom))
        if has_ratios:
            r_valid = all(0 <= v <= 1 for v in (field.ratio_left, field.ratio_top, field.ratio_right, field.ratio_bottom))
            if r_valid and field.ratio_right > field.ratio_left and field.ratio_bottom > field.ratio_top:
                left_px   = field.ratio_left   * page_w_px
                right_px  = field.ratio_right  * page_w_px
                top_px    = field.ratio_top    * page_h_px
                bottom_px = field.ratio_bottom * page_h_px
                width_px  = right_px  - left_px
                height_px = bottom_px - top_px
            else:
                print(f"[renderer] WARNING: Invalid ratios for {field.addr}, falling back to transformer")
                rect = cell_range_to_page_rect(layout, wb,
                                               field.col, field.row,
                                               field.col_end, field.row_end,
                                               sheet_index=sheet_index)
                left_px = rect["left_px"]
                top_px = rect["top_px"]
                width_px = rect["width_px"]
                height_px = rect["height_px"]
        else:
            rect = cell_range_to_page_rect(layout, wb,
                                           field.col, field.row,
                                           field.col_end, field.row_end,
                                           sheet_index=sheet_index)
            left_px = rect["left_px"]
            top_px = rect["top_px"]
            width_px = rect["width_px"]
            height_px = rect["height_px"]

        field_models.append(FieldModel(
            id=field.addr.replace("$", ""),
            label=field.addr.replace("$", ""),
            left_px=round(left_px),
            top_px=round(top_px),
            width_px=round(width_px),
            height_px=round(height_px),
            type=field.data_type or "text",
            required=field.required,
        ))

    # Generate debug overlay
    debug_png = None
    try:
        debug_path = os.path.join(output_dir, f"runtime_{template_id}_debug.png")
        _generate_debug_overlay(png_path, field_models, debug_path)
        debug_png = f"/preview/runtime_{template_id}_debug.png"
    except Exception as e:
        print(f"[renderer] Debug overlay failed: {e}")

    # Validation overlay — no calibration offsets
    if HAS_CV and fields and all(
        all(v is not None for v in (f.ratio_left, f.ratio_top, f.ratio_right, f.ratio_bottom))
        for f in fields
    ):
        try:
            val_path = os.path.join(output_dir, f"runtime_{template_id}_validation.png")
            generate_validation_overlay(png_path, fields, val_path, page_w_px, page_h_px)
            print(f"[renderer] Validation overlay saved: {val_path}")
        except Exception as e:
            print(f"[renderer] Validation overlay failed: {e}")

    return RenderResponse(
        page_width=page_w_px,
        page_height=page_h_px,
        background_image=f"/preview/{png_filename}",
        debug_image=debug_png,
        fields=field_models,
    )


def _generate_debug_overlay(bg_path: str, fields: list[FieldModel], output_path: str):
    if not HAS_PIL:
        print("[renderer] PIL not available, skipping debug overlay")
        return
    img = Image.open(bg_path).convert("RGBA")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    try:
        font = ImageFont.truetype("arial.ttf", 14)
    except:
        font = ImageFont.load_default()
    for f in fields:
        x, y, w, h = round(f.left_px), round(f.top_px), round(f.width_px), round(f.height_px)
        draw.rectangle([x, y, x + w, y + h], fill=(255, 255, 0, 64), outline=(255, 0, 0), width=2)
        draw.text((x + 2, y + 2), f.id, fill=(255, 0, 0, 255), font=font)
    img = Image.alpha_composite(img, overlay)
    img.convert("RGB").save(output_path, "PNG")
    print(f"[renderer] Debug overlay saved: {output_path}")
