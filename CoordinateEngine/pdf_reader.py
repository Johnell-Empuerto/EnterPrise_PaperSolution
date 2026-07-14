import math
from CoordinateEngine.models.field_def import ContentBounds


def _rect_is_valid(x0, y0, x1, y1):
    """Return True if the rectangle has positive finite area."""
    if any(math.isnan(v) or math.isinf(v) for v in (x0, y0, x1, y1)):
        return False
    if x1 <= x0 or y1 <= y0:
        return False
    return True


def measure_content_bounds(pdf_path):
    import fitz
    doc = fitz.open(pdf_path)
    page = doc[0]
    pw, ph = page.rect.width, page.rect.height
    bounds = []

    for b in page.get_text("blocks"):
        if b[6] == 0 and _rect_is_valid(b[0], b[1], b[2], b[3]):
            bounds.append((b[0], b[1], b[2], b[3]))

    for d in page.get_drawings():
        r = d.get("rect")
        if r and _rect_is_valid(r.x0, r.y0, r.x1, r.y1):
            bounds.append((r.x0, r.y0, r.x1, r.y1))

    for img in page.get_images(full=True):
        try:
            ir = page.get_image_bbox(img)
            if ir and _rect_is_valid(ir.x0, ir.y0, ir.x1, ir.y1):
                bounds.append((ir.x0, ir.y0, ir.x1, ir.y1))
        except Exception:
            pass

    doc.close()

    if not bounds:
        bounds = [(8, 8, pw - 8, ph - 8)]

    x0 = min(b[0] for b in bounds)
    y0 = min(b[1] for b in bounds)
    x1 = max(b[2] for b in bounds)
    y1 = max(b[3] for b in bounds)

    return ContentBounds(
        x0=x0, y0=y0, x1=x1, y1=y1,
        width_pt=x1 - x0, height_pt=y1 - y0,
        page_width_pt=pw, page_height_pt=ph,
        method="pdf"
    )


def detect_centering(content, ml_pt, mt_pt, mr_pt, mb_pt, threshold_pt=3.0):
    pw = content.page_width_pt
    ph = content.page_height_pt
    left_gap = content.x0 - ml_pt
    right_gap = (pw - mr_pt) - content.x1
    top_gap = content.y0 - mt_pt
    bottom_gap = (ph - mb_pt) - content.y1
    centered_h = abs(left_gap - right_gap) < threshold_pt
    centered_v = abs(top_gap - bottom_gap) < threshold_pt
    pw_adj = pw - ml_pt - mr_pt
    ph_adj = ph - mt_pt - mb_pt
    if content.width_pt > pw_adj * 0.95:
        centered_h = False
    if content.height_pt > ph_adj * 0.95:
        centered_v = False
    return centered_h, centered_v


def compute_effective_dims(content, ml_pt, mt_pt, mr_pt, mb_pt, centered_h, centered_v):
    pw = content.page_width_pt
    ph = content.page_height_pt
    pw_adj = pw - ml_pt - mr_pt
    ph_adj = ph - mt_pt - mb_pt
    if centered_h:
        eff_w = pw_adj - 2 * (content.x0 - ml_pt)
    else:
        eff_w = content.width_pt
    if centered_v:
        eff_h = ph_adj - 2 * (content.y0 - mt_pt)
    else:
        eff_h = content.height_pt
    return eff_w, eff_h
