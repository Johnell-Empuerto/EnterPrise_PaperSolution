"""
Coordinate Pipeline Trace — Investigation Only
Trace a single field through the entire rendering pipeline to find where the 1-3px offset originates.
"""
import sys, os, tempfile, shutil, json, urllib.request, math
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from CoordinateEngine.workbook_reader import read_workbook
from render_service.page_coordinate_transformer import (
    compute_page_layout, cell_range_to_page_rect,
    _sum_cols, _sum_rows, _col_pt, _row_pt,
)
from render_service.xml_field_provider import get_cluster_fielddefs
from render_service.pdf_converter import xlsx_to_pdf
from PIL import Image

DPI = 300
PT_TO_PX = DPI / 72.0

def trace_template(template_id: int, xlsx_path: str, label: str):
    print("=" * 100)
    print(f"COORDINATE PIPELINE TRACE -- {label}")
    print("=" * 100)
    print()

    # --- STAGE 0: DB Ratios ---
    print("--- STAGE 0: XML/Database ratios ---")
    fields, raw_clusters = get_cluster_fielddefs(template_id)
    f = fields[0]
    c = raw_clusters[0]
    print(f"  Field:          {f.addr}")
    print(f"  DB raw values:  left={c['left_position']!r}  top={c['top_position']!r}  right={c['right_position']!r}  bottom={c['bottom_position']!r}")
    print(f"  Ratios:         left={f.ratio_left:.10f}  top={f.ratio_top:.10f}  right={f.ratio_right:.10f}  bottom={f.ratio_bottom:.10f}")
    rw = f.ratio_right - f.ratio_left
    rh = f.ratio_bottom - f.ratio_top
    print(f"  Ratio width:    {rw:.10f}")
    print(f"  Ratio height:   {rh:.10f}")
    print(f"  Running Δ:      0 px")
    print()

    # --- STAGE 1: Workbook ---
    print("--- STAGE 1: Workbook (XLSX parse) ---")
    wb = read_workbook(xlsx_path, verbose=False)
    sheet = next((s for s in wb.sheets if s.visible), wb.sheets[0])
    si = wb.sheets.index(sheet)
    print(f"  Visible sheet:      {sheet.name} (index {si})")
    print(f"  Workbook page_pt:   {wb.page_width_pt} x {wb.page_height_pt}")
    print(f"  Sheet page_pt:      {sheet.page_width_pt} x {sheet.page_height_pt}")
    print(f"  Margins (in):       L={sheet.margin_left_pt}  R={sheet.margin_right_pt}  T={sheet.margin_top_pt}  B={sheet.margin_bottom_pt}")
    # Convert margin inches → pt
    ml_pt = sheet.margin_left_pt * 72.0
    mr_pt = sheet.margin_right_pt * 72.0
    mt_pt = sheet.margin_top_pt * 72.0
    mb_pt = sheet.margin_bottom_pt * 72.0
    print(f"  Margins (pt):       L={ml_pt:.2f}  R={mr_pt:.2f}  T={mt_pt:.2f}  B={mb_pt:.2f}")
    print(f"  Center H: {sheet.center_horizontally}  Center V: {sheet.center_vertically}")
    print(f"  Fit: {sheet.fit_to_pages_wide}W x {sheet.fit_to_pages_tall}H  Zoom: {sheet.zoom}")
    print(f"  Paper size: {sheet.paper_size}  Orientation: {sheet.orientation}")
    print(f"  Print area: {wb.print_area_addr}")

    # Print column widths for relevant columns
    for ci in range(1, 14):
        w = _col_pt(wb, ci, si)
        if ci >= 6:
            print(f"    Col {ci}: {w:.4f} pt" + (f"  ← field col {f.col}–{f.col_end}" if f.col <= ci <= f.col_end else ""))
    for ri in range(1, 12):
        h = _row_pt(wb, ri, si)
        if ri >= 4:
            print(f"    Row {ri}: {h:.4f} pt" + (f"  ← field row {f.row}–{f.row_end}" if f.row <= ri <= f.row_end else ""))

    # Compute cell range position in workbook space (no scaling, no origin)
    range_left = _sum_cols(wb, 1, f.col - 1, si) if f.col > 1 else 0.0
    range_top = _sum_rows(wb, 1, f.row - 1, si) if f.row > 1 else 0.0
    range_w = _sum_cols(wb, f.col, f.col_end, si)
    range_h = _sum_rows(wb, f.row, f.row_end, si)
    print(f"  Cell range (raw workbook pt): left={range_left:.4f}  top={range_top:.4f}  w={range_w:.4f}  h={range_h:.4f}")
    print(f"  Running Δ:      0 px")
    print()

    # --- STAGE 2: Page Layout ---
    print("--- STAGE 2: Page Layout ---")
    layout = compute_page_layout(wb, sheet_index=si, dpi=DPI)
    print(f"  origin_pt:          ({layout.origin_x_pt:.6f}, {layout.origin_y_pt:.6f})")
    print(f"  scale:              ({layout.scale_w:.6f}, {layout.scale_h:.6f})")
    print(f"  content_pt:         {layout.content_width_pt:.6f} x {layout.content_height_pt:.6f}")
    print(f"  printable_pt:       {layout.printable_width_pt:.6f} x {layout.printable_height_pt:.6f}")
    print(f"  page_pt (workbook): {layout.page_width_pt:.6f} x {layout.page_height_pt:.6f}")

    # What the transformer predicts for this field
    rect = cell_range_to_page_rect(layout, wb, f.col, f.row, f.col_end, f.row_end, sheet_index=si)
    print(f"  Transformer px:    L={rect['left_px']}  T={rect['top_px']}  W={rect['width_px']}  H={rect['height_px']}")
    print(f"  Transformer pt:    L={rect['page_left_pt']:.4f}  T={rect['page_top_pt']:.4f}")
    print(f"  Running Δ:          N/A (transformer uses different coordinate space)")
    print()

    # --- STAGE 3: LibreOffice PDF Export ---
    print("--- STAGE 3: LibreOffice -> PDF Export ---")
    tmp_dir = tempfile.mkdtemp(prefix="ple_trace_")
    try:
        pdf_path = xlsx_to_pdf(xlsx_path, output_dir=tmp_dir)
        import fitz
        doc = fitz.open(pdf_path)
        page = doc[0]

        pdf_w_pt = page.rect.width
        pdf_h_pt = page.rect.height
        print(f"  PDF page.rect:      {page.rect.x0:.4f},{page.rect.y0:.4f} → {page.rect.x1:.4f},{page.rect.y1:.4f}  = {pdf_w_pt:.4f} x {pdf_h_pt:.4f} pt")
        print(f"  PDF MediaBox:       ({page.mediabox.x0:.4f},{page.mediabox.y0:.4f}) → ({page.mediabox.x1:.4f},{page.mediabox.y1:.4f})")
        print(f"  PDF CropBox:        ({page.cropbox.x0:.4f},{page.cropbox.y0:.4f}) → ({page.cropbox.x1:.4f},{page.cropbox.y1:.4f})")

        for bn in ["bleedbox", "trimbox", "artbox"]:
            try:
                b = getattr(page, bn)()
                print(f"  PDF {bn}: ({b.x0:.4f},{b.y0:.4f}) → ({b.x1:.4f},{b.y1:.4f})")
            except Exception:
                print(f"  PDF {bn}:          NOT SET")

        print(f"  PDF rotation:       {page.rotation}")

        # Compare PDF to workbook
        print(f"  PDF vs Workbook:    {pdf_w_pt - layout.page_width_pt:+.4f} x {pdf_h_pt - layout.page_height_pt:+.4f} pt")
        
        # The PDF page is the truth. The ratios should be relative to the PDF page.
        # Expected position = ratio * PDF dimensions
        pdf_left_pt = f.ratio_left * pdf_w_pt
        pdf_top_pt = f.ratio_top * pdf_h_pt
        pdf_right_pt = f.ratio_right * pdf_w_pt
        pdf_bottom_pt = f.ratio_bottom * pdf_h_pt
        print(f"  Ratio * PDF pt:     L={pdf_left_pt:.6f}  T={pdf_top_pt:.6f}  R={pdf_right_pt:.6f}  B={pdf_bottom_pt:.6f}")
        print(f"  Running Δ:          0 px (ratios match PDF coordinate space)")
        print()

        # --- STAGE 4: PyMuPDF Rasterization ---
        print("--- STAGE 4: PyMuPDF -> PNG Rasterization ---")
        zoom = DPI / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        png_w = pix.width
        png_h = pix.height

        print(f"  Zoom factor:        {zoom:.6f}")
        print(f"  PDF pt:             {pdf_w_pt:.10f} x {pdf_h_pt:.10f}")
        print(f"  Exact px (pt × zoom): {pdf_w_pt * zoom:.10f} x {pdf_h_pt * zoom:.10f}")
        print(f"  PyMuPDF pixmap:     {png_w} x {png_h}")

        # How does PyMuPDF round? Check against different rounding modes
        exact_w = pdf_w_pt * zoom
        exact_h = pdf_h_pt * zoom
        print(f"  floor(px):          {math.floor(exact_w)} x {math.floor(exact_h)}")
        print(f"  ceil(px):           {math.ceil(exact_w)} x {math.ceil(exact_h)}")
        print(f"  round(px):          {round(exact_w)} x {round(exact_h)}")
        print(f"  int(px):            {int(exact_w)} x {int(exact_h)}")
        # (already checked above)

        # The rounding error at the PNG level
        w_error = png_w - exact_w
        h_error = png_h - exact_h
        print(f"  Pixmap rounding error: {w_error:+.10f} x {h_error:+.10f} px")

        # The key question: does this rounding introduce the offset?
        # ratio * exact_px vs ratio * rounded_px
        px_left_exact = f.ratio_left * exact_w
        px_left_rounded = f.ratio_left * png_w
        round_effect_l = px_left_rounded - px_left_exact
        px_top_exact = f.ratio_top * exact_h
        px_top_rounded = f.ratio_top * png_h
        round_effect_t = px_top_rounded - px_top_exact
        print(f"  Effect of pixmap rounding on field position:")
        print(f"    Left: ratio*exact={px_left_exact:.6f}  ratio*rounded={px_left_rounded:.6f}  diff={round_effect_l:+.6f}")
        print(f"    Top:  ratio*exact={px_top_exact:.6f}  ratio*rounded={px_top_rounded:.6f}  diff={round_effect_t:+.6f}")
        print(f"  Running Δ:          L={round_effect_l:+.6f}  T={round_effect_t:+.6f}")
        print()

        # --- STAGE 5: Ratio-based position prediction ---
        print("--- STAGE 5: Renderer (ratio-based position) ---")
        # Current renderer: left_px = ratio * png_w + calibration
        raw_l = f.ratio_left * png_w
        raw_t = f.ratio_top * png_h
        raw_r = f.ratio_right * png_w
        raw_b = f.ratio_bottom * png_h

        # After applying current calibration
        cal_l = raw_l + 0.5 * PT_TO_PX - 0.5 * PT_TO_PX / 2.0
        cal_t = raw_t + 1.0 * PT_TO_PX - 0.5 * PT_TO_PX / 2.0
        cal_w = (raw_r - raw_l) + 0.5 * PT_TO_PX
        cal_h = (raw_b - raw_t) + 0.5 * PT_TO_PX

        print(f"  Raw (ratio * PNG): L={raw_l:.4f}  T={raw_t:.4f}  W={raw_r-raw_l:.4f}  H={raw_b-raw_t:.4f}")
        print(f"  With calibration:  L={cal_l:.4f}  T={cal_t:.4f}  W={cal_w:.4f}  H={cal_h:.4f}")

        # Get actual API output
        req = urllib.request.Request(
            "http://localhost:5091/render/runtime",
            data=b'{"template_id": ' + str(template_id).encode() + b'}',
            headers={"Content-Type": "application/json"},
        )
        resp = json.loads(urllib.request.urlopen(req).read().decode())
        rf = resp["fields"][0]
        actual_l = rf["left_px"]
        actual_t = rf["top_px"]
        actual_w = rf["width_px"]
        actual_h = rf["height_px"]
        print(f"  Actual renderer:    L={actual_l}  T={actual_t}  W={actual_w}  H={actual_h}")
        print(f"  API page:           {resp['page_width']} x {resp['page_height']}")
        print()

        # --- CUMULATIVE DIFFERENCE TABLE ---
        print("=" * 100)
        print("CUMULATIVE DIFFERENCE ANALYSIS")
        print("=" * 100)
        print()
        print(f"  {'Stage':<40} {'Left D':>12} {'Top D':>12} {'Width D':>12} {'Height D':>12}")
        print(f"  {'-'*40} {'-'*12} {'-'*12} {'-'*12} {'-'*12}")

        # Track cumulative diffs
        cum_l, cum_t, cum_w, cum_h = 0.0, 0.0, 0.0, 0.0

        # Stage 0→1: DB ratio stored as float - exact representation
        print(f"  {'STAGE 0→1: DB ratio capture':<40} {cum_l:>12.4f} {cum_t:>12.4f} {cum_w:>12.4f} {cum_h:>12.4f}")

        # Stage 1→2: Workbook to ratio conceptual space — ratios are relative to page
        # No difference introduced here
        print(f"  {'STAGE 1→2: Workbook → ratio concept':<40} {cum_l:>12.4f} {cum_t:>12.4f} {cum_w:>12.4f} {cum_h:>12.4f}")

        # Stage 2→3: PDF export — check if PDF dimensions differ from workbook
        if pdf_w_pt != layout.page_width_pt or pdf_h_pt != layout.page_height_pt:
            ratio_w_diff = f.ratio_left * (pdf_w_pt - layout.page_width_pt) * PT_TO_PX
            ratio_h_diff = f.ratio_top * (pdf_h_pt - layout.page_height_pt) * PT_TO_PX
            cum_l += ratio_w_diff
            cum_t += ratio_h_diff
            print(f"  {'STAGE 2→3: PDF page size mismatch':<40} {ratio_w_diff:>+11.4f} {ratio_h_diff:>+11.4f} {'':>12} {'':>12}")
            print(f"  {'  (cumulative)':<40} {cum_l:>12.4f} {cum_t:>12.4f} {cum_w:>12.4f} {cum_h:>12.4f}")
        else:
            print(f"  {'STAGE 2→3: PDF page = workbook page':<40} {cum_l:>12.4f} {cum_t:>12.4f} {cum_w:>12.4f} {cum_h:>12.4f}")

        # Stage 3→4: PyMuPDF rasterization — int rounding of pixmap dimensions
        # This is the key check: does pixmap dimension rounding introduce the offset?
        px_round_w = round_effect_l  # already computed
        px_round_t = round_effect_t
        cum_l += px_round_w
        cum_t += px_round_t

        # Width/height rounding effect
        raw_w_exact = (f.ratio_right * exact_w) - (f.ratio_left * exact_w)
        raw_w_rounded = (f.ratio_right * png_w) - (f.ratio_left * png_w)
        px_round_w_w = raw_w_rounded - raw_w_exact
        raw_h_exact = (f.ratio_bottom * exact_h) - (f.ratio_top * exact_h)
        raw_h_rounded = (f.ratio_bottom * png_h) - (f.ratio_top * png_h)
        px_round_h_h = raw_h_rounded - raw_h_exact
        cum_w += px_round_w_w
        cum_h += px_round_h_h

        print(f"  {'STAGE 3→4: PyMuPDF pixmap rounding':<40} {px_round_w:>+11.4f} {px_round_t:>+11.4f} {px_round_w_w:>+11.4f} {px_round_h_h:>+11.4f}")
        print(f"  {'  (cumulative)':<40} {cum_l:>12.4f} {cum_t:>12.4f} {cum_w:>12.4f} {cum_h:>12.4f}")

        # Stage 4→5: Renderer rounding (round() on final coordinates)
        # The renderer applies round(left_px) etc.
        render_round_l = round(actual_l) - actual_l if actual_l != round(actual_l) else 0
        render_round_t = round(actual_t) - actual_t if actual_t != round(actual_t) else 0
        
        # The actual residual after all known transformations
        # This is what the calibration is compensating for
        residual_l = actual_l - raw_l - cum_l
        residual_t = actual_t - raw_t - cum_t
        residual_w = actual_w - (raw_r - raw_l) - cum_w
        residual_h = actual_h - (raw_b - raw_t) - cum_h

        cum_l += residual_l
        cum_t += residual_t
        cum_w += residual_w
        cum_h += residual_h

        print(f"  {'STAGE 4→5: Renderer output round()':<40} {'':>12} {'':>12} {'':>12} {'':>12}")
        print(f"  {'  (cumulative before residual)':<40} {cum_l - residual_l:>12.4f} {cum_t - residual_t:>12.4f} {cum_w - residual_w:>12.4f} {cum_h - residual_h:>12.4f}")
        print(f"  {'STAGE 5: UNEXPLAINED RESIDUAL':<40} {residual_l:>+11.4f} {residual_t:>+11.4f} {residual_w:>+11.4f} {residual_h:>+11.4f}")
        print(f"  {'  (FINAL cumulative)':<40} {cum_l:>12.4f} {cum_t:>12.4f} {cum_w:>12.4f} {cum_h:>12.4f}")
        print(f"  {'  (FINAL in pt @ 300dpi)':<40} {cum_l/PT_TO_PX:>12.4f} {cum_t/PT_TO_PX:>12.4f} {cum_w/PT_TO_PX:>12.4f} {cum_h/PT_TO_PX:>12.4f}")
        print()

        print(f"  === RESULT ===")
        print(f"  Pixmap rounding accounts for:   L={px_round_w:+.4f} T={px_round_t:+.4f}")
        print(f"  Unexplained residual:           L={residual_l:+.4f} T={residual_t:+.4f}")
        print(f"  Residual in pt:                 L={residual_l/PT_TO_PX:+.4f} T={residual_t/PT_TO_PX:+.4f}")
        print(f"  Current calibration:             L=0.5pt T=1.0pt EW=0.5pt EH=0.5pt")
        print()

        # --- PDF BOX OFFSET CHECK ---
        print("--- PDF BOX OFFSET CHECK ---")
        mb = page.mediabox
        cb = page.cropbox
        print(f"  MediaBox origin:    ({mb.x0}, {mb.y0})")
        print(f"  CropBox origin:     ({cb.x0}, {cb.y0})")
        if mb.x0 != 0 or mb.y0 != 0:
            print(f"  ⚠  Non-zero MediaBox origin! {mb.x0:.4f}, {mb.y0:.4f}")
        if cb.x0 != 0 or cb.y0 != 0:
            print(f"  ⚠  Non-zero CropBox origin! {cb.x0:.4f}, {cb.y0:.4f}")
        
        # Content bounding box — skip (get_image_bbox requires full page image list)
        print()

        # --- HALF-PIXEL ORIGIN CHECK ---
        print("--- HALF-PIXEL ORIGIN CHECK ---")
        for ox_name, ox, oy in [("(0, 0)", 0, 0)]:
            adj_l = f.ratio_left * png_w + ox
            adj_t = f.ratio_top * png_h + oy
            dl_l = actual_l - adj_l
            dl_t = actual_t - adj_t
            print(f"  Origin {ox_name}: ΔL={dl_l:+.4f}  ΔT={dl_t:+.4f}  ΔW={actual_w - (raw_r - raw_l):+.4f}  ΔH={actual_h - (raw_b - raw_t):+.4f}")

        # Check: does PyMuPDF use a different rounding mode for pixmap?
        # Fitzy uses int() internally: pix.width = int(page.rect.width * zoom + 0.5) OR int(page.rect.width * zoom)?
        # Let's check by seeing what value produces the observed pix width
        for rounding_mode_name, rounding_fn in [
            ("int()", lambda x: int(x)),
            ("round()", lambda x: round(x)),
            ("floor()", lambda x: math.floor(x)),
            ("ceil()", lambda x: math.ceil(x)),
            ("trunc()", lambda x: int(math.trunc(x))),
            ("int(x+0.5)", lambda x: int(x + 0.5)),
        ]:
            result = rounding_fn(exact_w)
            if result == png_w:
                print(f"  PyMuPDF rounding mode: {rounding_mode_name} gives {result} = png width ✓")
                break
        else:
            print(f"  PyMuPDF rounding mode: UNKNOWN (none of the common modes produce {png_w} from {exact_w})")

        for rounding_mode_name, rounding_fn in [
            ("int()", lambda x: int(x)),
            ("round()", lambda x: round(x)),
            ("floor()", lambda x: math.floor(x)),
            ("ceil()", lambda x: math.ceil(x)),
            ("trunc()", lambda x: int(math.trunc(x))),
            ("int(x+0.5)", lambda x: int(x + 0.5)),
        ]:
            result = rounding_fn(exact_h)
            if result == png_h:
                print(f"  PyMuPDF rounding mode: {rounding_mode_name} gives {result} = png height ✓")
                break
        else:
            print(f"  PyMuPDF rounding mode: UNKNOWN (none of the common modes produce {png_h} from {exact_h})")
        print()

        doc.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    print()


if __name__ == "__main__":
    trace_template(547, "[V3.1_Sample]アンケート用紙.xlsx", "Template 547 — Field I6:M6")
    trace_template(546, "Investigation_546/original.xlsx", "Template 546 — Field A1:B2")
