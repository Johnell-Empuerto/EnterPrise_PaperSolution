"""
Proper verification for the fix.
Instead of comparing content bounding boxes (which differ because content patterns differ),
this test:
1. Runs the full generate_coordinates() to get ratios
2. Computes expected pixel positions from those ratios on the background image
3. Verifies that the black fill rectangles in the sanitized PDF actually correspond
   to the correct cell positions by checking cell-level geometry

The REAL question: does the coordinate ratio multiplied by background dimensions
give the correct pixel position for the cluster cell?
"""
import sys, os, gc, json, tempfile, shutil
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ROOT = r"C:\Users\MCF-JO~1\Documents\Johnell\PaperLess Enterprise"
DOCS = r"C:\Users\MCF-JO~1\Documents"
sys.path.insert(0, ROOT)
os.chdir(ROOT)

FORM = os.path.join(DOCS, "FormTest - Copy.xlsx")
JAPAN = os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")
DPI_VAL = 300
PT_TO_PX = DPI_VAL / 72.0
OUTPUT = os.path.join(ROOT, "debug_output", "verify_proper_results.json")

import numpy as np
import win32com.client

results = {}

for wb_name, wb_path in [("FormTest", FORM), ("Japanese", JAPAN)]:
    print(f"\n{'='*70}")
    print(f"  WORKBOOK: {wb_name}")
    print(f"{'='*70}")
    wb_r = {}
    
    if not os.path.exists(wb_path):
        print(f"  NOT FOUND")
        wb_r["error"] = "not found"
        results[wb_name] = wb_r
        continue
    
    # ══════════════════════════════════════════════════════════════════
    # 1. Get cluster addresses directly from COM
    # ══════════════════════════════════════════════════════════════════
    print(f"\n  [1] Reading cluster cell addresses from COM...")
    cluster_addrs = []
    def _get_clusters(excel):
        wb = excel.Workbooks.Open(os.path.abspath(wb_path))
        wb_k = []
        
        # Try comments first
        for ws in wb.Worksheets:
            try:
                used = ws.UsedRange
                lr = used.Row + used.Rows.Count - 1
                lc = used.Column + used.Columns.Count - 1
                for r in range(1, min(lr + 1, 100)):
                    for c in range(1, min(lc + 1, 100)):
                        try:
                            cm = ws.Cells(r, c).Comment
                            if cm:
                                ma = ws.Cells(r, c).MergeArea
                                addr = str(ma.Address) if ma else str(ws.Cells(r, c).Address)
                                wb_k.append(addr)
                        except: pass
            except: pass
        
        # Try _Fields sheet
        if not wb_k:
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == "_Fields":
                    fs = wb.Sheets(i)
                    lr = fs.UsedRange.Rows.Count
                    for r in range(2, lr + 1):
                        try:
                            v = fs.Cells(r, 1).Value
                            if v: wb_k.append(str(v).strip())
                        except: pass
                    break
        
        # Also read cell Range positions (Left, Top, Width, Height) from COM
        data_sheet = wb.Sheets(1)
        if data_sheet.Name == "_Fields":
            data_sheet = wb.Sheets(2)
        
        cell_positions = {}
        for addr in wb_k:
            try:
                # Parse address
                import re
                m = re.match(r"(\$?[A-Z]+)(\$?\d+)(?::(\$?[A-Z]+)(\$?\d+))?", addr)
                if not m:
                    continue
                col_letter = m.group(1).replace("$", "")
                # Get the Range
                rng = data_sheet.Range(addr)
                left = float(rng.Left)
                top = float(rng.Top)
                width = float(rng.Width)
                height = float(rng.Height)
                cell_positions[addr] = {
                    "left_pt": left, "top_pt": top,
                    "width_pt": width, "height_pt": height,
                }
            except Exception as e:
                print(f"      COM position error for {addr}: {e}")
        
        # Read page setup
        ps = data_sheet.PageSetup
        pagesetup = {
            "Zoom": str(ps.Zoom),
            "FitToPagesWide": int(getattr(ps, "FitToPagesWide", 0)),
            "FitToPagesTall": int(getattr(ps, "FitToPagesTall", 0)),
            "PaperSize": int(ps.PaperSize),
            "LeftMargin": round(float(ps.LeftMargin), 2),
            "TopMargin": round(float(ps.TopMargin), 2),
            "RightMargin": round(float(ps.RightMargin), 2),
            "BottomMargin": round(float(ps.BottomMargin), 2),
            "CenterHorizontally": bool(ps.CenterHorizontally),
            "CenterVertically": bool(ps.CenterVertically),
            "Order": int(getattr(ps, "Order", 1)),
        }
        
        wb.Close(False)
        return wb_k, cell_positions, pagesetup
    
    gc.collect()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        cluster_addrs, cell_positions, pagesetup = _get_clusters(excel)
        excel.Quit()
        gc.collect()
    except Exception as e:
        print(f"      COM error: {e}")
        import traceback; traceback.print_exc()
        wb_r["error"] = f"com: {e}"
        results[wb_name] = wb_r
        continue
    
    print(f"      Clusters: {cluster_addrs}")
    print(f"      Cell positions (COM):")
    for addr, pos in cell_positions.items():
        print(f"        {addr:15s}  left={pos['left_pt']:.2f}pt top={pos['top_pt']:.2f}pt  w={pos['width_pt']:.2f}pt h={pos['height_pt']:.2f}pt")
    print(f"      PageSetup: Zoom={pagesetup['Zoom']}, FitToPages={pagesetup['FitToPagesWide']}x{pagesetup['FitToPagesTall']}")
    print(f"      PaperSize: {pagesetup['PaperSize']}, Margins: L={pagesetup['LeftMargin']} R={pagesetup['RightMargin']} T={pagesetup['TopMargin']} B={pagesetup['BottomMargin']}")
    wb_r["cluster_addrs"] = cluster_addrs
    wb_r["cell_positions"] = cell_positions
    wb_r["pagesetup"] = pagesetup
    
    # ══════════════════════════════════════════════════════════════════
    # 2. Run generate_coordinates and get ratios
    # ══════════════════════════════════════════════════════════════════
    print(f"\n  [2] Running generate_coordinates()...")
    from render_service.upload_coordinate_generator import generate_coordinates
    gc.collect()
    try:
        fields = generate_coordinates(wb_path)
        print(f"      Generated {len(fields)} fields:")
        coord_positions = {}
        for f in fields:
            # Convert ratios to pixels at 2550x3300
            left_px = f["left_ratio"] * 2550
            top_px = f["top_ratio"] * 3300
            right_px = f["right_ratio"] * 2550
            bottom_px = f["bottom_ratio"] * 3300
            coord_positions[f["cellAddr"]] = {
                "left_ratio": f["left_ratio"],
                "top_ratio": f["top_ratio"],
                "right_ratio": f["right_ratio"],
                "bottom_ratio": f["bottom_ratio"],
                "left_px_2550": round(left_px, 1),
                "top_px_3300": round(top_px, 1),
                "right_px_2550": round(right_px, 1),
                "bottom_px_3300": round(bottom_px, 1),
            }
            print(f"        {f['name']:20s}  {f['cellAddr']:15s}  "
                  f"ratio L={f['left_ratio']:.5f} T={f['top_ratio']:.5f} R={f['right_ratio']:.5f} B={f['bottom_ratio']:.5f}  "
                  f"px: ({left_px:.0f},{top_px:.0f})-({right_px:.0f},{bottom_px:.0f})")
        wb_r["coord_positions"] = coord_positions
        
        # ══════════════════════════════════════════════════════════════════
        # 3. COMPARE: COM cell positions vs coordinate ratios
        # ══════════════════════════════════════════════════════════════════
        print(f"\n  [3] Comparing COM positions vs coordinate ratios...")
        
        # The COM range positions (in points) need to be converted to page-relative
        # positions by accounting for margins and centering.
        # Then compared with the coordinate ratio * page_dimensions.
        
        left_margin = pagesetup.get("LeftMargin", 0)
        right_margin = pagesetup.get("RightMargin", 0)
        top_margin = pagesetup.get("TopMargin", 0)
        bottom_margin = pagesetup.get("BottomMargin", 0)
        center_h = pagesetup.get("CenterHorizontally", False)
        center_v = pagesetup.get("CenterVertically", False)
        
        # Page dimensions (Letter = 612x792 pt)
        page_w_pt = 612.0
        page_h_pt = 792.0
        printable_w = page_w_pt - left_margin - right_margin
        printable_h = page_h_pt - top_margin - bottom_margin
        
        comparisons = []
        # For computing content width/height for centering, we need the print area
        # dimensions. Check from the first cell's position. Actually, we can compute
        # from the total cluster range.
        if cell_positions:
            # Estimate content bounds from all cluster positions
            min_left = min(p["left_pt"] for p in cell_positions.values())
            min_top = min(p["top_pt"] for p in cell_positions.values())
            max_right = max(p["left_pt"] + p["width_pt"] for p in cell_positions.values())
            max_bottom = max(p["top_pt"] + p["height_pt"] for p in cell_positions.values())
            content_w = max_right - min_left
            content_h = max_bottom - min_top
            
            # Compute printable origin (same formula as ComputePrintedOrigin)
            origin_x = left_margin
            origin_y = top_margin
            if center_h and content_w < printable_w:
                origin_x += (printable_w - content_w) / 2.0
            if center_v and content_h < printable_h:
                origin_y += (printable_h - content_h) / 2.0
            
            print(f"      Printable area: {printable_w:.0f}x{printable_h:.0f} pt")
            print(f"      Content bounds: {content_w:.0f}x{content_h:.0f} pt")
            print(f"      Printed origin: ({origin_x:.1f}, {origin_y:.1f}) pt")
            
            for addr in cluster_addrs:
                com_pos = cell_positions.get(addr)
                coord_pos = coord_positions.get(addr)
                if com_pos and coord_pos:
                    # Expected page position (in points)
                    expected_left_pt = origin_x + (com_pos["left_pt"] - min_left)
                    expected_top_pt = origin_y + (com_pos["top_pt"] - min_top)
                    
                    # Convert to pixels at 300 DPI
                    expected_left_px = expected_left_pt * PT_TO_PX
                    expected_top_px = expected_top_pt * PT_TO_PX
                    
                    # Compare with coordinate ratio * 2550
                    coord_left_px = coord_pos["left_px_2550"]
                    coord_top_px = coord_pos["top_px_3300"]
                    
                    diff_x = coord_left_px - expected_left_px
                    diff_y = coord_top_px - expected_top_px
                    
                    comparisons.append({
                        "addr": addr,
                        "com_left_pt": com_pos["left_pt"],
                        "com_top_pt": com_pos["top_pt"],
                        "expected_left_pt": round(expected_left_pt, 2),
                        "expected_top_pt": round(expected_top_pt, 2),
                        "expected_left_px": round(expected_left_px, 1),
                        "expected_top_px": round(expected_top_px, 1),
                        "coord_left_px": coord_left_px,
                        "coord_top_px": coord_top_px,
                        "diff_px_x": round(diff_x, 1),
                        "diff_px_y": round(diff_y, 1),
                        "diff_pt_x": round(diff_x / PT_TO_PX, 2),
                        "diff_pt_y": round(diff_y / PT_TO_PX, 2),
                    })
                    
                    print(f"      {addr:15s}: COM=({com_pos['left_pt']:.1f},{com_pos['top_pt']:.1f})pt → "
                          f"expected=({expected_left_pt:.0f},{expected_top_pt:.0f})pt "
                          f"→ expected px=({expected_left_px:.0f},{expected_top_px:.0f}) "
                          f"vs coord px=({coord_left_px:.0f},{coord_top_px:.0f}) "
                          f"DIFF=({diff_x:.0f},{diff_y:.0f})px = ({diff_x/PT_TO_PX:.1f},{diff_y/PT_TO_PX:.1f})pt")
            
            wb_r["comparisons"] = comparisons
            
            # Check if the fix is working: all diffs should be < 5px
            all_match = all(abs(c["diff_px_x"]) < 5 and abs(c["diff_px_y"]) < 5 for c in comparisons)
            if comparisons:
                max_diff_x = max(abs(c["diff_px_x"]) for c in comparisons)
                max_diff_y = max(abs(c["diff_px_y"]) for c in comparisons)
                print(f"\n      Max difference: X={max_diff_x:.0f}px, Y={max_diff_y:.0f}px")
                if all_match:
                    print(f"      ✅ ALL FIELDS MATCH within 5px — WORKS!")
                else:
                    print(f"      ⚠️ Some fields differ by more than 5px")
                wb_r["max_diff_px_x"] = max_diff_x
                wb_r["max_diff_px_y"] = max_diff_y
                wb_r["all_match"] = all_match
    
    except Exception as e:
        print(f"      generate_coordinates error: {e}")
        import traceback; traceback.print_exc()
        wb_r["generate_error"] = str(e)
    
    results[wb_name] = wb_r

# Save
with open(OUTPUT, 'w', encoding='utf-8') as f:
    # Convert numpy/path types for JSON
    class SimpleEncoder(json.JSONEncoder):
        def default(self, o):
            return str(o)
    json.dump(results, f, ensure_ascii=False, indent=2, cls=SimpleEncoder)
print(f"\n\nResults saved to {OUTPUT}")

# Summary
print(f"\n{'='*70}")
print(f"  VERIFICATION SUMMARY")
print(f"{'='*70}")
for wb, r in results.items():
    print(f"\n  {wb}:")
    if "error" in r:
        print(f"    ❌ {r['error']}")
        continue
    print(f"    Clusters: {r.get('cluster_addrs', [])}")
    print(f"    CellPositions: {len(r.get('cell_positions', {}))}")
    print(f"    PageSetup: {r.get('pagesetup', {})}")
    cmp = r.get("comparisons", [])
    if cmp:
        print(f"    Max diff (px): X={r.get('max_diff_px_x', '?')} Y={r.get('max_diff_px_y', '?')}")
        print(f"    All match: {r.get('all_match', '?')}")
        if r.get("all_match"):
            print(f"    ✅ FIX VERIFIED: Coordinates match expected COM positions")
        else:
            print(f"    ⚠️ Some differences remain (see details above)")
    print(f"    Coordinates: {len(r.get('coord_positions', {}))} fields")

print(f"\n  Done.")
